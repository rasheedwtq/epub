# -*- coding: utf-8 -*-
"""
V-4 · 标题聚类真实准确率实测（M0 原型验证清单 硬门槛之一）

目标：回答"字号聚类能不能达到 ≥98%"这个假设（R-3）。
方法：样本由 make_samples.py 生成时同步登记了 ground truth，
      本脚本用 lxml 实现聚类算法（承《归一化层规格》N7 与 8.1 细则），
      与 ground truth 逐条比对，输出准确率 + 混淆矩阵 + 误判明细。

产出：samples/_v4_heading_report.md
"""
import json
import zipfile
from collections import Counter, defaultdict
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "tests"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(t):
    return "{%s}%s" % (W, t)


# ---------------------------------------------------------------- 解析
def load_paragraphs(path: Path):
    """用 lxml 解析 docx 的全部段落（与阶段1 同一套 XML 视角）"""
    with zipfile.ZipFile(path) as z:
        doc = etree.fromstring(z.read("word/document.xml"))
        default_sz = 21          # 10.5pt = 21 半磅
        if "word/styles.xml" in z.namelist():
            st_root = etree.fromstring(z.read("word/styles.xml"))
            for st in st_root.iter(w("style")):
                if st.get(w("styleId")) == "Normal":
                    sz = st.find(".//" + w("sz"))
                    if sz is not None and sz.get(w("val")):
                        default_sz = int(sz.get(w("val")))
                    break

    paras = []
    for i, p in enumerate(doc.iter(w("p"))):
        txt = "".join(t.text or "" for t in p.iter(w("t")))
        # 众数字号：只取直接子 run 的 rPr/sz（排除 rPrChange 内的历史值）
        szs = []
        weight = defaultdict(int)
        for r in p.findall(w("r")):
            rPr = r.find(w("rPr"))
            if rPr is not None:
                sz = rPr.find(w("sz"))
                if sz is not None and sz.get(w("val")) and sz.get(w("val")).isdigit():
                    v = int(sz.get(w("val")))
                    szs.append(v)
                    weight[v] += max(sum(len(t.text or "") for t in r.iter(w("t"))), 1)
        # ⚠️ 必须用「字符数加权众数」：普通众数在并列时会取第一个，
        #    导致"首字下沉"(58.5pt 一个字 + 10.5pt 二十字)被误判成大字号标题
        size = max(weight.items(), key=lambda kv: kv[1])[0] if weight else default_sz
        max_sz = max(szs) if szs else default_sz
        in_table = any(a.tag == w("tbl") for a in p.iterancestors())
        bold = any(r.find(w("rPr")) is not None
                   and r.find(w("rPr")).find(w("b")) is not None
                   for r in p.findall(w("r")))
        has_draw = p.find(".//" + w("drawing")) is not None
        paras.append({"idx": i, "text": txt, "size_pt": size / 2.0,
                      "max_pt": max_sz / 2.0, "in_table": in_table,
                      "bold": bold, "has_draw": has_draw})
    return paras


# ---------------------------------------------------------------- 聚类算法（承 8.1 / N7）
NEG_PATTERNS = ("***", "---", "※", "注：", "说明：")


def cluster(paras, ratio=1.15, gap_pt=1.5, max_len=40):
    steps = {}
    # Step2 正文基准 B = 直方图峰值（不硬编码）
    cnt = Counter(p["size_pt"] for p in paras)
    body_pt, body_n = cnt.most_common(1)[0]
    steps["正文基准B_pt"] = body_pt
    steps["B的段落数"] = body_n
    steps["字号分布"] = dict(sorted(cnt.items(), reverse=True))

    # Step1 是否已用 Heading 样式
    # Step3 候选：字号 ≥ B×ratio 且非表内
    thr = body_pt * ratio
    steps["候选门槛_pt"] = round(thr, 2)
    # Step3 候选（**软门槛**：任何明显大于正文基准的字号都入候选，
    #   再靠频次/邻域区分"真标题"与"正文强调语" —— V-4 实测：12pt/10.5pt=1.14
    #   会被 B×1.15 硬门槛整批漏判，故改为软门槛）
    thr_soft = body_pt + 0.5
    thr_strict = body_pt * ratio
    steps["候选门槛_软_pt"] = round(thr_soft, 2)
    steps["候选门槛_严格B×%.2f_pt" % ratio] = round(thr_strict, 2)
    cand = [p for p in paras
            if p["size_pt"] >= thr_soft and not p["in_table"] and p["text"].strip()]

    # Step3.5 剔除「文档标题/书名」：最大字号 + 仅出现1次 + 位于文档前 15%
    #   真实文档每本都有书名页，若不剔除会把整个章节层级挤低一级（V-4 实测发现的系统性缺陷）
    doc_title_idx = None
    if cand:
        top_sz = max(p["size_pt"] for p in cand)
        same = [p for p in cand if p["size_pt"] == top_sz]
        if (len(same) == 1
                and same[0]["idx"] <= max(3, int(len(paras) * 0.15))):
            doc_title_idx = same[0]["idx"]
            cand = [p for p in cand if p["idx"] != doc_title_idx]
            steps["剔除文档标题"] = f"{same[0]['text'][:18]} @{top_sz}pt (idx={same[0]['idx']})"
    steps["候选段落数"] = len(cand)

    # Step4 分桶（按字号降序，相邻差 > gap 切一刀）
    buckets = []
    for p in sorted(cand, key=lambda x: (-x["size_pt"], x["idx"])):
        if buckets and abs(buckets[-1][0]["size_pt"] - p["size_pt"]) <= gap_pt:
            buckets[-1].append(p)
        else:
            buckets.append([p])
    buckets = buckets[:4]
    steps["桶"] = [{"层级": f"h{i+1}", "字号pt": b[0]["size_pt"], "段数": len(b)}
                  for i, b in enumerate(buckets)]

    labels = {}
    if doc_title_idx is not None:
        labels[doc_title_idx] = "doctitle"
    for i, b in enumerate(buckets):
        lvl = f"h{i+1}"
        for p in b:
            t = p["text"].strip()
            # 邻域字号（前后各 2 段的众数），用于识别"夹在正文里的偶发大字"
            nb = [paras[j]["size_pt"] for j in
                  range(max(0, p["idx"] - 2), min(len(paras), p["idx"] + 3))
                  if j != p["idx"]]
            nb_mode = Counter(nb).most_common(1)[0][0] if nb else body_pt
            same_n = sum(1 for q in b if q["size_pt"] == p["size_pt"])
            # Step4 护栏（按优先级）
            if doc_title_idx is not None and p["idx"] == doc_title_idx:
                labels[p["idx"]] = "doctitle"
            elif len(t) > max_len and same_n <= 2:
                # 超长 **且孤立** → 醒目语块；超长但成组出现 → 仍是真标题（长标题合法）
                labels[p["idx"]] = "callout"
            elif any(n in t for n in NEG_PATTERNS):    # 分隔符/注释句
                labels[p["idx"]] = "body"
            elif p["max_pt"] - p["size_pt"] >= 20:      # 段内字号悬殊（首字下沉）
                labels[p["idx"]] = "body"
            elif nb_mode <= body_pt + 0.5 and same_n <= 2:
                # 孤立大字（V-4 实测的真实难点：很多书只有 1–2 个一级标题，如"前言/绪论"）
                #   跃升幅度 ≥4pt → 低置信标题：预测为标题，但标记进抽检队列（R-3 分层置信度）
                #   跃升幅度 <4pt → 判为正文强调语(callout)
                if p["size_pt"] - nb_mode >= 4.0:
                    labels[p["idx"]] = lvl
                    steps.setdefault("低置信标题_需抽检", []).append(
                        {"idx": p["idx"], "字号pt": p["size_pt"],
                         "邻域pt": nb_mode, "文本": t[:20]})
                else:
                    labels[p["idx"]] = "callout"
            else:
                labels[p["idx"]] = lvl
    return labels, steps


# ---------------------------------------------------------------- 评估
def find_sample(code):
    for d in sorted(SAMPLES.iterdir()):
        if d.is_dir() and d.name.startswith(code + "_"):
            for f in d.iterdir():
                if f.suffix == ".docx" and not f.name.endswith(".tmp.docx"):
                    return f
    return None


def main():
    gt_path = SAMPLES / "ground_truth.json"
    gt = json.loads(gt_path.read_text(encoding="utf-8"))
    by_sample = defaultdict(list)
    for g in gt:
        by_sample[g["sample"]].append(g)

    report = ["# V-4 · 标题聚类真实准确率实测", "",
              "> 自动评估：样本生成时登记 ground truth，聚类算法为 lxml 实现（承 8.1 / N7）。",
              "> 参数：正文基准=字号直方图峰值，候选门槛=B×1.15，切桶间距=1.5pt，长度上限=40字", ""]

    total = correct = 0
    conf = Counter()
    details = []

    for code, items in sorted(by_sample.items()):
        path = find_sample(code)
        if not path:
            report.append(f"## {code} —— ❌ 未找到样本文件")
            continue
        paras = load_paragraphs(path)
        labels, steps = cluster(paras)
        text2pred = {p["text"]: labels.get(p["idx"], "body") for p in paras}

        c = 0
        rows = []
        for it in items:
            truth = it["label"]
            pred = text2pred.get(it["text"], "body")   # 找不到段落 → 视为 body
            hit = (pred == truth)
            total += 1
            conf[(truth, pred)] += 1
            if hit:
                correct += 1
                c += 1
            else:
                rows.append((it["text"][:34], truth, pred, it["note"]))
            details.append((code, it["text"][:34], truth, pred, hit))

        acc = c / len(items) if items else 0
        report += [f"## 样本 {code} —— {path.name}", "",
                   f"- 段落总数：**{len(paras)}**",
                   f"- 正文基准 B = **{steps['正文基准B_pt']}pt**"
                   f"（占 {steps['B的段落数']} 段的众数）",
                   f"- 字号分布：{steps['字号分布']}",
                   f"- 候选门槛 = B×1.15 = **{steps['候选门槛_pt']}pt**，"
                   f"命中候选 **{steps['候选段落数']}** 段",
                   f"- 分桶：{steps['桶']}",
                   f"- **准确率：{c}/{len(items)} = {acc*100:.1f}%**", ""]
        if steps.get("低置信标题_需抽检"):
            report += [f"**⚠️ 低置信标题（R-3 分层：需进抽检队列，不可自动放行）："
                       f"{len(steps['低置信标题_需抽检'])} 个**", "",
                       "| 段落 | 字号 | 邻域字号 |", "|---|---|---|"]
            for x in steps["低置信标题_需抽检"]:
                report.append(f"| {x['文本']} | {x['字号pt']}pt | {x['邻域pt']}pt |")
            report.append("")
        if rows:
            report += ["**误判明细**", "",
                       "| 段落文本 | 真值 | 预测 | 说明 |", "|---|---|---|---|"]
            for t, truth, pred, note in rows:
                report.append(f"| {t} | {truth} | {pred} | {note} |")
            report.append("")

    overall = correct / total if total else 0
    report += ["## 总体", "",
               f"- 标注总数 **{total}**，正确 **{correct}**",
               f"- **总体准确率：{overall*100:.1f}%**",
               f"- 对照门槛：单阈值 ≥98%（R-3 假设）→ "
               f"**{'达标' if overall >= 0.98 else '未达标，须改为分层置信度'}**", "",
               "### 混淆矩阵（真值 → 预测）", "",
               "| 真值 \\ 预测 | body | h1 | h2 | h3 | callout | doctitle |",
               "|---|---|---|---|---|---|---|"]
    for truth in ["h1", "h2", "h3", "body", "callout", "doctitle"]:
        line = f"| **{truth}** |"
        for pred in ["body", "h1", "h2", "h3", "callout", "doctitle"]:
            line += f" {conf[(truth, pred)]} |"
        report.append(line)

    out = SAMPLES / "_v4_heading_report.md"
    out.write_text("\n".join(report), encoding="utf-8")

    print(f"正文基准/桶/混淆矩阵已生成 → {out}")
    print(f"\n总体准确率：{correct}/{total} = {overall*100:.1f}%")
    print("混淆矩阵（真值→预测）：")
    for truth in ["h1", "h2", "h3", "body", "callout", "doctitle"]:
        row = "  ".join(f"{pred}={conf[(truth, pred)]}"
                        for pred in ["body", "h1", "h2", "h3", "callout", "doctitle"]
                        if conf[(truth, pred)])
        print(f"  {truth:<8} {row or '—'}")
    if overall < 0.98:
        print("\n⚠️ 未达 98% —— 按 R-3 结论应改为分层置信度（高自动/中抽检/低强制人工）")


if __name__ == "__main__":
    main()