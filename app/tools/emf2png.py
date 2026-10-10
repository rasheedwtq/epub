# -*- coding: utf-8 -*-
"""V-13b 公式高清栅格化：Word COM CopyAsPicture → 剪贴板 EMF → GDI 矢量渲染 PNG
纯 ctypes GDI，无外部程序依赖。输入 EMF bytes，输出 PIL Image。
"""
import ctypes
from ctypes import Structure, c_long, c_uint32, c_void_p, byref, sizeof
from PIL import Image

gdi32 = ctypes.windll.gdi32
user32 = ctypes.windll.user32

class RECT(Structure):
    _fields_ = [("left", c_long), ("top", c_long), ("right", c_long), ("bottom", c_long)]

class SIZEL(Structure):
    _fields_ = [("cx", c_long), ("cy", c_long)]

class ENHMETAHEADER(Structure):
    _fields_ = [("iType", c_uint32), ("nSize", c_uint32),
                ("rclBounds", RECT), ("rclFrame", RECT),
                ("dSignature", c_uint32), ("nVersion", c_uint32),
                ("nBytes", c_uint32), ("nRecords", c_uint32),
                ("sHandles", c_uint32), ("sReserved", c_uint32),
                ("nDescription", c_uint32), ("offDescription", c_uint32),
                ("nPalEntries", c_uint32), ("szlDevice", SIZEL), ("szlMillimeters", SIZEL)]

class BITMAPINFOHEADER(Structure):
    _fields_ = [("biSize", c_uint32), ("biWidth", c_long), ("biHeight", c_long),
                ("biPlanes", ctypes.c_ushort), ("biBitCount", ctypes.c_ushort), ("biCompression", c_uint32),
                ("biSizeImage", c_uint32), ("biXPelsPerMeter", c_long), ("biYPelsPerMeter", c_long),
                ("biClrUsed", c_uint32), ("biClrImportant", c_uint32)]

class RGBQUAD(Structure):
    _fields_ = [("rgbBlue", ctypes.c_ubyte), ("rgbGreen", ctypes.c_ubyte),
                ("rgbRed", ctypes.c_ubyte), ("rgbReserved", ctypes.c_ubyte)]

class BITMAPINFO(Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", RGBQUAD * 1)]

# Win64 句柄是 64 位：必须显式声明 restype/argtypes，否则句柄被截断
for fn in ("SetEnhMetaFileBits", "GetEnhMetaFileHeader", "DeleteEnhMetaFile",
           "CreateCompatibleDC", "CreateDIBSection", "SelectObject", "DeleteObject",
           "DeleteDC", "PlayEnhMetaFile", "PatBlt"):
    getattr(gdi32, fn).restype = c_void_p
gdi32.SetEnhMetaFileBits.argtypes = [ctypes.c_uint, ctypes.c_char_p]
gdi32.GetEnhMetaFileHeader.argtypes = [c_void_p, ctypes.c_uint, c_void_p]
gdi32.PlayEnhMetaFile.argtypes = [c_void_p, c_void_p, ctypes.POINTER(RECT)]
gdi32.CreateDIBSection.argtypes = [c_void_p, c_void_p, ctypes.c_uint, ctypes.POINTER(c_void_p), c_void_p, ctypes.c_uint]
gdi32.PatBlt.argtypes = [c_void_p, c_long, c_long, c_long, c_long, ctypes.c_uint]
gdi32.DeleteEnhMetaFile.argtypes = [c_void_p]
gdi32.CreateCompatibleDC.argtypes = [c_void_p]
gdi32.SelectObject.argtypes = [c_void_p, c_void_p]
gdi32.DeleteObject.argtypes = [c_void_p]
gdi32.DeleteDC.argtypes = [c_void_p]
user32.GetDC.restype = c_void_p
user32.ReleaseDC.argtypes = [c_void_p, c_void_p]

WHITENESS = 0x00FF0062

def emf_bytes_to_png(emf_bytes: bytes, scale: int = 4) -> Image.Image:
    """EMF 二进制 → 高分辨率 PNG（白底 RGB）。scale 相对 EMF 逻辑尺寸的放大倍数。"""
    hemf = gdi32.SetEnhMetaFileBits(len(emf_bytes), emf_bytes)
    if not hemf:
        raise RuntimeError("SetEnhMetaFileBits failed")
    try:
        hdr = ENHMETAHEADER()
        if not gdi32.GetEnhMetaFileHeader(hemf, sizeof(hdr), byref(hdr)):
            raise RuntimeError("GetEnhMetaFileHeader failed")
        b = hdr.rclBounds
        f = hdr.rclFrame
        # 物理比例以 rclFrame(0.01mm) 为准，避免 bounds/frame 不一致导致拉伸
        if f.right > f.left and f.bottom > f.top:
            fw_mm = (f.right - f.left) / 100.0
            fh_mm = (f.bottom - f.top) / 100.0
            px_per_mm = 96.0 / 25.4
            W = round(fw_mm * px_per_mm * scale)
            H = round(fh_mm * px_per_mm * scale)
        else:
            w0, h0 = b.right - b.left, b.bottom - b.top
            W, H = w0 * scale, h0 * scale
        if W <= 0 or H <= 0:
            raise RuntimeError(f"bad EMF frame: {f}")
        # 32bpp top-down DIB
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = W
        bmi.bmiHeader.biHeight = -H          # top-down
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0      # BI_RGB
        hdcScreen = user32.GetDC(0)
        memDC = gdi32.CreateCompatibleDC(hdcScreen)
        ppv = c_void_p()
        hbmp = gdi32.CreateDIBSection(memDC, byref(bmi), 0, byref(ppv), None, 0)
        if not hbmp or not ppv.value:
            raise RuntimeError("CreateDIBSection failed")
        old = gdi32.SelectObject(memDC, hbmp)
        try:
            gdi32.PatBlt(memDC, 0, 0, W, H, WHITENESS)   # 白底
            rect = RECT(0, 0, W, H)
            if not gdi32.PlayEnhMetaFile(memDC, hemf, byref(rect)):
                raise RuntimeError("PlayEnhMetaFile failed")
            buf = ctypes.string_at(ppv.value, W * H * 4)
        finally:
            gdi32.SelectObject(memDC, old)
            gdi32.DeleteObject(hbmp)
            gdi32.DeleteDC(memDC)
            user32.ReleaseDC(0, hdcScreen)
        img = Image.frombuffer("RGBA", (W, H), buf, "raw", "BGRA", 0, 1)
        return img.convert("RGB")
    finally:
        gdi32.DeleteEnhMetaFile(hemf)

if __name__ == "__main__":
    import os, glob
    for f in sorted(glob.glob(r"D:\project\epub\work\v13b_math_hd\math_*.emf")):
        data = open(f, "rb").read()
        img = emf_bytes_to_png(data)
        out = f[:-4] + "_gdi.png"
        img.save(out)
        print(os.path.basename(f), "→", img.size, out)
