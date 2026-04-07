"""
PDF 每页去白边并拼回（可变页尺寸）。
用图像算法检测内容边界，直接设置 CropBox 裁剪 PDF，避免重渲染导致文字模糊。
"""
import argparse
import sys

import numpy as np

try:
    import fitz
except ImportError:
    try:
        import pymupdf as fitz
    except ImportError:
        fitz = None


def _content_bbox_from_pixmap(pix, white_thresh=250, margin_ratio=0.02):
    """从 pixmap 检测内容边界（像素），并外扩 margin_ratio 比例的白框。返回 (x0,y0,x1,y1)。"""
    h, w = pix.height, pix.width
    n = pix.n
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, n)
    gray = arr.max(axis=2) if n >= 3 else arr.squeeze()
    nonwhite = gray < white_thresh

    rows = nonwhite.any(axis=1)
    cols = nonwhite.any(axis=0)
    if not rows.any() or not cols.any():
        return 0, 0, w, h

    y0, y1 = int(np.argmax(rows)), int(h - 1 - np.argmax(rows[::-1]))
    x0, x1 = int(np.argmax(cols)), int(w - 1 - np.argmax(cols[::-1]))

    mw = max(1, int(w * margin_ratio))
    mh = max(1, int(h * margin_ratio))
    x0 = max(0, x0 - mw)
    y0 = max(0, y0 - mh)
    x1 = min(w, x1 + mw)
    y1 = min(h, y1 + mh)
    return x0, y0, x1, y1


def _pixel_bbox_to_pdf_rect(px0, py0, px1, py1, page_rect, pix_w, pix_h):
    """将像素 bbox 转为 PDF 坐标 (fitz.Rect)。"""
    sx = page_rect.width / pix_w
    sy = page_rect.height / pix_h
    return fitz.Rect(
        page_rect.x0 + px0 * sx,
        page_rect.y0 + py0 * sy,
        page_rect.x0 + px1 * sx,
        page_rect.y0 + py1 * sy,
    )


def crop_pdf_pages(
    input_path,
    output_path,
    *,
    margin_ratio=0.02,
    white_thresh=250,
    bbox_dpi=150,
):
    """
    读入 PDF，每页用图像算法检测内容边界并设 CropBox 裁掉白边、保留 margin 白框，
    再按页拼回新 PDF；每页尺寸可不同。
    不重渲染内容，仅用低分辨率图算 bbox，避免文字失真。
    """
    if fitz is None:
        raise RuntimeError(
            "请安装 pymupdf: pip install pymupdf numpy。"
            "若已安装仍报错，请用当前环境的 python 运行（例如 python Pdftopng.py input.pdf），不要用 py -3。"
        )

    doc = fitz.open(input_path)
    out = fitz.open()

    for i in range(len(doc)):
        page = doc[i]
        rect = page.rect
        pix = page.get_pixmap(dpi=bbox_dpi)
        px0, py0, px1, py1 = _content_bbox_from_pixmap(pix, white_thresh, margin_ratio)
        crop_rect = _pixel_bbox_to_pdf_rect(px0, py0, px1, py1, rect, pix.width, pix.height)
        page.set_cropbox(crop_rect)
        out.insert_pdf(doc, from_page=i, to_page=i)

    out.save(output_path, deflate=True)
    out.close()
    doc.close()


def main():
    parser = argparse.ArgumentParser(description="PDF 每页去白边并拼回（可变页尺寸）")
    parser.add_argument("input", nargs="?", default="input.pdf", help="输入 PDF")
    parser.add_argument("output", nargs="?", default=None, help="输出 PDF，默认 input_cropped.pdf")
    parser.add_argument("--margin", type=float, default=0.02, help="保留白框比例 (0~1)，默认 0.02")
    parser.add_argument("--white-thresh", type=int, default=250, help="低于此灰度视为内容，默认 250")
    parser.add_argument("--dpi", type=int, default=150, help="仅用于检测边界的渲染 DPI，默认 150")
    args = parser.parse_args()

    out = args.output
    if out is None:
        base = args.input.rsplit(".", 1)[0] if "." in args.input else args.input
        out = f"{base}_cropped.pdf"

    try:
        crop_pdf_pages(
            args.input,
            out,
            margin_ratio=args.margin,
            white_thresh=args.white_thresh,
            bbox_dpi=args.dpi,
        )
        print(f"已保存: {out}")
    except FileNotFoundError:
        print(f"错误: 找不到文件 {args.input}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
