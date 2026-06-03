"""
C宝 PDF 阅读 - 把 PDF 页面转图片，用豆包视觉模型识别
用法: python read_pdf.py <PDF路径> [页码，默认全部] [可选问题]
"""
import sys
import os
import fitz  # PyMuPDF
import tempfile
import subprocess

VISION_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vision.py")

def pdf_to_images(pdf_path, page_num=None):
    """将 PDF 页面转为临时图片"""
    doc = fitz.open(pdf_path)
    total = doc.page_count
    temp_dir = tempfile.mkdtemp(prefix="pdf_pages_")

    pages_to_convert = [page_num] if page_num is not None else range(total)
    image_paths = []

    for i in pages_to_convert:
        page = doc[i]
        # 放大 2 倍，保证清晰
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat)
        img_path = os.path.join(temp_dir, f"page_{i+1}.png")
        pix.save(img_path)
        image_paths.append((i + 1, img_path))

    doc.close()
    return image_paths, total

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python read_pdf.py <PDF路径> [页码] [可选问题]")
        print("示例: python read_pdf.py report.pdf")
        print("示例: python read_pdf.py report.pdf 3")
        print("示例: python read_pdf.py report.pdf 1 \"这个图表中的数据是多少？\"")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"[C宝] 找不到文件: {pdf_path}")
        sys.exit(1)

    page_num = int(sys.argv[2]) - 1 if len(sys.argv) > 2 and sys.argv[2].isdigit() else None
    question = sys.argv[3] if len(sys.argv) > 3 else None

    if page_num is None and len(sys.argv) > 2 and sys.argv[2].isdigit():
        # 只有页码没有问题
        pass
    elif page_num is None and len(sys.argv) > 2:
        # 第二个参数是问题
        question = sys.argv[2]

    print(f"[C宝] 正在阅读 PDF: {pdf_path}")

    images, total = pdf_to_images(pdf_path, page_num)
    print(f"[C宝] 共 {total} 页，正在识别 {len(images)} 页...")

    for page_no, img_path in images:
        print(f"\n{'='*60}")
        print(f"  第 {page_no} 页")
        print(f"{'='*60}")

        cmd = ["python", VISION_SCRIPT, img_path]
        if question:
            cmd.append(question)

        subprocess.run(cmd, encoding="utf-8", errors="replace")

    # 清理临时文件
    import shutil
    shutil.rmtree(os.path.dirname(images[0][1]))
