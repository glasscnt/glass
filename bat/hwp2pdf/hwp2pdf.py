# HWP -> PDF 변환 후 파일명 순서대로 병합
# 사용법: python hwp2pdf.py <대상폴더> [통합본파일명]
# 필요: 한컴오피스(한글), Python 패키지 pywin32, pypdf
import sys, os, glob
import win32com.client as win32
from pypdf import PdfWriter

src = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.getcwd())
merged_name = sys.argv[2] if len(sys.argv) > 2 and sys.argv[2] else "0. 통합본.pdf"
if not merged_name.lower().endswith(".pdf"):
    merged_name += ".pdf"
out_dir = os.path.join(src, "PDF")
os.makedirs(out_dir, exist_ok=True)

files = sorted(glob.glob(os.path.join(src, "*.hwp")) + glob.glob(os.path.join(src, "*.hwpx")),
               key=lambda p: os.path.basename(p))
files = [f for f in files if not os.path.basename(f).startswith("~")]
if not files:
    print("HWP 파일이 없습니다:", src)
    sys.exit(1)

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
try:
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
except Exception:
    pass
hwp.SetMessageBoxMode(0x00020021)

pdfs = []
try:
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0]
        out = os.path.join(out_dir, name + ".pdf")
        hwp.Open(f, "", "forceopen:true")
        hwp.SaveAs(out, "PDF", "")
        hwp.Clear(1)
        if os.path.exists(out):
            print("변환 완료:", os.path.basename(out))
            pdfs.append(out)
        else:
            print("변환 실패:", os.path.basename(f))
finally:
    hwp.Quit()

w = PdfWriter()
for p in pdfs:
    w.append(p)
merged = os.path.join(out_dir, merged_name)
with open(merged, "wb") as fh:
    w.write(fh)
print("병합 완료:", merged_name, f"({len(pdfs)}개 파일)")
