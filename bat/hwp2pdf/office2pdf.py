
# -*- coding: utf-8 -*-
"""
한글 / 엑셀 / 워드 / PDF 일괄 변환 및 병합

- 지정 폴더 내부 파일만 처리
- 하위 폴더 탐색 제외
- 파일명 기준 자연 정렬
- PDF 폴더 자동 생성
- 원본 파일 보존

필요 환경:
    Windows
    한컴오피스 한글
    Microsoft Excel
    Microsoft Word

패키지 설치:
    pip install pywin32 pypdf
"""

import os
import re
import sys
import shutil
import uuid
import traceback

try:
    import pythoncom
    import win32com.client
    from pypdf import PdfReader, PdfWriter

except ImportError as e:
    print(f"[오류] 필요한 패키지가 없습니다: {e}")
    print("설치 명령: pip install pywin32 pypdf")
    sys.exit(2)


# ==================================================
# 1. 기본 설정
# ==================================================

HWP_EXTENSIONS = (".hwp", ".hwpx")

EXCEL_EXTENSIONS = (
    ".xls", ".xlsx", ".xlsm", ".xlsb"
)

WORD_EXTENSIONS = (
    ".doc", ".docx", ".docm"
)

PDF_EXTENSIONS = (".pdf",)

ALL_EXTENSIONS = (
    HWP_EXTENSIONS
    + EXCEL_EXTENSIONS
    + WORD_EXTENSIONS
    + PDF_EXTENSIONS
)

OUTPUT_DIR_NAME = "PDF"
DEFAULT_MERGED_NAME = "0. 통합본.pdf"

HWP_MESSAGE_MODE = 0x00020021


# ==================================================
# 2. 파일명 자연 정렬
# ==================================================

def natural_key(name):
    parts = re.split(r"([0-9]+)", name.lower())

    return [
        int(p) if i % 2 == 1 else p
        for i, p in enumerate(parts)
    ]


# ==================================================
# 3. 해당 폴더의 파일만 검색
# ==================================================

def find_files(folder, merged_name):

    files = []

    for name in os.listdir(folder):

        path = os.path.join(folder, name)

        # 하위 폴더 제외
        if not os.path.isfile(path):
            continue

        # 임시 파일 제외
        if name.startswith("~"):
            continue

        # 기존 통합본 제외
        if name.lower() == merged_name.lower():
            continue

        ext = os.path.splitext(name)[1].lower()

        if ext in ALL_EXTENSIONS:
            files.append(path)

    return sorted(
        files,
        key=lambda p: natural_key(os.path.basename(p))
    )


# ==================================================
# 4. PDF 파일명 중복 방지
# ==================================================

def make_output_names(files, merged_name):

    used = {merged_name.lower()}
    results = []

    for src in files:

        stem, ext = os.path.splitext(
            os.path.basename(src)
        )

        pdf_name = stem + ".pdf"

        if pdf_name.lower() in used:
            pdf_name = f"{stem}_{ext[1:]}.pdf"

        count = 2
        base_name = os.path.splitext(pdf_name)[0]

        while pdf_name.lower() in used:
            pdf_name = f"{base_name}_{count}.pdf"
            count += 1

        used.add(pdf_name.lower())
        results.append((src, pdf_name))

    return results


# ==================================================
# 5. PDF 정상 여부 확인
# ==================================================

def validate_pdf(path):

    if not os.path.isfile(path):
        raise RuntimeError("PDF 파일이 생성되지 않았습니다.")

    if os.path.getsize(path) == 0:
        raise RuntimeError("PDF 파일의 용량이 0입니다.")

    with open(path, "rb") as f:
        reader = PdfReader(f)

        if reader.is_encrypted:
            raise RuntimeError("암호화된 PDF입니다.")

        pages = len(reader.pages)

        if pages == 0:
            raise RuntimeError("PDF 페이지가 없습니다.")

    return pages


# ==================================================
# 6. 한글 프로그램 실행
# ==================================================

def open_hwp():

    hwp = win32com.client.DispatchEx(
        "HWPFrame.HwpObject"
    )

    try:
        hwp.RegisterModule(
            "FilePathCheckDLL",
            "FilePathCheckerModule"
        )
    except Exception:
        pass

    try:
        hwp.SetMessageBoxMode(HWP_MESSAGE_MODE)
    except Exception:
        pass

    try:
        hwp.XHwpWindows.Item(0).Visible = False
    except Exception:
        pass

    return hwp


# ==================================================
# 7. 엑셀 프로그램 실행
# ==================================================

def open_excel():

    excel = win32com.client.DispatchEx(
        "Excel.Application"
    )

    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False

    # 매크로 자동 실행 방지
    excel.AutomationSecurity = 3

    return excel


# ==================================================
# 8. 워드 프로그램 실행
# ==================================================

def open_word():

    word = win32com.client.DispatchEx(
        "Word.Application"
    )

    word.Visible = False
    word.DisplayAlerts = 0

    # 매크로 자동 실행 방지
    word.AutomationSecurity = 3

    return word


# ==================================================
# 9. 한글 PDF 변환
# ==================================================

def convert_hwp(hwp, src, out_path):

    try:
        result = hwp.Open(
            src,
            "",
            "forceopen:true"
        )

        if result is not None and not result:
            raise RuntimeError("한글 문서 열기 실패")

        result = hwp.SaveAs(
            out_path,
            "PDF",
            ""
        )

        if result is not None and not result:
            raise RuntimeError("한글 PDF 저장 실패")

    finally:
        try:
            hwp.Clear(1)
        except Exception:
            pass


# ==================================================
# 10. 엑셀 PDF 변환
# ==================================================

def convert_excel(excel, src, out_path):

    workbook = None

    try:
        workbook = excel.Workbooks.Open(
            Filename=src,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False
        )

        workbook.ExportAsFixedFormat(
            Type=0,
            Filename=out_path,
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)


# ==================================================
# 11. 워드 PDF 변환
# ==================================================

def convert_word(word, src, out_path):

    document = None

    try:
        document = word.Documents.Open(
            FileName=src,
            ReadOnly=True,
            AddToRecentFiles=False,
            ConfirmConversions=False
        )

        document.ExportAsFixedFormat(
            OutputFileName=out_path,
            ExportFormat=17,
            OpenAfterExport=False
        )

    finally:
        if document is not None:
            document.Close(SaveChanges=0)


# ==================================================
# 12. Office 프로그램 관리
# ==================================================

class OfficeManager:

    def __init__(self):
        self.hwp = None
        self.excel = None
        self.word = None

    def convert(self, src, out_path):

        ext = os.path.splitext(src)[1].lower()

        if ext in HWP_EXTENSIONS:

            if self.hwp is None:
                self.hwp = open_hwp()

            convert_hwp(
                self.hwp,
                src,
                out_path
            )

        elif ext in EXCEL_EXTENSIONS:

            if self.excel is None:
                self.excel = open_excel()

            convert_excel(
                self.excel,
                src,
                out_path
            )

        elif ext in WORD_EXTENSIONS:

            if self.word is None:
                self.word = open_word()

            convert_word(
                self.word,
                src,
                out_path
            )

        elif ext in PDF_EXTENSIONS:

            # 기존 PDF는 변환 없이 복사
            shutil.copy2(src, out_path)

        else:
            raise RuntimeError(
                f"지원하지 않는 확장자: {ext}"
            )

    def close(self):

        for name, app in (
            ("한글", self.hwp),
            ("엑셀", self.excel),
            ("워드", self.word)
        ):

            if app is not None:
                try:
                    app.Quit()
                except Exception as e:
                    print(f"[경고] {name} 종료 실패: {e}")


# ==================================================
# 13. 개별 파일 변환
# ==================================================

def convert_one(manager, src, out_path):

    # 기존 PDF 보존을 위한 임시 파일
    temp_path = os.path.join(
        os.path.dirname(out_path),
        f"_temp_{uuid.uuid4().hex}.pdf"
    )

    try:

        manager.convert(src, temp_path)

        pages = validate_pdf(temp_path)

        # 정상 변환된 경우에만 기존 결과 교체
        os.replace(temp_path, out_path)

        return pages

    finally:

        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


# ==================================================
# 14. PDF 병합
# ==================================================

def merge_pdfs(pdf_paths, merged_path):

    writer = PdfWriter()

    expected_pages = 0

    temp_path = os.path.join(
        os.path.dirname(merged_path),
        f"_merge_{uuid.uuid4().hex}.pdf"
    )

    try:

        for path in pdf_paths:

            pages = validate_pdf(path)

            writer.append(path)

            expected_pages += pages

        with open(temp_path, "wb") as f:
            writer.write(f)

        actual_pages = validate_pdf(temp_path)

        # 페이지 누락 검증
        if actual_pages != expected_pages:

            raise RuntimeError(
                f"페이지 수 불일치: "
                f"예상 {expected_pages}, "
                f"실제 {actual_pages}"
            )

        os.replace(temp_path, merged_path)

        return actual_pages

    finally:

        writer.close()

        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


# ==================================================
# 15. 메인 실행
# ==================================================

def main():

    # 대상 폴더
    root = (
        os.path.abspath(sys.argv[1])
        if len(sys.argv) > 1 and sys.argv[1]
        else os.path.dirname(
            os.path.abspath(__file__)
        )
    )

    # 통합본 파일명
    merged_name = (
        sys.argv[2]
        if len(sys.argv) > 2 and sys.argv[2]
        else DEFAULT_MERGED_NAME
    )

    if not merged_name.lower().endswith(".pdf"):
        merged_name += ".pdf"

    # 통합본 파일명 검증
    if (
        os.path.basename(merged_name) != merged_name
        or re.search(r'[<>:"/\\|?*]', merged_name)
    ):
        print("[오류] 통합본 파일명이 올바르지 않습니다.")
        return 1

    if not os.path.isdir(root):
        print(f"[오류] 폴더가 없습니다: {root}")
        return 1

    files = find_files(root, merged_name)

    if not files:
        print("[알림] 변환할 파일이 없습니다.")
        return 1

    out_dir = os.path.join(
        root,
        OUTPUT_DIR_NAME
    )

    os.makedirs(out_dir, exist_ok=True)

    merged_path = os.path.join(
        out_dir,
        merged_name
    )

    total = len(files)

    success = []
    failures = []
    total_pages = 0

    manager = OfficeManager()

    print("=" * 60)
    print("Office 문서 PDF 일괄 변환 및 병합")
    print("=" * 60)
    print(f"대상 폴더: {root}")
    print(f"대상 파일: {total}개")
    print(f"저장 위치: {out_dir}")
    print("하위 폴더: 처리하지 않음")
    print("=" * 60)

    pythoncom.CoInitialize()

    try:

        jobs = make_output_names(
            files,
            merged_name
        )

        for index, (src, pdf_name) in enumerate(
            jobs,
            start=1
        ):

            out_path = os.path.join(
                out_dir,
                pdf_name
            )

            progress = index / total * 100

            print(
                f"\n[{index}/{total}] "
                f"({progress:.1f}%) "
                f"{os.path.basename(src)}"
            )

            try:

                pages = convert_one(
                    manager,
                    src,
                    out_path
                )

                success.append(out_path)
                total_pages += pages

                print(
                    f"  [성공] {pdf_name} "
                    f"({pages}페이지)"
                )

            except Exception as e:

                failures.append(
                    (os.path.basename(src), str(e))
                )

                print(f"  [실패] {e}")

    except KeyboardInterrupt:

        print("\n[중단] 사용자가 작업을 중단했습니다.")
        return 1

    finally:

        manager.close()
        pythoncom.CoUninitialize()

    # --------------------------------------------------
    # PDF 통합본 생성
    # --------------------------------------------------

    merged_ok = False

    if success:

        # 실패 파일이 있으면 부분 통합본으로 구분
        if failures:

            stem = os.path.splitext(merged_name)[0]

            partial_name = (
                f"{stem}_일부누락_"
                f"{uuid.uuid4().hex[:8]}.pdf"
            )

            merged_path = os.path.join(
                out_dir,
                partial_name
            )

        print("\n" + "=" * 60)
        print("PDF 병합 시작")
        print("=" * 60)

        try:

            merged_pages = merge_pdfs(
                success,
                merged_path
            )

            merged_ok = True

            print(
                f"[병합 완료] "
                f"{os.path.basename(merged_path)}"
            )

            print(f"병합 파일: {len(success)}개")
            print(f"전체 페이지: {merged_pages}페이지")

        except Exception as e:

            failures.append(
                ("통합본 PDF", str(e))
            )

            print(f"[병합 실패] {e}")

    # --------------------------------------------------
    # 결과 출력
    # --------------------------------------------------

    print("\n" + "=" * 60)
    print("최종 작업 결과")
    print("=" * 60)

    print(f"전체 대상 : {total}개")
    print(f"처리 성공 : {len(success)}개")
    print(f"원본 실패 : {total - len(success)}개")
    print(f"성공 페이지 : {total_pages}페이지")
    print(f"전체 오류 : {len(failures)}건")

    if failures:

        print("\n[실패 목록]")

        for name, reason in failures:
            print(f"- {name}: {reason}")

        print("\n[주의] 누락된 파일을 확인하세요.")

    if merged_ok:

        print("\n[통합본 저장 위치]")
        print(merged_path)

    # --------------------------------------------------
    # 작업 결과 로그 저장
    # --------------------------------------------------

    log_path = os.path.join(
        out_dir,
        "변환결과.txt"
    )

    try:

        with open(
            log_path,
            "w",
            encoding="utf-8-sig"
        ) as log:

            log.write("PDF 일괄 변환 결과\n")
            log.write("=" * 50 + "\n")
            log.write(f"대상 폴더: {root}\n")
            log.write(f"전체 대상: {total}개\n")
            log.write(f"처리 성공: {len(success)}개\n")
            log.write(f"전체 오류: {len(failures)}건\n")
            log.write(f"성공 페이지: {total_pages}\n")

            log.write("\n[성공 파일]\n")

            for path in success:
                log.write(
                    os.path.basename(path) + "\n"
                )

            log.write("\n[실패 목록]\n")

            for name, reason in failures:
                log.write(f"{name}: {reason}\n")

            if merged_ok:
                log.write(
                    f"\n통합본: {merged_path}\n"
                )

    except OSError as e:

        print(f"[경고] 로그 저장 실패: {e}")

    # --------------------------------------------------
    # 종료 코드
    # --------------------------------------------------

    if failures or not merged_ok:
        return 1

    print("\n[완료] 모든 작업이 정상적으로 완료되었습니다.")

    return 0


if __name__ == "__main__":

    try:
        sys.exit(main())

    except Exception:

        traceback.print_exc()
        sys.exit(2)
