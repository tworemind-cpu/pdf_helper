# -*- coding: utf-8 -*-
"""
PDF 도우미
 메뉴 순서
 1) 합치기 · 분리하기 · 회전하기
     · 여러 PDF를 목록에 담고(추가/삭제/순서 변경), 모든 페이지를 미리보기로 확인
     · 미리보기에서 페이지를 끌어서 순서 바꾸기 (체크한 여러 쪽을 한꺼번에 끌어 옮길 수 있음)
     · 미리보기 그림 크기를 슬라이더로 조절
     · 페이지 삭제
     · [NEW] 체크한 페이지를 왼쪽 90° / 오른쪽 90° / 180° 로 회전 (애니메이션 없이 바로 미리보기에 반영)
       (회전한 상태 그대로 저장·낱장 분리에 반영되며, 원래 'PDF 회전하기' 메뉴는 이 화면으로 합쳐졌습니다)
     · 전체를 하나로 합쳐 저장 / 체크한 페이지만 합쳐 저장 / 체크한 페이지를 낱장으로 분리
     · 파일 목록의 순서를 바꾸면 페이지 순서도 목록 순서대로 다시 정렬됨
       (이미 페이지 순서를 조정했거나 지우거나 회전한 상태라면 초기화된다는 경고창이 뜸)
     · '쪽 번호 넣기'를 체크하면 저장되는 PDF의 아래쪽 가운데에 '1/전체쪽수' 형태로 번호가 찍힘
       (기본값은 체크 해제 = 번호 없음. 낱장 분리 시에는 미리보기에 보이는 순서 번호가 찍힘)
 2) PDF로 변환하기 (PDF·한글·엑셀·워드·파워포인트·이미지 -> 하나의 PDF)
     · 여기서도 '쪽 번호 넣기'를 체크하면 합쳐진 결과에 1/전체쪽수 번호가 이어서 찍힘
 3) [NEW] PDF 마스킹 (개인정보 등 가릴 곳을 검은색으로 '삭제')
     · 첫 쪽 미리보기에서 마우스로 영역을 끌어 정하면, 모든 쪽의 같은 위치가 검은색으로 바뀜
     · 단순히 검은 도형을 덮는 것이 아니라 그 자리의 글자·그림 자체를 지우므로 복구할 수 없음
     · 영역은 여러 개 지정할 수 있고, 쪽 크기가 달라도 같은 비율 위치에 적용됨
 4) 이미지 변환하기 (PDF -> jpg / png / bmp)
 5) 용량 줄이기 (페이지별 진행도 표시)
 - 모든 화면에서 파일을 창 위로 끌어다 놓기(드래그앤드롭) 가능
 - 파일 목록은 마우스로 끌어서 순서를 바꿀 수 있음
 - 파일 목록: Delete 키로 선택 파일 삭제, 파일 이름 오름차순/내림차순 정렬
 - 프로그램을 켜면 새 버전이 있는지 확인하고, 있으면 업데이트 안내 팝업을 띄움
   (인터넷이 막혀 있거나 확인에 실패하면 아무 팝업 없이 조용히 넘어감)

설치:  pip install pymupdf customtkinter pillow pywin32 tkinterdnd2
       (업데이트 확인 기능은 파이썬 기본 라이브러리만 사용하므로 추가 설치 없음)
실행:  python pdf_helper.py

[version.json 예시]  (VERSION_INFO_SOURCE 위치에 올려 두세요)
{
  "version": "2.1.0",
  "download_url": "https://github.com/내계정/내저장소/releases/latest",
  "notes": "- 엑셀 취합 속도 개선\\n- 오류 수정"
}
"""
import io
import json
import os
import re
import shutil
import tempfile
import threading
import time
import unicodedata
import urllib.request
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pymupdf  # PyMuPDF (예전 이름 fitz 는 경고가 떠서 pymupdf 로 바꿨습니다)
from PIL import Image, ImageTk
from tkinterdnd2 import TkinterDnD, DND_FILES

OFFICE_EXT = {".doc", ".docx", ".xls", ".xlsx", ".xlsm",
              ".ppt", ".pptx", ".hwp", ".hwpx"}
IMG_EXT = {".jpg", ".jpeg", ".png", ".bmp"}
ALL_EXT = {".pdf"} | OFFICE_EXT | IMG_EXT
BG = "#EEF1F6"
FONT = "Malgun Gothic"

# 미리보기 그림을 '시계 방향으로 90/180/270도' 돌릴 때 쓰는 방법 (Pillow 버전이 달라도 동작하도록 처리)
_T = getattr(Image, "Transpose", Image)
ROTATE_OPS = {90: _T.ROTATE_270, 180: _T.ROTATE_180, 270: _T.ROTATE_90}

# ============================================================
# ▼ 업데이트 확인 설정 (주무관님이 수정하실 부분) ▼
# ============================================================
# 현재 프로그램의 버전 (exe를 새로 만들 때마다 이 숫자를 올려주세요)
CURRENT_VERSION = "3.0.0"

# 최신 버전 정보(version.json)가 있는 위치
# 방법 A) 인터넷 주소 (예: GitHub)
VERSION_INFO_SOURCE = "https://raw.githubusercontent.com/tworemind-cpu/pdf_helper/main/version.json"
# 방법 B) 사내 공유 폴더 (인터넷이 막힌 환경이면 위 줄 대신 아래 형태로 쓰세요)
# VERSION_INFO_SOURCE = r"\\서버이름\공유폴더\version.json"

# 정보를 읽을 때 기다리는 최대 시간(초)
UPDATE_TIMEOUT_SEC = 5
# ============================================================
# 조정하기 화면에서 '어느 파일의 쪽인지' 구분하기 위한 색깔 띠 (파일마다 하나씩 돌아가며 사용)
FILE_COLORS = ["#3B8ED0", "#2F9E7A", "#E8913A", "#9B59B6",
               "#D9534F", "#17A2B8", "#8E9A3B", "#C0558F"]


# ───────────────────────── 공통 도우미 ─────────────────────────
def fmt_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def parse_pages(text, total):
    """'1-3,5,8' -> {0,1,2,4,7} (0부터 시작하는 페이지 번호)"""
    result = set()
    for part in text.replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a), int(b)
        else:
            a = b = int(part)
        if a > b:
            a, b = b, a
        if a < 1 or b > total:
            raise ValueError(f"페이지 범위는 1~{total} 사이여야 합니다.")
        result.update(range(a - 1, b))
    return result


def natural_key(path):
    """파일 이름 정렬용 키. '2번 문서'가 '10번 문서'보다 앞에 오도록 숫자는 숫자로 비교합니다."""
    name = unicodedata.normalize("NFC", os.path.basename(path)).lower()
    parts = re.split(r"(\d+)", name)  # [글자, 숫자, 글자, 숫자, ...] 순서로 나뉨
    return [int(p) if i % 2 else p for i, p in enumerate(parts)]


def same_path(a, b):
    """두 경로가 같은 파일을 가리키는지 비교합니다 (대소문자·슬래시 차이 무시)."""
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def stamp_page_numbers(doc, numbers=None, total=None):
    """PDF의 각 쪽 아래쪽 가운데에 '번호/전체쪽수' 글자를 찍습니다. (저장 직전의 새 PDF에만 사용)
    - numbers: 각 쪽에 찍을 번호 목록. 비워 두면 1, 2, 3... 순서대로 찍습니다.
    - total  : '/' 뒤에 찍을 전체 쪽수. 비워 두면 이 PDF의 전체 쪽수를 씁니다.
    회전된 쪽도 화면에서 보이는 방향 기준으로 바르게 읽히도록 위치와 방향을 맞춥니다."""
    total = total or doc.page_count
    for i, page in enumerate(doc):
        no = numbers[i] if numbers else i + 1
        text = f"{no}/{total}"
        rect = page.rect                                   # 화면에 보이는 쪽의 크기
        size = max(8.0, min(rect.width, rect.height) / 60)  # 쪽 크기에 비례한 글자 크기 (A4 약 10pt)
        width = pymupdf.get_text_length(text, fontname="helv", fontsize=size)
        # 화면에서 본 '아래쪽 가운데' 위치 -> 회전되지 않은 원래 좌표로 변환
        visible = pymupdf.Point((rect.width - width) / 2, rect.height - size * 1.8)
        point = visible * page.derotation_matrix
        page.insert_text(point, text, fontname="helv", fontsize=size,
                         rotate=page.rotation, color=(0.2, 0.2, 0.2), overlay=True)


# ───────────────────────── 업데이트 확인 ─────────────────────────
def parse_version(text):
    """'1.2.10' 같은 글자를 (1, 2, 10) 숫자 묶음으로 바꿉니다.
    글자 그대로 비교하면 '1.10.0'이 '1.9.0'보다 작다고 잘못 나오기 때문입니다."""
    parts = []
    for piece in str(text).strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in piece if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def is_newer(latest, current):
    """latest가 current보다 높은 버전이면 True"""
    a, b = parse_version(latest), parse_version(current)
    length = max(len(a), len(b))
    a += (0,) * (length - len(a))
    b += (0,) * (length - len(b))
    return a > b


def fetch_version_info():
    """version.json 내용을 읽어 딕셔너리로 돌려줍니다. 실패하면 None."""
    try:
        source = VERSION_INFO_SOURCE
        if source.lower().startswith(("http://", "https://")):
            request = urllib.request.Request(source, headers={"User-Agent": "update-checker"})
            with urllib.request.urlopen(request, timeout=UPDATE_TIMEOUT_SEC) as response:
                raw = response.read().decode("utf-8-sig")
        else:
            with open(source, "r", encoding="utf-8-sig") as f:
                raw = f.read()
        info = json.loads(raw)
        if "version" not in info or "download_url" not in info:
            return None
        return info
    except Exception:
        # 네트워크 차단, 파일 없음, 형식 오류 등 모든 경우 조용히 넘어갑니다.
        return None


def show_update_popup(info, parent=None):
    """업데이트 안내 팝업. '예'를 누르면 다운로드 페이지를 엽니다."""
    message = (
        f"새 버전이 있습니다.\n\n"
        f"현재 버전: {CURRENT_VERSION}\n"
        f"최신 버전: {info['version']}\n"
    )
    notes = str(info.get("notes", "")).replace("\\n", "\n").strip()
    if notes:
        message += f"\n[변경 내용]\n{notes}\n"
    message += "\n지금 다운로드 페이지를 여시겠습니까?"

    if messagebox.askyesno("업데이트 안내", message, parent=parent):
        webbrowser.open(info["download_url"])


def check_for_update_async(root):
    """프로그램 창(root)이 뜬 뒤에 호출합니다.
    확인 작업은 뒤에서 따로 진행하므로 프로그램이 멈추지 않습니다."""

    def worker():
        info = fetch_version_info()
        if info and is_newer(info["version"], CURRENT_VERSION):
            # 팝업은 반드시 메인 창 쪽에서 띄워야 안전하므로 root.after를 사용합니다.
            root.after(0, lambda: show_update_popup(info, parent=root))

    threading.Thread(target=worker, daemon=True).start()


def open_pdf(path):
    doc = pymupdf.open(path)
    if doc.needs_pass:
        doc.close()
        raise ValueError("암호가 걸린 PDF는 처리할 수 없습니다.")
    return doc


def page_to_pil(page, dpi=100):
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def wait_for_file(path, timeout):
    """파일이 만들어지고 크기가 더 이상 변하지 않을 때까지 기다립니다. 성공하면 True."""
    end = time.time() + timeout
    last = -1
    while time.time() < end:
        if os.path.exists(path):
            size = os.path.getsize(path)
            if size > 0 and size == last:
                return True
            last = size
        time.sleep(0.3)
    return False


def hwp_to_pdf(src, dst, ext):
    """한글 파일(hwp/hwpx)을 PDF로 저장합니다.
    - 한글 종료(Quit) 단계에서 오류가 나도 '진짜 원인'이 가려지지 않게 따로 처리합니다.
    - 저장이 실패(False)하거나 늦어져도 확인 후 대체 방법을 시도합니다.
    - 한글이 중간에 멈춘 경우를 대비해 한 번 더 새로 시도합니다."""
    import win32com.client as wc
    last_err = None
    for attempt in range(2):
        hwp = None
        try:
            try:
                # 타입 정보를 미리 만들어 쓰는 방식이 더 안정적입니다.
                hwp = wc.gencache.EnsureDispatch("HWPFrame.HwpObject")
            except Exception:
                hwp = wc.Dispatch("HWPFrame.HwpObject")
            try:
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass

            # 1) 파일 열기: 확장자에 맞는 형식 -> 안 되면 한글이 알아서 판단하도록 ""
            opened = False
            first_fmt = "HWPX" if ext == ".hwpx" else "HWP"
            for fmt in (first_fmt, ""):
                try:
                    if hwp.Open(src, fmt, "forceopen:true") is not False:
                        opened = True
                        break
                except Exception as e:
                    last_err = e
            if not opened:
                raise RuntimeError("한글에서 파일을 열지 못했습니다.")

            # 2) PDF로 저장 (SaveAs는 실패해도 오류를 내지 않고 False만 돌려줄 수 있음)
            if os.path.exists(dst):
                os.remove(dst)
            ok = None
            try:
                ok = hwp.SaveAs(dst, "PDF", "")
            except Exception as e:
                last_err = e
            if wait_for_file(dst, 1.5 if ok is False else 20):
                return dst

            # 3) 대체 방법
            try:
                act = hwp.HAction
                pset = hwp.HParameterSet.HFileOpenSave
                act.GetDefault("FileSaveAs_S", pset.HSet)
                pset.filename = dst
                pset.Format = "PDF"
                pset.Attributes = 0
                act.Execute("FileSaveAs_S", pset.HSet)
            except Exception as e:
                last_err = e
            if wait_for_file(dst, 20):
                return dst
            raise RuntimeError("한글이 PDF 파일을 만들지 못했습니다.")
        except Exception as e:
            last_err = e
            time.sleep(1)  # 잠깐 쉬었다가 새로 시도
        finally:
            if hwp is not None:
                try:
                    hwp.Clear(1)  # 문서를 저장하지 않고 닫기
                except Exception:
                    pass
                try:
                    hwp.Quit()  # 여기서 오류가 나도 무시 (진짜 원인을 가리지 않도록)
                except Exception:
                    pass
                hwp = None
    raise RuntimeError(f"한글 변환 실패: {last_err}")


def convert_to_pdf(path, tmpdir, idx):
    """어떤 파일이든 PDF 경로로 바꿔서 돌려줍니다 (PDF는 그대로)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return path
    out = os.path.join(tmpdir, f"{idx:03d}.pdf")

    if ext in IMG_EXT:  # 이미지는 A4 용지에 맞춰 넣기
        pix = pymupdf.Pixmap(path)
        w, h = pix.width, pix.height
        a4 = (842, 595) if w > h else (595, 842)
        doc = pymupdf.open()
        page = doc.new_page(width=a4[0], height=a4[1])
        page.insert_image(page.rect, filename=path, keep_proportion=True)
        doc.save(out)
        doc.close()
        return out

    # 오피스/한글 파일은 설치된 프로그램을 불러와 PDF로 저장
    import pythoncom
    import win32com.client as wc
    pythoncom.CoInitialize()
    src, dst = os.path.abspath(path), os.path.abspath(out)
    try:
        if ext in (".xls", ".xlsx", ".xlsm"):
            app = wc.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(src, ReadOnly=True)
                wb.ExportAsFixedFormat(0, dst)
                wb.Close(False)
            finally:
                app.Quit()
        elif ext in (".doc", ".docx"):
            app = wc.DispatchEx("Word.Application")
            app.Visible = False
            try:
                d = app.Documents.Open(src, ReadOnly=True)
                d.ExportAsFixedFormat(dst, 17)
                d.Close(False)
            finally:
                app.Quit()
        elif ext in (".ppt", ".pptx"):
            app = wc.DispatchEx("PowerPoint.Application")
            try:
                pres = app.Presentations.Open(src, True, False, False)
                pres.SaveAs(dst, 32)
                pres.Close()
            finally:
                app.Quit()
        elif ext in (".hwp", ".hwpx"):
            hwp_to_pdf(src, dst, ext)
        else:
            raise ValueError("지원하지 않는 형식입니다.")
    finally:
        pythoncom.CoUninitialize()
    if not os.path.exists(dst):
        raise RuntimeError("PDF 변환 결과 파일이 만들어지지 않았습니다.")
    return dst


# ───────────────────────── 파일 목록 위젯 ─────────────────────────
class FileList(ctk.CTkFrame):
    """추가 / 삭제 / 정렬 / 끌어서 순서 바꾸기가 되는 파일 목록

    - exts: 받아들일 확장자 (기본: 모든 지원 형식)
    - unique: True 이면 같은 파일을 두 번 추가하지 않음
    - on_change(kind, old_paths): 목록이 바뀔 때마다 불리는 함수 (조정하기 화면에서 사용)
        kind 는 "add" / "remove" / "clear" / "reorder"
        "reorder" 일 때 False 를 돌려주면 순서 변경을 취소하고 원래대로 되돌립니다.
    - delete_needs_focus: True 이면 목록을 클릭해 둔 상태에서만 Delete 키로 삭제
    """

    def __init__(self, master, filetypes, exts=None, unique=False,
                 on_change=None, delete_needs_focus=False):
        super().__init__(master, fg_color="transparent")
        self.paths = []
        self.filetypes = filetypes
        self.exts = exts or ALL_EXT
        self.unique = unique
        self.on_change = on_change
        self.delete_needs_focus = delete_needs_focus

        # 드래그 상태를 기억하는 변수들
        self.drag_start_y = 0
        self.drag_start_index = -1
        self.dragged_indices = ()
        self.is_dragging = False
        self.order_before = []  # 드래그를 시작하기 직전의 순서 (취소할 때 되돌리기 위해)

        box = ctk.CTkFrame(self, corner_radius=10, fg_color="white",
                           border_width=1, border_color="#D5DAE3")
        box.pack(side="left", fill="both", expand=True)
        # exportselection=False: 다른 입력칸을 눌러도 목록의 선택이 풀리지 않게 함
        self.lb = tk.Listbox(box, selectmode="extended", bd=0, bg="white",
                             highlightthickness=0, activestyle="none",
                             exportselection=False,
                             font=(FONT, 11), selectbackground="#3B8ED0")
        sb = ctk.CTkScrollbar(box, command=self.lb.yview)
        self.lb.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(0, 2), pady=6)
        self.lb.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        # 키보드 Delete 로 삭제 (창 전체에서 받되, 이 화면이 보일 때만 동작)
        top = self.winfo_toplevel()
        top.bind("<Delete>", self.on_delete_key, add="+")
        top.bind("<BackSpace>", self.on_delete_key, add="+")

        # 마우스로 끌어서 순서 바꾸기
        self.lb.bind("<Button-1>", self.on_drag_start)
        self.lb.bind("<B1-Motion>", self.on_drag_motion)
        self.lb.bind("<ButtonRelease-1>", self.on_drag_release)

        side = ctk.CTkFrame(self, fg_color="transparent")
        side.pack(side="left", fill="y", padx=(12, 0))
        for text, cmd in [("＋ 파일 추가", self.add), ("－ 선택 삭제", self.remove),
                          ("전체 비우기", self.clear)]:
            ctk.CTkButton(side, text=text, width=110, command=cmd).pack(pady=4)
        ctk.CTkButton(side, text="이름 ↑ 오름차순", width=110, fg_color="#8A94A6",
                      command=lambda: self.sort_by_name(False)).pack(pady=(14, 4))
        ctk.CTkButton(side, text="이름 ↓ 내림차순", width=110, fg_color="#8A94A6",
                      command=lambda: self.sort_by_name(True)).pack(pady=4)
        ctk.CTkLabel(side, text="※ 목록의 파일을\n끌어서 순서를\n바꿀 수 있습니다.\n(Delete 키로 삭제)",
                     text_color="#667085", justify="left",
                     font=ctk.CTkFont(size=11)).pack(pady=(10, 0))

    def _notify(self, kind, old):
        """목록이 바뀐 것을 화면(페이지)에 알립니다. 순서 변경은 화면이 거절하면 되돌립니다."""
        if self.on_change is None:
            return
        if self.on_change(kind, old) is False and kind == "reorder":
            self.paths = list(old)
            self.refresh()

    def refresh(self, select=()):
        self.lb.delete(0, "end")
        for i, p in enumerate(self.paths, 1):
            self.lb.insert("end", f"  {i:>2}.  {os.path.basename(p)}")
        for i in select:
            self.lb.selection_set(i)

    def add(self):
        files = filedialog.askopenfilenames(filetypes=self.filetypes)
        if not files:
            return
        added, skipped = self.add_paths(files)
        if skipped:
            messagebox.showinfo("안내", f"{added}개를 추가했습니다.\n"
                                        f"지원하지 않는 형식이거나 이미 추가된 파일 {skipped}개는 건너뛰었습니다.")

    def add_paths(self, paths):
        """파일 추가(끌어다 놓기 포함). 추가된 개수와 건너뛴 개수를 돌려줍니다."""
        old = list(self.paths)
        added = skipped = 0
        for p in paths:
            ok = os.path.isfile(p) and os.path.splitext(p)[1].lower() in self.exts
            if ok and self.unique and any(same_path(p, q) for q in self.paths):
                ok = False  # 이미 목록에 있는 파일
            if ok:
                self.paths.append(p)
                added += 1
            else:
                skipped += 1
        self.refresh()
        if added:
            self._notify("add", old)
        return added, skipped

    def remove(self):
        sel = self.lb.curselection()
        if not sel:
            return
        old = list(self.paths)
        first = sel[0]
        for i in reversed(sel):
            del self.paths[i]
        # 삭제 후 그 자리의 다음 파일을 선택해 두면 Delete 를 계속 눌러 연속 삭제할 수 있음
        nxt = min(first, len(self.paths) - 1)
        self.refresh([nxt] if nxt >= 0 else ())
        self.lb.focus_set()
        self._notify("remove", old)

    def on_delete_key(self, event):
        """Delete / BackSpace 키 처리. 글자 입력칸에서 누른 키는 건드리지 않습니다."""
        if not self.winfo_viewable():  # 다른 화면을 보고 있으면 무시
            return
        focus = self.focus_get()
        if focus is not None and focus.winfo_class() in ("Entry", "Text", "TEntry"):
            return
        if self.delete_needs_focus and focus is not self.lb:
            return  # 목록을 클릭해 둔 상태가 아니면 무시 (페이지 작업 중 실수 방지)
        if self.lb.curselection():
            self.remove()
            return "break"

    def clear(self):
        if not self.paths:
            return
        old = list(self.paths)
        self.paths.clear()
        self.refresh()
        self._notify("clear", old)

    def sort_by_name(self, reverse=False):
        """파일 이름 기준 정렬 (숫자는 크기대로: 2번 < 10번)"""
        if not self.paths:
            return
        old = list(self.paths)
        self.paths.sort(key=natural_key, reverse=reverse)
        self.refresh()
        if self.paths != old:
            self._notify("reorder", old)

    # ── 끌어서 순서 바꾸기 ──
    def on_drag_start(self, event):
        self.lb.focus_set()  # 키보드(Delete) 입력을 받을 수 있도록 포커스를 목록에 둠
        self.order_before = list(self.paths)
        self.drag_start_y = event.y
        self.is_dragging = False
        index = self.lb.nearest(event.y)

        # 항목이 없는 빈 곳을 눌렀으면 선택 해제
        bbox = self.lb.bbox(index)
        if bbox and (event.y < bbox[1] or event.y > bbox[1] + bbox[3]):
            self.lb.selection_clear(0, "end")
            self.dragged_indices = ()
            return

        # Shift / Ctrl 로 여러 개 고르는 중이면 기본 동작에 맡김
        if event.state & 0x0001 or event.state & 0x0004:
            self.after(10, self._record_drag_start, event.y)
            return

        # 이미 여러 개가 선택된 상태에서 그중 하나를 누르면, 선택을 유지한 채 함께 끌기
        if len(self.lb.curselection()) > 1 and index in self.lb.curselection():
            self.drag_start_index = index
            self.dragged_indices = self.lb.curselection()
            return "break"
        self.after(10, self._record_drag_start, event.y)

    def _record_drag_start(self, y):
        self.drag_start_index = self.lb.nearest(y)
        self.dragged_indices = self.lb.curselection()

    def on_drag_motion(self, event):
        if abs(event.y - self.drag_start_y) > 3:
            self.is_dragging = True
        if not self.is_dragging or not self.dragged_indices:
            return "break"

        target = self.lb.nearest(event.y)
        if target < 0 or target >= self.lb.size() or target == self.drag_start_index:
            return "break"

        step = 1 if target > self.drag_start_index else -1
        new_indices = [i + step for i in self.dragged_indices]
        if min(new_indices) < 0 or max(new_indices) >= len(self.paths):
            return "break"

        # 선택한 항목을 한 칸씩 위/아래로 이동
        if step == 1:
            for i in reversed(self.dragged_indices):
                self.paths[i], self.paths[i + 1] = self.paths[i + 1], self.paths[i]
        else:
            for i in self.dragged_indices:
                self.paths[i], self.paths[i - 1] = self.paths[i - 1], self.paths[i]

        self.refresh(new_indices)
        self.dragged_indices = new_indices
        self.drag_start_index += step
        return "break"

    def on_drag_release(self, event):
        # 끌어서 순서가 실제로 바뀌었다면 화면(페이지)에 알림 (필요하면 경고창 후 되돌림)
        if self.is_dragging and self.paths != self.order_before:
            before = self.order_before
            self.is_dragging = False
            self._notify("reorder", before)
            return
        # 끌지 않고 그냥 클릭만 한 경우: 그 항목 하나만 선택
        if not self.is_dragging and not (event.state & 0x0001 or event.state & 0x0004):
            index = self.lb.nearest(event.y)
            bbox = self.lb.bbox(index)
            if bbox and bbox[1] <= event.y <= bbox[1] + bbox[3]:
                self.lb.selection_clear(0, "end")
                self.lb.selection_set(index)
                self._record_drag_start(event.y)
        self.is_dragging = False


# ───────────────────────── 화면(페이지)들 ─────────────────────────
class BasePage(ctk.CTkFrame):
    def __init__(self, master, app, title, desc):
        super().__init__(master, fg_color=BG, corner_radius=0)
        self.app = app
        ctk.CTkLabel(self, text=title, anchor="w",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(fill="x", padx=28, pady=(24, 0))
        ctk.CTkLabel(self, text=desc, anchor="w", justify="left", text_color="#667085").pack(fill="x", padx=28, pady=(2, 14))

    def pick_pdf(self):
        return filedialog.askopenfilename(filetypes=[("PDF 파일", "*.pdf")])

    # ── 드래그앤드롭 ──
    def enable_drop(self, widget=None):
        """이 화면의 모든 위젯을 '파일 놓는 곳'으로 등록합니다 (화면 구성이 끝난 뒤 호출)."""
        widget = widget or self
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass
        for child in widget.winfo_children():
            self.enable_drop(child)

    def _on_drop(self, event):
        paths = list(self.tk.splitlist(event.data))
        self.handle_drop(paths)
        return event.action

    def handle_drop(self, paths):
        """기본 동작: 끌어다 놓은 파일 중 첫 번째 PDF를 불러옵니다."""
        pdfs = [p for p in paths if p.lower().endswith(".pdf") and os.path.isfile(p)]
        if not pdfs:
            return messagebox.showinfo("안내", "이 화면에는 PDF 파일을 끌어다 놓아 주세요.")
        self.load(pdfs[0])


class MergePage(BasePage):
    def __init__(self, master, app):
        super().__init__(master, app, "PDF로 변환하기",
                         "PDF, 한글, 파워포인트, 엑셀, 워드, 이미지 파일을 목록에 담고 순서를 정해 하나의 PDF로 저장합니다.\n"
                         "※ 한글 파일 추가 시 접근 허용 팝업이 뜨면 허용을 눌러주세요.")
        types = [("지원 파일", "*.pdf *.hwp *.hwpx *.xls *.xlsx *.xlsm *.doc *.docx *.ppt *.pptx *.jpg *.jpeg *.png *.bmp"),
                 ("모든 파일", "*.*")]
        self.fl = FileList(self, types)
        self.fl.pack(fill="both", expand=True, padx=28)

        # 쪽 번호 옵션 (기본값: 체크 해제 = 번호 없음)
        self.num_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(self, text="쪽 번호 넣기  (아래쪽 가운데에 '1/전체쪽수' 형태로 표시)",
                        variable=self.num_var).pack(pady=(14, 0))

        ctk.CTkButton(self, text="목록 순서대로 PDF 저장", height=44,
                      font=ctk.CTkFont(size=15, weight="bold"),
                      command=self.save).pack(pady=18)
        self.enable_drop()

    def handle_drop(self, paths):
        added, skipped = self.fl.add_paths(paths)
        if skipped:
            messagebox.showinfo("안내", f"{added}개를 추가했습니다.\n"
                                        f"지원하지 않는 형식이거나 폴더인 {skipped}개는 건너뛰었습니다.")

    def save(self):
        paths = list(self.fl.paths)
        if not paths:
            return messagebox.showinfo("안내", "먼저 파일을 추가해 주세요.")
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                           initialfile="합친문서.pdf")
        if not out:
            return
        numbering = self.num_var.get()  # 별도 스레드에서 화면 값을 읽지 않도록 미리 가져옴

        def work():
            tmp = tempfile.mkdtemp()
            failed = []
            try:
                merged = pymupdf.open()
                for i, p in enumerate(paths):
                    self.app.report(i, len(paths), f"{os.path.basename(p)} 처리 중", "개")
                    try:
                        pdf = convert_to_pdf(p, tmp, i)
                        with open_pdf(pdf) as d:
                            merged.insert_pdf(d)
                    except Exception as e:
                        failed.append(f"{os.path.basename(p)} ({e})")
                self.app.busy("PDF 저장 중...")
                if merged.page_count:
                    if numbering:
                        stamp_page_numbers(merged)  # 합쳐진 결과의 순서대로 1/전체쪽수
                    merged.save(out, garbage=3, deflate=True)
                merged.close()
            finally:
                shutil.rmtree(tmp, ignore_errors=True)
            return failed

        def done(failed, err):
            if err:
                return messagebox.showerror("오류", str(err))
            msg = "PDF 저장이 완료되었습니다."
            if failed:
                msg += "\n\n처리하지 못한 파일:\n" + "\n".join(failed)
            messagebox.showinfo("완료", msg)

        self.app.run_bg(work, done, "PDF를 만드는 중...")


class MaskPage(BasePage):
    """PDF 마스킹: 첫 쪽 미리보기에서 영역을 끌어 정하면, 모든 쪽의 같은 위치를 검은색으로 '삭제'합니다.

    ★ 핵심: 검은 네모를 위에 덮는 방식이 아니라 PyMuPDF 의 '교정(redaction)' 기능을 씁니다.
      - add_redact_annot  : 지울 영역과 채울 색(검정)을 표시
      - apply_redactions  : 그 영역 안의 글자·그림 픽셀·선 그림을 파일에서 실제로 제거하고 검정으로 채움
      -> 저장된 PDF에는 가려진 내용이 남아 있지 않아서, 검은 칠을 지워도 복구할 수 없습니다.

    영역은 '쪽 크기에 대한 비율(0~1)'로 기억합니다. 그래서 쪽마다 크기가 달라도(가로/세로 혼합 등)
    첫 쪽에서 고른 것과 같은 상대 위치가 가려집니다."""
    MIN_DRAG_PX = 6          # 이보다 작게 끌면 실수로 보고 영역으로 만들지 않음
    PREVIEW_DPI = 150        # 미리보기 기준 그림 해상도 (화면 크기에 맞춰 줄여서 보여줌)
    PREVIEW_MAX_SIDE = 2200  # 아주 큰 도면도 감당할 수 있게 그림의 긴 변 최대 픽셀 제한

    def __init__(self, master, app):
        super().__init__(master, app, "PDF 마스킹",
                         "첫 쪽 미리보기에서 가릴 곳을 마우스로 끌어 네모를 그리면, 모든 쪽의 같은 위치가 검은색으로 바뀝니다. (여러 곳 지정 가능)\n"
                         "※ 위에 덮어 가리는 것이 아니라 그 자리의 글자·그림 자체를 삭제하므로 복구할 수 없습니다. 원본은 그대로 남습니다.")
        self.src = None
        self.base_img = None     # 첫 쪽을 그림으로 만들어 둔 것 (화면 크기에 맞춰 줄여서 표시)
        self.rects = []          # 지정한 영역들: (왼쪽, 위, 오른쪽, 아래) 를 쪽 크기 대비 비율(0~1)로 기억
        self.view = None         # 화면에 그림이 놓인 위치와 크기: (x, y, 가로, 세로)
        self.tk_img = None       # 화면에 표시 중인 그림 (파이썬이 지워버리지 않게 붙들어 두기)
        self.drag_start = None
        self.drag_id = None
        self._redraw_job = None

        # 아래쪽 버튼 (먼저 배치해 두어야 창이 작아져도 가려지지 않음)
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(side="bottom", pady=14)
        ctk.CTkButton(bottom, text="마지막 영역 취소", width=130, height=40, fg_color="#8A94A6",
                      command=self.undo_rect).pack(side="left", padx=6)
        ctk.CTkButton(bottom, text="영역 모두 지우기", width=130, height=40, fg_color="#8A94A6",
                      command=self.clear_rects).pack(side="left", padx=6)
        ctk.CTkButton(bottom, text="모든 쪽에 마스킹하여 저장", height=40, width=220,
                      fg_color="#1F2A44", hover_color="#33415F",
                      font=ctk.CTkFont(size=14, weight="bold"),
                      command=self.run).pack(side="left", padx=(20, 6))

        r1 = ctk.CTkFrame(self, fg_color="transparent")
        r1.pack(fill="x", padx=28)
        ctk.CTkButton(r1, text="PDF 파일 선택", width=130, command=self.load).pack(side="left")
        self.name = ctk.CTkLabel(r1, text="선택된 파일 없음", text_color="#667085")
        self.name.pack(side="left", padx=12)
        self.info = ctk.CTkLabel(r1, text="", text_color="#667085")
        self.info.pack(side="right")

        box = ctk.CTkFrame(self, fg_color="white", corner_radius=10)
        box.pack(fill="both", expand=True, padx=28, pady=(8, 0))
        self.cv = tk.Canvas(box, bg="#D5DAE3", highlightthickness=0, cursor="crosshair")
        self.cv.pack(fill="both", expand=True, padx=8, pady=8)
        self.cv.bind("<ButtonPress-1>", self._press)
        self.cv.bind("<B1-Motion>", self._motion)
        self.cv.bind("<ButtonRelease-1>", self._release)
        self.cv.bind("<Configure>", self._on_canvas_resize)

        self._show_preview()
        self.enable_drop()

    # ───────── 파일 불러오기 ─────────
    def load(self, p=None):
        p = p or self.pick_pdf()
        if not p:
            return
        try:
            with open_pdf(p) as doc:
                if doc.page_count == 0:
                    raise ValueError("페이지가 없는 PDF입니다.")
                page = doc[0]
                # page.rect 는 회전까지 반영된 '눈에 보이는' 쪽 크기라서 미리보기와 정확히 맞습니다.
                zoom = min(self.PREVIEW_DPI / 72.0, self.PREVIEW_MAX_SIDE / max(page.rect.width, page.rect.height))
                pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                pages = doc.page_count
        except Exception as e:
            return messagebox.showerror("오류", f"PDF를 열지 못했습니다.\n{e}")
        self.src = p
        self.base_img = img
        self.rects = []
        self.name.configure(text=f"{os.path.basename(p)}  (총 {pages}쪽)")
        self._show_preview()

    # ───────── 화면 그리기 ─────────
    def _on_canvas_resize(self, event):
        # 창 크기를 줄였다 늘렸다 할 때 계속 다시 그리지 않도록 잠깐 기다렸다가 한 번만 그립니다.
        if self._redraw_job:
            self.after_cancel(self._redraw_job)
        self._redraw_job = self.after(80, self._show_preview)

    def _show_preview(self):
        self._redraw_job = None
        self.cv.delete("all")
        self.drag_id = None
        cw, ch = self.cv.winfo_width(), self.cv.winfo_height()
        if self.base_img is None or cw < 20 or ch < 20:
            self.view = None
            self.info.configure(text="")
            if cw >= 20 and ch >= 20:
                self.cv.create_text(cw // 2, ch // 2, fill="#667085", font=(FONT, 13),
                                    text="PDF 파일을 선택하거나 이 창 위로 끌어다 놓으세요.\n첫 쪽이 여기에 나타납니다.")
            return
        w, h = self.base_img.size
        s = min(cw / w, ch / h)
        iw, ih = max(1, int(w * s)), max(1, int(h * s))
        ox, oy = (cw - iw) // 2, (ch - ih) // 2
        self.view = (ox, oy, iw, ih)
        self.tk_img = ImageTk.PhotoImage(self.base_img.resize((iw, ih), Image.LANCZOS))
        self.cv.create_image(ox, oy, anchor="nw", image=self.tk_img)
        self.cv.create_rectangle(ox, oy, ox + iw, oy + ih, outline="#8A94A6")
        for n, (a, b, c, d) in enumerate(self.rects, 1):
            x0, y0, x1, y1 = ox + a * iw, oy + b * ih, ox + c * iw, oy + d * ih
            self.cv.create_rectangle(x0, y0, x1, y1, fill="black", stipple="gray50",
                                     outline="#E5322D", width=2)
            self.cv.create_text(x0 + 4, y0 + 2, text=str(n), anchor="nw",
                                fill="#E5322D", font=(FONT, 11, "bold"))
        self.info.configure(text=f"지정한 영역 {len(self.rects)}개")

    # ───────── 마우스로 영역 그리기 ─────────
    def _clamp(self, x, y):
        """마우스 위치를 미리보기 그림 안쪽으로 제한합니다."""
        ox, oy, iw, ih = self.view
        return min(max(x, ox), ox + iw), min(max(y, oy), oy + ih)

    def _press(self, event):
        if not self.view:
            return
        self.drag_start = self._clamp(event.x, event.y)
        if self.drag_id:
            self.cv.delete(self.drag_id)
            self.drag_id = None

    def _motion(self, event):
        if not self.drag_start or not self.view:
            return
        x, y = self._clamp(event.x, event.y)
        x0, y0 = self.drag_start
        if self.drag_id:
            self.cv.coords(self.drag_id, x0, y0, x, y)
        else:
            self.drag_id = self.cv.create_rectangle(x0, y0, x, y, outline="#E5322D",
                                                    width=2, dash=(4, 2))

    def _release(self, event):
        if not self.drag_start or not self.view:
            return
        x, y = self._clamp(event.x, event.y)
        x0, y0 = self.drag_start
        self.drag_start = None
        if self.drag_id:
            self.cv.delete(self.drag_id)
            self.drag_id = None
        if abs(x - x0) < self.MIN_DRAG_PX or abs(y - y0) < self.MIN_DRAG_PX:
            return  # 너무 작게 끌었으면 무시 (그냥 클릭한 경우)
        ox, oy, iw, ih = self.view
        a, c = sorted(((x0 - ox) / iw, (x - ox) / iw))
        b, d = sorted(((y0 - oy) / ih, (y - oy) / ih))
        self.rects.append((a, b, c, d))
        self._show_preview()

    def undo_rect(self):
        if self.rects:
            self.rects.pop()
            self._show_preview()

    def clear_rects(self):
        if self.rects:
            self.rects = []
            self._show_preview()

    # ───────── 마스킹 (진짜 삭제) ─────────
    @staticmethod
    def mask_pdf(src, dst, rects, report, busy):
        """모든 쪽의 같은 비율 위치를 '지우고' 검은색으로 채워 dst 에 저장합니다. (별도 스레드에서 실행)"""
        img_mode = getattr(pymupdf, "PDF_REDACT_IMAGE_PIXELS", 2)                  # 그림은 겹친 부분의 픽셀만 지움
        gfx_mode = getattr(pymupdf, "PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED", 2)    # 선·도형도 닿으면 지움
        doc = open_pdf(src)
        try:
            n = doc.page_count
            for i in range(n):
                report(i, n, "마스킹 중")
                page = doc[i]
                r = page.rect  # 화면에 보이는 쪽 크기 (회전 반영)
                targets = []
                for a, b, c, d in rects:
                    visible = pymupdf.Rect(r.x0 + a * r.width, r.y0 + b * r.height,
                                           r.x0 + c * r.width, r.y0 + d * r.height)
                    # 회전된 쪽은 PDF 내부 좌표가 다르므로 원래 좌표로 되돌려서 적용
                    targets.append(visible * page.derotation_matrix)

                # 가릴 곳에 걸친 메모·스탬프·입력칸 같은 주석에도 내용이 숨어 있을 수 있으니 함께 삭제
                try:
                    annot = page.first_annot
                    while annot:
                        if any(annot.rect.intersects(t) for t in targets):
                            annot = page.delete_annot(annot)
                        else:
                            annot = annot.next
                except Exception:
                    pass

                for t in targets:
                    page.add_redact_annot(t, fill=(0, 0, 0))  # fill=(0,0,0) 이 검은색
                try:
                    page.apply_redactions(images=img_mode, graphics=gfx_mode, text=0)
                except TypeError:  # 옛 PyMuPDF 버전은 graphics/text 옵션이 없음
                    page.apply_redactions(images=img_mode)
            report(n, n, "마스킹 중")
            busy("저장 중...")
            # garbage=4: 지워진 내용이 파일 안에 찌꺼기로 남지 않도록 쓰지 않는 데이터를 모두 정리
            doc.save(dst, garbage=4, deflate=True, clean=True)
            return n
        finally:
            doc.close()

    def run(self):
        if not self.src:
            return messagebox.showinfo("안내", "먼저 PDF 파일을 선택해 주세요.")
        if not self.rects:
            return messagebox.showinfo("안내", "가릴 영역이 없습니다.\n미리보기 위에서 마우스를 끌어 영역을 먼저 지정해 주세요.")
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                           initialfile=os.path.splitext(os.path.basename(self.src))[0] + "_마스킹.pdf")
        if not out:
            return
        if same_path(out, self.src):
            return messagebox.showwarning("안내", "원본과 다른 이름으로 저장해 주세요.")
        src, rects = self.src, list(self.rects)

        def work():
            return self.mask_pdf(src, out, rects, self.app.report, self.app.busy)

        self.app.run_bg(work, lambda n, e: messagebox.showerror("오류", str(e)) if e
                        else messagebox.showinfo("완료", f"{n}쪽 모두에 마스킹을 적용했습니다.\n"
                                                        "가려진 부분의 내용은 파일에서 삭제되었습니다."),
                        "마스킹 중...")


class ExtractPage(BasePage):
    """합치기 · 분리하기 · 회전하기
    왼쪽: PDF 파일 목록 / 오른쪽: 모든 페이지 미리보기(끌어서 순서 변경, 삭제, 회전, 크기 조절)
    self.pages 의 각 항목은 {"path": 원본 파일, "no": 원본의 몇 번째 쪽(0부터),
                           "cell": 미리보기 칸, "cap": 설명 글자, "var": 체크 상태,
                           "lbl": 그림이 들어가는 글자판, "img": 지금 보이는 그림,
                           "dims": 가장 크게 볼 때의 그림 크기, "rot": 시계 방향 회전 각도(0/90/180/270)} 입니다.
    self.pages 의 순서가 곧 저장될 페이지 순서입니다."""
    # 미리보기 그림은 가장 크게 볼 때(MAX_W) 기준으로 한 번만 만들어 두고,
    # 슬라이더로 크기를 바꿀 때는 이 그림을 줄이거나 늘려 보여줍니다. (다시 만들지 않아 빠름)
    RENDER_W, RENDER_H = 180, 252
    MIN_W, MAX_W, DEFAULT_W = 60, 180, 100
    C_NORMAL, C_DRAG, C_TARGET = "#F6F7FA", "#BBD7F2", "#FFE7A3"

    def __init__(self, master, app):
        super().__init__(master, app, "PDF 합치기 · 분리하기 · 회전하기",
                         "여러 PDF를 목록에 담고, 페이지 미리보기에서 순서를 바꾸거나 지우거나 회전한 뒤 하나로 합치거나 일부만 저장·낱장 분리합니다.\n"
                         "PDF 파일을 끌어다 놓아도 목록에 추가됩니다.")
        self.pages = []
        self.thumb_pil = {}     # (파일, 쪽번호) -> 회전하지 않은 미리보기 그림 (회전할 때 이 그림을 바로 돌려 씀)
        self.page_counts = {}   # 파일 -> 전체 쪽 수
        self.file_colors = {}   # 파일 -> 색깔 띠
        self.color_no = 0
        self.thumb_w = self.DEFAULT_W   # 지금 보여주는 미리보기 칸의 가로 크기
        self.grid_px = 0                # 미리보기 영역의 현재 가로 폭 (실제 화면 픽셀, 칸 개수 계산용)
        self.cols = 4
        self.drag = None
        self._relayout_job = None
        self._size_job = None

        # 아래쪽 저장 버튼 (먼저 배치해 두어야 창이 작아져도 가려지지 않음)
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(side="bottom", pady=14)
        ctk.CTkButton(bottom, text="전체 페이지를 하나의 PDF로 저장", height=42,
                      command=lambda: self.save_pages(False)).pack(side="left", padx=6)
        ctk.CTkButton(bottom, text="선택한 페이지만 하나의 PDF로 저장", height=42,
                      command=lambda: self.save_pages(True)).pack(side="left", padx=6)
        ctk.CTkButton(bottom, text="선택한 페이지 낱장 분리", height=42, fg_color="#2F9E7A",
                      hover_color="#25805F", command=self.split_checked).pack(side="left", padx=6)

        # 쪽 번호 옵션 (저장 버튼 바로 위, 기본값: 체크 해제 = 번호 없음)
        self.num_var = tk.BooleanVar(value=False)
        opt = ctk.CTkFrame(self, fg_color="transparent")
        opt.pack(side="bottom", pady=(6, 0))
        ctk.CTkCheckBox(opt, text="쪽 번호 넣기  (아래쪽 가운데에 '1/전체쪽수' 형태로 표시 · 낱장 분리 시에는 미리보기 순서 번호)",
                        variable=self.num_var).pack()

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=28)

        # ── 왼쪽: 파일 목록 ──
        left = ctk.CTkFrame(body, fg_color="transparent", width=330)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        ctk.CTkLabel(left, text="① PDF 파일 목록", anchor="w",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x", pady=(0, 6))
        self.fl = FileList(left, [("PDF 파일", "*.pdf")], exts={".pdf"}, unique=True,
                           on_change=self.on_files_changed, delete_needs_focus=True)
        self.fl.pack(fill="both", expand=True)

        # ── 오른쪽: 페이지 미리보기 ──
        right = ctk.CTkFrame(body, fg_color="transparent")
        right.pack(side="left", fill="both", expand=True, padx=(14, 0))
        ctk.CTkLabel(right, text="② 페이지 미리보기  (끌어서 순서 변경 · 체크한 쪽은 함께 이동·회전)", anchor="w",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(fill="x")

        r1 = ctk.CTkFrame(right, fg_color="transparent")
        r1.pack(fill="x", pady=(6, 0))
        self.entry = ctk.CTkEntry(r1, width=150, placeholder_text="예: 1-3,5,8")
        self.entry.pack(side="left")
        ctk.CTkButton(r1, text="범위 선택", width=80, command=self.apply_range).pack(side="left", padx=6)
        ctk.CTkButton(r1, text="전체 선택", width=80, fg_color="#8A94A6",
                      command=lambda: self.set_all(True)).pack(side="left", padx=2)
        ctk.CTkButton(r1, text="전체 해제", width=80, fg_color="#8A94A6",
                      command=lambda: self.set_all(False)).pack(side="left", padx=2)

        r2 = ctk.CTkFrame(right, fg_color="transparent")
        r2.pack(fill="x", pady=6)
        ctk.CTkButton(r2, text="◀ 앞으로", width=80,
                      command=lambda: self.move_checked(-1)).pack(side="left")
        ctk.CTkButton(r2, text="뒤로 ▶", width=80,
                      command=lambda: self.move_checked(1)).pack(side="left", padx=6)
        ctk.CTkButton(r2, text="선택 삭제", width=80, fg_color="#D9534F", hover_color="#B8413D",
                      command=self.delete_checked).pack(side="left", padx=2)
        ctk.CTkButton(r2, text="순서·삭제·회전 초기화", width=140, fg_color="#8A94A6",
                      command=self.ask_reset).pack(side="left", padx=6)

        # 회전 버튼 (체크한 페이지를 돌립니다. 미리보기에 바로 반영되고 저장할 때도 그대로 적용됩니다)
        r2b = ctk.CTkFrame(right, fg_color="transparent")
        r2b.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(r2b, text="선택한 쪽 회전").pack(side="left", padx=(0, 8))
        ctk.CTkButton(r2b, text="↺ 왼쪽 90°", width=100, fg_color="#2F9E7A", hover_color="#25805F",
                      command=lambda: self.rotate_checked(-90)).pack(side="left")
        ctk.CTkButton(r2b, text="↻ 오른쪽 90°", width=100, fg_color="#2F9E7A", hover_color="#25805F",
                      command=lambda: self.rotate_checked(90)).pack(side="left", padx=6)
        ctk.CTkButton(r2b, text="180° 회전", width=100, fg_color="#2F9E7A", hover_color="#25805F",
                      command=lambda: self.rotate_checked(180)).pack(side="left")

        # 미리보기 그림 크기 조절 슬라이더
        r3 = ctk.CTkFrame(right, fg_color="transparent")
        r3.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(r3, text="미리보기 크기").pack(side="left")
        ctk.CTkLabel(r3, text="작게", text_color="#667085",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(8, 2))
        self.size_slider = ctk.CTkSlider(r3, from_=self.MIN_W, to=self.MAX_W, width=170,
                                         number_of_steps=(self.MAX_W - self.MIN_W) // 10,
                                         command=self.on_size_slider)
        self.size_slider.set(self.DEFAULT_W)
        self.size_slider.pack(side="left")
        ctk.CTkLabel(r3, text="크게", text_color="#667085",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(2, 8))
        self.size_label = ctk.CTkLabel(r3, text="100%", text_color="#667085", width=44, anchor="w")
        self.size_label.pack(side="left")

        self.info = ctk.CTkLabel(right, text="", anchor="w", text_color="#667085")
        self.info.pack(fill="x")

        self.grid_frame = ctk.CTkScrollableFrame(right, fg_color="white", corner_radius=10)
        self.grid_frame.pack(fill="both", expand=True, pady=(4, 0))
        # 창 크기에 맞춰 한 줄에 보이는 칸 수를 자동으로 조절합니다.
        # ★ 실제로 그림이 보이는 '캔버스'의 폭을 기준으로 삼습니다. (바깥 틀의 폭은 실제 보이는 폭과 다를 수 있음)
        try:
            self.grid_frame._parent_canvas.bind("<Configure>", self._on_resize, add="+")
        except Exception:
            self.grid_frame.bind("<Configure>", self._on_resize, add="+")

        self.update_info()
        self.enable_drop()

    # ───────── 끌어다 놓기 / 파일 목록 변화 ─────────
    def handle_drop(self, paths):
        added, skipped = self.fl.add_paths(paths)
        if skipped:
            messagebox.showinfo("안내", f"{added}개를 추가했습니다.\n"
                                        f"PDF가 아니거나 이미 추가된 파일, 폴더 {skipped}개는 건너뛰었습니다.")

    def on_files_changed(self, kind, old):
        """파일 목록이 바뀔 때마다 FileList 가 불러주는 함수. (kind: add / remove / clear / reorder)"""
        if kind == "add":
            self.load_new([p for p in self.fl.paths if p not in old])
        elif kind in ("remove", "clear"):
            self.drop_files(set(old) - set(self.fl.paths))
        elif kind == "reorder":
            if self.is_edited(old):
                ok = messagebox.askokcancel(
                    "페이지 순서 초기화 안내",
                    "파일 목록의 순서를 바꾸면, 미리보기에서 직접 조정한 페이지 순서와\n"
                    "지우거나 회전해 둔 내용이 모두 초기화됩니다. (페이지가 목록 순서대로 다시 배열됩니다.)\n\n"
                    "계속할까요?")
                if not ok:
                    return False  # 목록 순서를 원래대로 되돌림
            self.reset_pages()
        return True

    def is_edited(self, order):
        """페이지 순서를 직접 바꿨거나, 페이지를 지웠거나, 회전했는지 확인합니다.
        (파일 목록이 order 순서일 때의 '기본 배열'과 지금 배열을 비교)"""
        default = [(p, i) for p in order for i in range(self.page_counts.get(p, 0))]
        current = [(e["path"], e["no"]) for e in self.pages]
        return default != current or any(e["rot"] for e in self.pages)

    # ───────── 미리보기 만들기 ─────────
    def _render_thumbs(self, p):
        """파일의 모든 쪽을 작은 그림으로 만들어 보관합니다."""
        if p in self.page_counts:
            return
        with open_pdf(p) as doc:
            n = doc.page_count
            if n == 0:
                raise ValueError("페이지가 없는 PDF입니다.")
            for i in range(n):
                page = doc[i]
                s = min(self.RENDER_W / page.rect.width, self.RENDER_H / page.rect.height)
                pix = page.get_pixmap(matrix=pymupdf.Matrix(s, s), alpha=False)
                self.thumb_pil[(p, i)] = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                if i % 5 == 4:
                    self.app.status.configure(
                        text=f"미리보기 만드는 중...  {os.path.basename(p)}  {i + 1}/{n}쪽")
                    self.update_idletasks()
            self.page_counts[p] = n
        if p not in self.file_colors:
            self.file_colors[p] = FILE_COLORS[self.color_no % len(FILE_COLORS)]
            self.color_no += 1

    def _forget_file(self, p):
        self.page_counts.pop(p, None)
        self.file_colors.pop(p, None)
        for k in [k for k in self.thumb_pil if k[0] == p]:
            del self.thumb_pil[k]

    def _fit_image(self, key, rot):
        """보관해 둔 그림을 시계 방향으로 rot도 돌려서 돌려줍니다. (rot: 0/90/180/270)
        90·270도로 돌리면 가로세로가 바뀌므로, 칸 크기가 들쭉날쭉해지지 않게
        원래 그림 상자(RENDER_W x RENDER_H) 안에 들어가도록 줄입니다."""
        img = self.thumb_pil[key]
        if not rot:
            return img
        img = img.transpose(ROTATE_OPS[rot])  # 90도 단위 회전은 아주 빨라서 애니메이션이 필요 없음
        s = min(self.RENDER_W / img.width, self.RENDER_H / img.height, 1.0)
        if s < 1.0:
            img = img.resize((max(1, round(img.width * s)), max(1, round(img.height * s))), Image.BILINEAR)
        return img

    def _ctk_image(self, img):
        """PIL 그림을 지금 슬라이더 크기에 맞는 화면용 그림으로 만듭니다."""
        f = self.thumb_w / self.RENDER_W
        return ctk.CTkImage(light_image=img,
                            size=(max(1, int(img.width * f)), max(1, int(img.height * f))))

    def _make_cell(self, path, no):
        """페이지 한 장의 미리보기 칸을 만듭니다 (그림 + 설명 + 선택 체크박스)."""
        img = self._fit_image((path, no), 0)
        cimg = self._ctk_image(img)
        cell = ctk.CTkFrame(self.grid_frame, fg_color=self.C_NORMAL, corner_radius=8)
        bar = ctk.CTkFrame(cell, height=5, fg_color=self.file_colors.get(path, "#8A94A6"), corner_radius=2)
        bar.pack(fill="x", padx=6, pady=(6, 2))
        lbl = ctk.CTkLabel(cell, image=cimg, text="")
        lbl.pack(padx=6, pady=2)
        cap = ctk.CTkLabel(cell, text="", font=ctk.CTkFont(size=11), text_color="#667085",
                           wraplength=self.thumb_w + 10)
        cap.pack(padx=4)
        var = tk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(cell, text="선택", variable=var, width=60, command=self.update_info)
        chk.pack(pady=(2, 6))
        entry = {"path": path, "no": no, "cell": cell, "cap": cap, "var": var,
                 "lbl": lbl, "img": cimg, "dims": (img.width, img.height), "rot": 0}
        for w in (cell, bar, lbl, cap):  # 체크박스를 뺀 나머지를 눌러 끌 수 있게 함
            w.bind("<ButtonPress-1>", lambda e, en=entry: self._drag_press(e, en), add="+")
            w.bind("<B1-Motion>", self._drag_move, add="+")
            w.bind("<ButtonRelease-1>", self._drag_release, add="+")
        self.enable_drop(cell)  # 새로 만든 칸에도 파일 끌어다 놓기 허용
        return entry

    def _set_caption(self, i, en):
        """칸 아래 설명 글자를 씁니다. 예) '3번 · 문서 2쪽 · ↻90°' (회전한 쪽만 각도 표시)"""
        name = os.path.splitext(os.path.basename(en["path"]))[0]
        if len(name) > 10:
            name = name[:10] + "…"
        text = f"{i + 1}번 · {name} {en['no'] + 1}쪽"
        if en["rot"]:
            text += f" · ↻{en['rot']}°"
        en["cap"].configure(text=text)

    def _place(self):
        """현재 self.cols 에 맞춰 칸을 격자에 놓고, 번호 글자를 새로 씁니다."""
        old_conf = getattr(self, "_conf_cols", 0)
        for c in range(max(old_conf, self.cols)):
            # 모든 칸의 폭을 가장 넓은 칸에 맞춰 똑같이 (그래야 폭 계산이 정확해짐)
            self.grid_frame.grid_columnconfigure(c, uniform="pg" if c < self.cols else "", weight=0)
        self._conf_cols = self.cols
        for i, en in enumerate(self.pages):
            en["cell"].grid(row=i // self.cols, column=i % self.cols, padx=6, pady=6)
            self._set_caption(i, en)

    def _measure_cols(self):
        """★ 근본 수정: 추측하지 않고 '실제로 잰 값'으로 한 줄에 들어갈 칸 수를 구합니다.
        - 칸 하나가 실제로 차지하는 폭(가장 넓은 칸 기준)
        - 실제로 눈에 보이는 영역의 폭 (창 바깥으로 밀려 잘린 부분은 제외)"""
        try:
            self.update_idletasks()
            cv = self.grid_frame._parent_canvas
            if cv.winfo_width() < 100 or not self.pages:
                return self.cols
            top = self.winfo_toplevel()
            left = cv.winfo_rootx()
            right = min(left + cv.winfo_width(), top.winfo_rootx() + top.winfo_width())
            sc = self._scaling()
            avail = right - left - int(10 * sc)
            cell_w = max(e["cell"].winfo_reqwidth() for e in self.pages)
            slot = cell_w + int(12 * sc) + 2     # 칸 + 좌우 여백(6씩)
            return max(1, int(avail // slot))
        except Exception:
            return self.cols

    def relayout(self):
        """self.pages 의 순서대로 칸을 격자에 다시 배치하고 번호를 새로 매깁니다."""
        self._relayout_job = None
        self.cols = self._calc_cols()   # 1차: 대략 계산해서 배치
        self._place()
        real = self._measure_cols()     # 2차: 실제로 재서 맞지 않으면 다시 배치
        if real != self.cols:
            self.cols = real
            self._place()
        self.update_info()

    # ───────── 칸 개수 / 미리보기 크기 조절 ─────────
    def _scaling(self):
        """윈도우 '디스플레이 배율'(125%, 150% 등)을 돌려줍니다. 배율이 없으면 1.0"""
        try:
            return float(self.grid_frame._get_widget_scaling())
        except Exception:
            return 1.0

    def _calc_cols(self):
        """미리보기 영역 폭과 그림 크기에 맞는 한 줄 칸 수.
        ★ 핵심 수정: 화면에서 잰 폭(실제 픽셀)을 배율로 나눠서, 칸 크기(논리 크기)와 같은 단위로 맞춥니다.
          (예전에는 배율을 고려하지 않아 '실제로 보이는 칸 수'보다 많이 계산되어 오른쪽이 잘렸습니다.)"""
        if self.grid_px < 100:
            return self.cols
        avail = self.grid_px / self._scaling()          # 보이는 폭 (논리 크기)
        cell_w = max(self.thumb_w + 18, 90)              # 칸 하나의 폭 (그림 + 안쪽 여백, 체크박스 최소 폭 고려)
        slot = cell_w + 12 + 16                          # 칸 사이 여백(좌우 6씩) + 안전 여유 (1차 대략 계산용)
        return max(1, int((avail - 16) // slot))         # 16: 좌우 가장자리 여유

    def _on_resize(self, event):
        if event.width < 100:
            return
        self.grid_px = event.width
        # 창 크기가 바뀔 때마다 잠깐 기다렸다가(연속 호출 방지) 다시 배치합니다.
        if self._relayout_job:
            self.after_cancel(self._relayout_job)
        self._relayout_job = self.after(150, self._refit)

    def _refit(self):
        """창 크기가 바뀐 뒤, 실제로 잰 칸 수가 지금과 다를 때만 다시 배치합니다."""
        self._relayout_job = None
        if self.pages and self._measure_cols() != self.cols:
            self.relayout()

    def on_size_slider(self, value):
        """슬라이더를 움직일 때마다 불립니다. 너무 자주 다시 그리지 않도록 잠깐 기다렸다가 적용합니다."""
        self.thumb_w = int(round(value))
        self.size_label.configure(text=f"{int(round(self.thumb_w / self.DEFAULT_W * 100))}%")
        if self._size_job:
            self.after_cancel(self._size_job)
        self._size_job = self.after(60, self.apply_thumb_size)

    def apply_thumb_size(self):
        """각 칸 그림의 표시 크기를 바꾸고 칸을 다시 배치합니다. (회전한 쪽은 회전된 크기 기준)"""
        self._size_job = None
        f = self.thumb_w / self.RENDER_W
        for en in self.pages:
            rw, rh = en["dims"]
            en["img"].configure(size=(max(1, int(rw * f)), max(1, int(rh * f))))
            en["cap"].configure(wraplength=self.thumb_w + 10)
        self.relayout()

    def update_info(self):
        sel = sum(1 for e in self.pages if e["var"].get())
        self.info.configure(text=f"파일 {len(self.fl.paths)}개  ·  전체 {len(self.pages)}쪽  ·  선택 {sel}쪽")

    # ───────── 파일 추가 / 삭제 / 초기화 ─────────
    def load_new(self, added):
        """새로 추가된 파일의 페이지를 맨 뒤에 붙입니다. (이미 조정해 둔 순서는 그대로 유지)"""
        errors = []
        for p in added:
            try:
                self._render_thumbs(p)
            except Exception as e:
                errors.append(f"{os.path.basename(p)} ({e})")
                if p in self.fl.paths:
                    self.fl.paths.remove(p)
                self._forget_file(p)
                continue
            for i in range(self.page_counts[p]):
                self.pages.append(self._make_cell(p, i))
        self.fl.refresh()
        self.relayout()
        self.app.status.configure(text="준비됨")
        if errors:
            messagebox.showwarning("안내", "불러오지 못한 파일이 있습니다.\n\n" + "\n".join(errors))

    def drop_files(self, removed):
        """목록에서 빠진 파일의 페이지를 미리보기에서도 지웁니다."""
        if not removed:
            return
        keep = []
        for en in self.pages:
            if en["path"] in removed:
                en["cell"].destroy()
            else:
                keep.append(en)
        self.pages = keep
        for p in removed:
            self._forget_file(p)
        self.relayout()

    def reset_pages(self):
        """파일 목록 순서대로 모든 페이지를 처음 상태(회전 없음)로 다시 배열합니다."""
        for en in self.pages:
            en["cell"].destroy()
        self.pages = []
        for p in self.fl.paths:
            for i in range(self.page_counts.get(p, 0)):
                self.pages.append(self._make_cell(p, i))
        self.relayout()

    def ask_reset(self):
        if not self.pages and not self.fl.paths:
            return
        if not self.is_edited(self.fl.paths):
            return messagebox.showinfo("안내", "이미 처음 상태입니다.")
        if messagebox.askokcancel("초기화", "페이지 순서와 삭제·회전한 내용을 모두 처음 상태로 되돌립니다.\n계속할까요?"):
            self.reset_pages()

    # ───────── 선택 / 이동 / 삭제 / 회전 ─────────
    def set_all(self, flag):
        for e in self.pages:
            e["var"].set(flag)
        self.update_info()

    def apply_range(self):
        if not self.pages:
            return
        try:
            sel = parse_pages(self.entry.get(), len(self.pages))
        except ValueError as e:
            msg = str(e)
            if "페이지 범위" not in msg:
                msg = "숫자와 - , 만 사용해 주세요. (예: 1-3,5,8)"
            return messagebox.showwarning("안내", f"범위를 확인해 주세요.\n(미리보기에 표시된 번호 기준입니다.)\n{msg}")
        for i, e in enumerate(self.pages):
            e["var"].set(i in sel)
        self.update_info()

    def move_checked(self, step):
        """체크한 페이지를 한 칸씩 앞(-1) 또는 뒤(+1)로 옮깁니다."""
        idx = [i for i, e in enumerate(self.pages) if e["var"].get()]
        if not idx:
            return messagebox.showinfo("안내", "옮길 페이지를 체크해 주세요.")
        pg = self.pages
        if step < 0:
            if idx[0] == 0:
                return
            for i in idx:
                pg[i], pg[i - 1] = pg[i - 1], pg[i]
        else:
            if idx[-1] == len(pg) - 1:
                return
            for i in reversed(idx):
                pg[i], pg[i + 1] = pg[i + 1], pg[i]
        self.relayout()

    def rotate_checked(self, delta):
        """체크한 페이지를 delta도(-90 / 90 / 180) 회전합니다. 애니메이션 없이 미리보기 그림을 바로 바꿉니다.
        회전 각도는 각 쪽에 기억해 두었다가 저장·낱장 분리할 때 그대로 적용합니다."""
        sel = [(i, e) for i, e in enumerate(self.pages) if e["var"].get()]
        if not sel:
            return messagebox.showinfo("안내", "회전할 페이지를 체크해 주세요.\n"
                                               "(모두 돌리려면 '전체 선택'을 누르세요.)")
        for i, en in sel:
            en["rot"] = (en["rot"] + delta) % 360
            img = self._fit_image((en["path"], en["no"]), en["rot"])
            cimg = self._ctk_image(img)
            en["lbl"].configure(image=cimg)
            en["img"] = cimg
            en["dims"] = (img.width, img.height)
            self._set_caption(i, en)

    def delete_checked(self):
        sel = [e for e in self.pages if e["var"].get()]
        if not sel:
            return messagebox.showinfo("안내", "삭제할 페이지를 체크해 주세요.")
        if not messagebox.askyesno("확인", f"선택한 {len(sel)}쪽을 미리보기에서 지웁니다.\n"
                                           "(원본 파일은 바뀌지 않으며, '순서·삭제·회전 초기화'로 되살릴 수 있습니다.)\n\n"
                                           "계속할까요?"):
            return
        self.pages = [e for e in self.pages if not e["var"].get()]
        for e in sel:
            e["cell"].destroy()
        self.relayout()

    # ───────── 끌어서 순서 바꾸기 (여러 쪽 함께 이동 가능) ─────────
    def _canvas_box(self):
        """미리보기 영역(스크롤 되는 부분)의 화면상 위치를 돌려줍니다."""
        try:
            c = self.grid_frame._parent_canvas
            return c, c.winfo_rootx(), c.winfo_rooty(), c.winfo_width(), c.winfo_height()
        except Exception:
            return None

    def _index_at(self, x, y):
        """화면 좌표 (x, y) 아래에 있는 페이지 칸의 번호를 돌려줍니다. 없으면 None."""
        box = self._canvas_box()
        if box and not (box[1] <= x < box[1] + box[3] and box[2] <= y < box[2] + box[4]):
            return None  # 미리보기 영역 바깥
        for i, en in enumerate(self.pages):
            c = en["cell"]
            cx, cy = c.winfo_rootx(), c.winfo_rooty()
            if cx <= x < cx + c.winfo_width() and cy <= y < cy + c.winfo_height():
                return i
        return None

    def _drag_press(self, event, entry):
        """칸을 누른 순간. 체크된 쪽이 여럿이고 그중 하나를 눌렀다면 체크된 쪽 전체를 함께 끕니다."""
        checked = [e for e in self.pages if e["var"].get()]
        if entry["var"].get() and len(checked) > 1:
            group = checked          # 화면 순서대로 들어 있음
        else:
            group = [entry]          # 체크 안 된 칸이거나 하나만 체크: 그 칸만 이동
        self.drag = {"entry": entry, "group": group, "x": event.x_root, "y": event.y_root,
                     "moved": False, "target": None}

    def _drag_move(self, event):
        d = self.drag
        if not d:
            return
        group_ids = {id(g) for g in d["group"]}
        if not d["moved"]:
            if abs(event.x_root - d["x"]) + abs(event.y_root - d["y"]) < 8:
                return  # 살짝 흔들린 것은 클릭으로 봄
            d["moved"] = True
            for g in d["group"]:  # 함께 끌려가는 칸들을 파랗게 표시
                g["cell"].configure(fg_color=self.C_DRAG)
        t = self._index_at(event.x_root, event.y_root)
        te = self.pages[t] if t is not None else None
        if te is not None and te is not d["target"]:
            old = d["target"]
            if old is not None and id(old) not in group_ids:
                old["cell"].configure(fg_color=self.C_NORMAL)
            d["target"] = te
            if id(te) not in group_ids:  # 놓을 위치는 노랗게 표시
                te["cell"].configure(fg_color=self.C_TARGET)
        # 위/아래 가장자리로 끌면 자동으로 스크롤
        box = self._canvas_box()
        if box:
            c, _, top, _, h = box
            if event.y_root < top + 30:
                c.yview_scroll(-1, "units")
            elif event.y_root > top + h - 30:
                c.yview_scroll(1, "units")

    def _drag_release(self, event):
        d, self.drag = self.drag, None
        if not d or not d["moved"]:
            return
        group = d["group"]
        group_ids = {id(g) for g in group}
        t = self._index_at(event.x_root, event.y_root)
        te = self.pages[t] if t is not None else d["target"]
        for en in self.pages:
            en["cell"].configure(fg_color=self.C_NORMAL)
        if te is None or id(te) in group_ids:
            return  # 자기 자신(또는 함께 끄는 칸) 위에 놓으면 이동하지 않음
        src = self.pages.index(d["entry"])
        dst = self.pages.index(te)
        # 끄는 칸들을 모두 빼낸 뒤, 놓은 칸의 앞(위로 이동) 또는 뒤(아래로 이동)에 한 덩어리로 끼워 넣음
        rest = [e for e in self.pages if id(e) not in group_ids]
        pos = rest.index(te) + (1 if dst > src else 0)
        rest[pos:pos] = group
        self.pages = rest
        self.relayout()

    # ───────── 저장 ─────────
    def build_pdf(self, entries, out, numbering=False):
        """entries: [(파일, 쪽번호, 회전각도), ...] 순서대로 하나의 PDF로 저장합니다. (별도 스레드에서 실행)
        회전각도가 0이 아니면 그 쪽을 시계 방향으로 돌려서 저장합니다.
        numbering 이 True 이면 저장되는 PDF의 각 쪽에 1/전체쪽수 번호를 찍습니다."""
        runs = []  # 같은 파일의 연속된 쪽(회전각도도 같은 것)은 묶어서 한 번에 복사하면 훨씬 빠름
        for p, n, rot in entries:
            if runs and runs[-1][0] == p and runs[-1][2] == n - 1 and runs[-1][3] == rot:
                runs[-1][2] = n
            else:
                runs.append([p, n, n, rot])
        docs = {}
        try:
            with pymupdf.open() as dst:
                done = 0
                for p, a, b, rot in runs:
                    if p not in docs:
                        docs[p] = open_pdf(p)
                    start = dst.page_count
                    dst.insert_pdf(docs[p], from_page=a, to_page=b)
                    if rot:
                        for k in range(start, dst.page_count):
                            pg = dst[k]
                            pg.set_rotation((pg.rotation + rot) % 360)
                    done += b - a + 1
                    self.app.report(done, len(entries), "저장 중")
                self.app.busy("PDF 저장 중...")
                if numbering:
                    stamp_page_numbers(dst)  # 저장될 결과의 순서대로 1/전체쪽수 (회전한 쪽도 바르게 찍힘)
                dst.save(out, garbage=3, deflate=True)
        finally:
            for d in docs.values():
                d.close()

    def save_pages(self, only_checked):
        if not self.pages:
            return messagebox.showinfo("안내", "먼저 PDF 파일을 추가해 주세요.")
        entries = [(e["path"], e["no"], e["rot"]) for e in self.pages if e["var"].get() or not only_checked]
        if not entries:
            return messagebox.showinfo("안내", "저장할 페이지를 체크해 주세요.")
        files = []
        for p, _, _ in entries:
            if p not in files:
                files.append(p)
        if len(files) == 1:
            stem = os.path.splitext(os.path.basename(files[0]))[0]
            initial = f"{stem}_추출.pdf" if only_checked else f"{stem}_편집.pdf"
        else:
            initial = "추출문서.pdf" if only_checked else "합친문서.pdf"
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                           initialfile=initial)
        if not out:
            return
        if any(same_path(out, p) for p in files):
            return messagebox.showwarning("안내", "원본 파일과 다른 이름으로 저장해 주세요.")
        numbering = self.num_var.get()  # 별도 스레드에서 화면 값을 읽지 않도록 미리 가져옴

        def work():
            self.build_pdf(entries, out, numbering)
            return len(entries)

        self.app.run_bg(work, lambda n, e: messagebox.showerror("오류", str(e)) if e
                        else messagebox.showinfo("완료", f"{n}쪽을 하나의 PDF로 저장했습니다."),
                        "PDF를 만드는 중...")

    def split_checked(self):
        if not self.pages:
            return messagebox.showinfo("안내", "먼저 PDF 파일을 추가해 주세요.")
        # (원본 파일, 원본 쪽번호, 미리보기에 보이는 순서 번호, 회전각도) - 순서 번호는 쪽 번호를 찍을 때 사용
        entries = [(e["path"], e["no"], i + 1, e["rot"]) for i, e in enumerate(self.pages) if e["var"].get()]
        if not entries:
            return messagebox.showinfo("안내", "분리할 페이지를 체크해 주세요.\n"
                                               "(모두 분리하려면 '전체 선택'을 누르세요.)")
        folder = filedialog.askdirectory(title="저장할 폴더 선택")
        if not folder:
            return
        numbering = self.num_var.get()  # 별도 스레드에서 화면 값을 읽지 않도록 미리 가져옴
        total = len(self.pages)         # '3/20' 에서 20 에 해당하는 값 (미리보기의 전체 쪽수)

        def work():
            docs, used = {}, set()
            try:
                for k, (p, no, order, rot) in enumerate(entries):
                    self.app.report(k, len(entries), "분리 저장 중")
                    if p not in docs:
                        docs[p] = open_pdf(p)
                    # 원래 페이지 번호를 파일명에 사용 (예: 문서_005.pdf). 이름이 겹치면 _2, _3 을 붙임
                    base = f"{os.path.splitext(os.path.basename(p))[0]}_{no + 1:03d}"
                    name, c = base, 2
                    while name.lower() in used:
                        name = f"{base}_{c}"
                        c += 1
                    used.add(name.lower())
                    with pymupdf.open() as d:
                        d.insert_pdf(docs[p], from_page=no, to_page=no)
                        if rot:
                            d[0].set_rotation((d[0].rotation + rot) % 360)
                        if numbering:
                            stamp_page_numbers(d, numbers=[order], total=total)
                        d.save(os.path.join(folder, name + ".pdf"), deflate=True)
            finally:
                for d in docs.values():
                    d.close()
            return len(entries)

        self.app.run_bg(work, lambda n, e: messagebox.showerror("오류", str(e)) if e
                        else messagebox.showinfo("완료", f"선택한 {n}쪽을 낱장 파일로 분리했습니다."),
                        "분리 저장 중...")


class ImagePage(BasePage):
    def __init__(self, master, app):
        super().__init__(master, app, "PDF → 이미지 변환하기",
                         "PDF의 모든 페이지를 이미지 파일로 저장합니다. PDF 파일을 끌어다 놓아도 불러옵니다.")
        self.src = None
        card = ctk.CTkFrame(self, fg_color="white", corner_radius=12)
        card.pack(fill="x", padx=28)
        r1 = ctk.CTkFrame(card, fg_color="transparent")
        r1.pack(fill="x", padx=20, pady=(20, 8))
        ctk.CTkButton(r1, text="PDF 파일 선택", width=130, command=self.load).pack(side="left")
        self.name = ctk.CTkLabel(r1, text="선택된 파일 없음", text_color="#667085")
        self.name.pack(side="left", padx=12)

        r2 = ctk.CTkFrame(card, fg_color="transparent")
        r2.pack(fill="x", padx=20, pady=8)
        ctk.CTkLabel(r2, text="이미지 형식").pack(side="left")
        self.fmt = ctk.CTkSegmentedButton(r2, values=["jpg", "png", "bmp"])
        self.fmt.set("jpg")
        self.fmt.pack(side="left", padx=(10, 30))
        ctk.CTkLabel(r2, text="해상도(DPI)").pack(side="left")
        self.dpi = ctk.CTkComboBox(r2, values=["72", "100", "150", "200", "300"], width=90)
        self.dpi.set("150")
        self.dpi.pack(side="left", padx=10)

        ctk.CTkButton(card, text="변환하여 폴더에 저장", height=42,
                      command=self.convert).pack(pady=(14, 22))
        self.enable_drop()

    def load(self, p=None):
        p = p or self.pick_pdf()
        if p:
            self.src = p
            self.name.configure(text=os.path.basename(p))

    def convert(self):
        if not self.src:
            return messagebox.showinfo("안내", "먼저 PDF 파일을 선택해 주세요.")
        try:
            dpi = int(self.dpi.get())
        except ValueError:
            return messagebox.showwarning("안내", "DPI는 숫자로 입력해 주세요.")
        folder = filedialog.askdirectory(title="저장할 폴더 선택")
        if not folder:
            return
        ext = self.fmt.get()
        stem = os.path.splitext(os.path.basename(self.src))[0]

        def work():
            with open_pdf(self.src) as doc:
                n = doc.page_count
                for i, page in enumerate(doc):
                    self.app.report(i, n, "이미지 변환 중")
                    img = page_to_pil(page, dpi)
                    path = os.path.join(folder, f"{stem}_{i + 1:03d}.{ext}")
                    if ext == "jpg":
                        img.save(path, "JPEG", quality=95)
                    else:
                        img.save(path, ext.upper())
            return n

        self.app.run_bg(work, lambda n, e: messagebox.showerror("오류", str(e)) if e
                        else messagebox.showinfo("완료", f"{n}장의 이미지를 저장했습니다."), "이미지 변환 중...")


class CompressPage(BasePage):
    LEVELS = [("max", "가장 많이 줄이기", "용량 최우선. 이미지를 가장 낮은 화질로 줄이고, 그래도 큰 쪽은 통째로 그림으로 바꿀 수 있습니다. "
                                         "(그림이 된 쪽은 글자 선택·검색 불가)"),
              ("high", "많이 줄이기", "이미지 해상도·화질을 눈에 띄게 낮춥니다. 글자는 그대로 유지됩니다."),
              ("normal", "보통으로 줄이기", "시인성 우선. 해상도가 너무 높은 이미지만 줄입니다.")]

    # thr: 이 해상도(dpi)보다 높은 이미지를 줄임 / tgt: 줄일 목표 해상도 / q: JPEG 화질(낮을수록 용량↓)
    # always: 해상도가 낮아도 화질만 낮춰 다시 저장할지 / raster: (dpi, 화질) 쪽 전체를 그림으로 바꾸는 최후 수단
    CFG = {
        "normal": dict(thr=200, tgt=150, q=75, always=False, raster=None),
        "high":   dict(thr=120, tgt=90,  q=45, always=True,  raster=None),
        "max":    dict(thr=90,  tgt=72,  q=30, always=True,  raster=(90, 35)),
    }

    def __init__(self, master, app):
        super().__init__(master, app, "PDF 용량 줄이기",
                         "압축 강도를 고르고 저장하면 새 파일로 만들어집니다. 원본은 그대로 남습니다. "
                         "PDF 파일을 끌어다 놓아도 불러옵니다.\n"
                         "※ 어떤 경우에도 결과가 원본보다 커지지는 않습니다. (더 줄일 수 없으면 원본 용량 그대로 저장)")
        self.src = None
        card = ctk.CTkFrame(self, fg_color="white", corner_radius=12)
        card.pack(fill="x", padx=28)
        r1 = ctk.CTkFrame(card, fg_color="transparent")
        r1.pack(fill="x", padx=20, pady=(20, 8))
        ctk.CTkButton(r1, text="PDF 파일 선택", width=130, command=self.load).pack(side="left")
        self.name = ctk.CTkLabel(r1, text="선택된 파일 없음", text_color="#667085")
        self.name.pack(side="left", padx=12)

        self.level = tk.StringVar(value="high")
        for key, title, desc in self.LEVELS:
            row = ctk.CTkFrame(card, fg_color="transparent")
            row.pack(fill="x", padx=20, pady=4)
            ctk.CTkRadioButton(row, text=title, variable=self.level, value=key,
                               font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
            ctk.CTkLabel(row, text=desc, text_color="#667085", anchor="w", justify="left",
                         wraplength=720).pack(anchor="w", padx=28)
        ctk.CTkButton(card, text="압축하여 저장", height=42, command=self.run).pack(pady=(14, 22))
        self.enable_drop()

    def load(self, p=None):
        p = p or self.pick_pdf()
        if p:
            self.src = p
            self.name.configure(text=f"{os.path.basename(p)}  ({fmt_size(os.path.getsize(p))})")

    # ── 압축 알고리즘 ──
    @staticmethod
    def _shrink_images(doc, page, cfg, seen, stats):
        """이 쪽에 쓰인 이미지를 줄입니다. 같은 이미지가 여러 쪽에 쓰였어도 한 번만 처리합니다."""
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            stats["images"] += 1
            try:
                if img[1]:  # 투명 마스크가 있는 이미지는 건드리지 않음
                    continue
                info = doc.extract_image(xref)
                if not info or info.get("bpc", 8) == 1:  # 흑백(1비트) 이미지는 이미 작음
                    continue
                if info.get("colorspace") not in (1, 3):  # 흑백/컬러만 처리
                    continue
                old_len = len(doc.xref_stream_raw(xref))
                if old_len < 8 * 1024:  # 너무 작은 이미지는 건너뜀
                    continue
                rects = page.get_image_rects(xref)
                if not rects:
                    continue
                w_in = max(r.width for r in rects) / 72.0
                w, h = info["width"], info["height"]
                eff_dpi = w / w_in if w_in > 0 else 0
                scale = 1.0
                if eff_dpi > cfg["thr"]:
                    scale = cfg["tgt"] / eff_dpi
                elif not cfg["always"]:
                    continue
                im = Image.open(io.BytesIO(info["image"]))
                if im.mode in ("RGBA", "LA", "PA"):
                    continue
                if im.mode not in ("RGB", "L"):
                    im = im.convert("RGB")
                if scale < 1.0:
                    im = im.resize((max(1, round(w * scale)), max(1, round(h * scale))), Image.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=cfg["q"], optimize=True)
                data = buf.getvalue()
                if len(data) < old_len * 0.95:  # 5% 이상 줄 때만 교체
                    page.replace_image(xref, stream=data)
                    stats["shrunk"] += 1
            except Exception:
                continue  # 문제가 있는 이미지는 그대로 두고 다음으로

    @staticmethod
    def _maybe_rasterize(doc, page, raster, stats):
        """이미지 용량이 큰 쪽을 낮은 화질의 그림 한 장으로 바꿔서 더 작아질 때만 교체합니다."""
        if page.rotation != 0:
            return
        img_bytes = 0
        for img in page.get_images(full=True):
            for x in (img[0], img[1]):
                if x:
                    try:
                        img_bytes += len(doc.xref_stream_raw(x))
                    except Exception:
                        pass
        if img_bytes < 30 * 1024:  # 이미지가 거의 없는 쪽(글자 위주)은 그대로 둠 -> 용량이 늘어나는 것을 방지
            return
        dpi, q = raster
        buf = io.BytesIO()
        page_to_pil(page, dpi).save(buf, "JPEG", quality=q, optimize=True)
        data = buf.getvalue()
        if len(data) < img_bytes * 0.7:  # 확실히 작아질 때만
            for cx in page.get_contents():
                doc.update_stream(cx, b" ")
            page.insert_image(page.rect, stream=data)
            stats["rasterized"] += 1

    @classmethod
    def compress(cls, src, dst, level, report, busy):
        cfg = cls.CFG[level]
        doc = open_pdf(src)
        n = doc.page_count
        stats = {"images": 0, "shrunk": 0, "rasterized": 0}
        seen = set()
        can_replace = hasattr(pymupdf.Page, "replace_image")
        try:
            for i in range(n):
                report(i, n, "압축 중")
                page = doc[i]
                if can_replace:
                    cls._shrink_images(doc, page, cfg, seen, stats)
                if cfg["raster"]:
                    cls._maybe_rasterize(doc, page, cfg["raster"], stats)
            report(n, n, "압축 중")
            if not can_replace:  # 구버전 PyMuPDF 대비
                try:
                    doc.rewrite_images(dpi_threshold=cfg["thr"], dpi_target=cfg["tgt"], quality=cfg["q"])
                except Exception:
                    pass
            busy("마무리 정리·저장 중... (큰 파일은 시간이 걸릴 수 있습니다)")
            try:
                doc.subset_fonts()
            except Exception:
                pass
            doc.save(dst, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
        finally:
            doc.close()
        return stats

    def run(self):
        if not self.src:
            return messagebox.showinfo("안내", "먼저 PDF 파일을 선택해 주세요.")
        out = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                           initialfile=os.path.splitext(os.path.basename(self.src))[0] + "_압축.pdf")
        if not out:
            return
        if os.path.abspath(out) == os.path.abspath(self.src):
            return messagebox.showwarning("안내", "원본과 다른 이름으로 저장해 주세요.")
        level, src = self.level.get(), self.src

        def work():
            stats = self.compress(src, out, level, self.app.report, self.app.busy)
            stats["same"] = False
            if os.path.getsize(out) >= os.path.getsize(src):
                shutil.copyfile(src, out)  # 커졌다면 원본 그대로 저장
                stats["same"] = True
            return stats

        def done(stats, e):
            if e:
                return messagebox.showerror("오류", str(e))
            a, b = os.path.getsize(src), os.path.getsize(out)
            if stats["same"] or b >= a:
                msg = (f"{fmt_size(a)} → {fmt_size(b)}\n이 파일은 이미 충분히 압축되어 있어 더 줄어들지 않았습니다.\n"
                       "(원본과 같은 내용으로 저장했습니다.)")
            else:
                msg = f"{fmt_size(a)} → {fmt_size(b)}  ({(1 - b / a) * 100:.0f}% 감소)"
                detail = []
                if stats["images"]:
                    detail.append(f"이미지 {stats['images']}개 중 {stats['shrunk']}개 축소")
                if stats["rasterized"]:
                    detail.append(f"{stats['rasterized']}쪽을 그림으로 변환")
                if detail:
                    msg += "\n" + " / ".join(detail)
            messagebox.showinfo("완료", msg)

        self.app.run_bg(lambda: work(), done, "압축 준비 중...")


# ───────────────────────── 메인 창 ─────────────────────────
class App(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self):
        super().__init__()
        self.TkdndVersion = TkinterDnD._require(self)  # 드래그앤드롭 기능 켜기
        self.title("PDF 도우미")
        self.geometry("1240x760")
        self.minsize(1100, 680)

        side = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color="#1F2A44")
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        ctk.CTkLabel(side, text="PDF 도우미", text_color="white",
                     font=ctk.CTkFont(size=24, weight="bold")).pack(pady=(30, 24))

        main = ctk.CTkFrame(self, corner_radius=0, fg_color=BG)
        main.pack(side="left", fill="both", expand=True)
        bar = ctk.CTkFrame(main, height=36, fg_color="white", corner_radius=0)
        bar.pack(side="bottom", fill="x")
        self.status = ctk.CTkLabel(bar, text="준비됨", text_color="#667085", anchor="w")
        self.status.pack(side="left", padx=16)
        self.progress = ctk.CTkProgressBar(bar, mode="indeterminate", width=220)
        self._determinate = False

        self.pages, self.buttons = {}, {}
        # (키, 버튼 이름, 화면 클래스, 버튼 아래에 작게 붙일 주석) - 이 순서대로 왼쪽 메뉴에 표시됩니다.
        menu = [("extract", "✂  합치기·분리하기·회전하기", ExtractPage, ""),
            ("merge", "📎  PDF로 변환하기", MergePage, ""),
            ("mask", "⬛  PDF 마스킹", MaskPage, ""),
            ("image", "🖼  이미지 변환하기", ImagePage, ""),
            ("compress", "🗜  용량 줄이기", CompressPage, "")]
        for key, label, cls, note in menu:
            self.pages[key] = cls(main, self)
            b = ctk.CTkButton(side, text=label, anchor="w", height=44, corner_radius=8,
                              fg_color="transparent", hover_color="#33415F",
                              font=ctk.CTkFont(size=14), command=lambda k=key: self.show(k))
            b.pack(fill="x", padx=14, pady=(3, 0 if note else 3))
            self.buttons[key] = b
            if note:
                ctk.CTkLabel(side, text=note, anchor="w", text_color="#9FB0D0",
                             font=ctk.CTkFont(size=11)).pack(fill="x", padx=(48, 14), pady=(0, 3))
        self.show("extract")
        check_for_update_async(self)  # 창이 뜬 뒤 새 버전 확인 (뒤에서 조용히 진행)

    def show(self, key):
        for k, p in self.pages.items():
            p.pack_forget()
            self.buttons[k].configure(fg_color="#3B8ED0" if k == key else "transparent")
        self.pages[key].pack(fill="both", expand=True)

    # ── 상태·진행도 표시 (다른 스레드에서도 안전하게 호출) ──
    def say(self, text):
        self.after(0, lambda: self.status.configure(text=text))

    def report(self, i, n, text="", unit="쪽"):
        """진행도 표시: n개 중 i개 완료. 진행 막대가 채워지고 '12/80쪽 (15%)' 같은 글자가 나옵니다."""
        def ui():
            if not self._determinate:
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self._determinate = True
            ratio = (i / n) if n else 0
            self.progress.set(ratio)
            self.status.configure(text=f"{text}  {i}/{n}{unit} ({int(ratio * 100)}%)")
        self.after(0, ui)

    def busy(self, text):
        """개수를 셀 수 없는 단계(저장·정리 등)에서는 막대를 좌우로 움직이는 표시로 바꿉니다."""
        def ui():
            if self._determinate:
                self.progress.configure(mode="indeterminate")
                self._determinate = False
                self.progress.start()
            self.status.configure(text=text)
        self.after(0, ui)

    def run_bg(self, work, done, msg="작업 중..."):
        """오래 걸리는 작업을 별도 스레드에서 실행해 창이 멈추지 않게 합니다."""
        self.status.configure(text=msg)
        self._determinate = False
        self.progress.configure(mode="indeterminate")
        self.progress.pack(side="right", padx=16)
        self.progress.start()

        def target():
            try:
                res, err = work(), None
            except Exception as e:
                res, err = None, e

            def finish():
                self.progress.stop()
                self.progress.pack_forget()
                self.status.configure(text="준비됨")
                done(res, err)
            self.after(0, finish)

        threading.Thread(target=target, daemon=True).start()


if __name__ == "__main__":
    try:
        ctk.ThemeManager.theme["CTkFont"]["family"] = FONT
    except Exception:
        pass
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    App().mainloop()
