# -*- coding: utf-8 -*-
"""
i4_core.py —— PDF 奇偶拼接打印 · 核心引擎
==========================================
需求（2026-10-27）：
  多份 Word / Excel（作业/报告，1~5 页）→ 转 PDF → 拼成一份大 PDF 一键打印。
  打印机没有自动翻页（手动双面）→ 双面模式给奇数页文档末尾补空白页凑偶数，
  防止第二份作业印到第一份背面；单面模式直拼不补。

转换链路移植自项目3（PDF转换拼接4.0）Word/Excel→PDF 部分，
只保留 Office 两族，去掉 CAD / BarTender / Visio。
新增：奇数页补白（pad_blank_page）+ PDF 拼接（merge_pdfs）。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple

try:
    import pymupdf as fitz
except ImportError:
    import fitz  # type: ignore


# ══════════════════════════════════════════════════════════════════════
# ① 数据模型
# ══════════════════════════════════════════════════════════════════════

class PrintMode(str, Enum):
    DUPLEX = "duplex"       # 双面打印：奇数页补白
    SIMPLEX = "simplex"     # 单面打印：不补白


class ConvStatus(str, Enum):
    PENDING = "pending"
    OK = "ok"
    FAIL = "fail"


@dataclass
class ConvItem:
    src: Path
    status: ConvStatus = ConvStatus.PENDING
    pdf_path: Optional[Path] = None
    pages: int = 0
    padded: bool = False
    error: str = ""


@dataclass
class MergePlan:
    items: List[ConvItem] = field(default_factory=list)
    mode: PrintMode = PrintMode.DUPLEX
    total_pages: int = 0
    blank_added: int = 0


# ══════════════════════════════════════════════════════════════════════
# ② 文件类型识别
# ══════════════════════════════════════════════════════════════════════

WORD_EXTS = {".doc", ".docx", ".docm", ".rtf"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm"}
PDF_EXTS = {".pdf"}
ALL_EXTS = WORD_EXTS | EXCEL_EXTS | PDF_EXTS


def is_supported(p: Path) -> bool:
    return p.suffix.lower() in ALL_EXTS


def family_of(p: Path) -> str:
    ext = p.suffix.lower()
    if ext in WORD_EXTS:
        return "word"
    if ext in EXCEL_EXTS:
        return "excel"
    if ext in PDF_EXTS:
        return "pdf"
    return ""


# ══════════════════════════════════════════════════════════════════════
# ③ Word / Excel → PDF 转换（移植自项目3）
# ══════════════════════════════════════════════════════════════════════

def _com_retry(call, *args, **kwargs):
    """COM 调用带重试（Office 偶发 busy）。"""
    last = None
    for _ in range(3):
        try:
            return call(*args, **kwargs)
        except Exception as exc:  # noqa: BLE001
            last = exc
            text = str(exc).lower()
            if any(k in text for k in ("busy", "正在使用", "0x800a0046")):
                import time
                time.sleep(1.0)
                continue
            raise
    raise last  # type: ignore[misc]


def _quiet_close(obj) -> None:
    if obj is None:
        return
    try:
        obj.Close(SaveChanges=False)
    except Exception:  # noqa: BLE001
        pass


def _win_convert_document(app, src: Path, dst: Path) -> None:
    doc = None
    try:
        doc = _com_retry(app.Documents.Open, str(src),
                         ReadOnly=True, AddToRecentFiles=False)
        _com_retry(doc.SaveAs, str(dst), FileFormat=17)  # wdFormatPDF
    finally:
        _quiet_close(doc)


def _win_convert_workbook(app, src: Path, dst: Path) -> None:
    wb = None
    try:
        wb = _com_retry(app.Workbooks.Open, str(src), ReadOnly=True)
        _com_retry(wb.ExportAsFixedFormat, 0, str(dst))  # xlTypePDF
    finally:
        _quiet_close(wb)


WIN_RUNNERS = {
    "word": _win_convert_document,
    "excel": _win_convert_workbook,
}


def _win_convert(src: Path, dst: Path) -> str:
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except ImportError:
        return "缺少 pywin32（pip install pywin32）"

    fam = family_of(src)
    runner = WIN_RUNNERS.get(fam)
    if not runner:
        return f"不支持的类型：{src.suffix}"

    pythoncom.CoInitialize()
    app = None
    try:
        prog_id = "Word.Application" if fam == "word" else "Excel.Application"
        app = win32com.client.Dispatch(prog_id)
        app.Visible = False
        try:
            app.DisplayAlerts = False
        except Exception:  # noqa: BLE001
            pass
        runner(app, src, dst)
        return ""
    except Exception as exc:  # noqa: BLE001
        return _classify_error(str(exc))
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass


def _mac_convert(src: Path, dst: Path) -> str:
    fam = family_of(src)
    if fam == "word":
        app_name = "Microsoft Word"
    elif fam == "excel":
        app_name = "Microsoft Excel"
    else:
        return f"不支持的类型：{src.suffix}"

    script = f'''
tell application "{app_name}"
    set beforeCount to (count of documents)
    open (POSIX file "{src}" as alias)
    set d to missing value
    repeat 40 times
        delay 0.25
        try
            if (count of documents) > beforeCount then set d to document (beforeCount + 1)
        end try
        if d is not missing value then exit repeat
    end repeat
    if d is not missing value then
        save as d file name (POSIX file "{dst}" as text) file format format PDF
        close d saving no
    end if
end tell
'''
    try:
        r = subprocess.run(["osascript", "-e", script],
                           capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            return _classify_error(r.stderr)
        return "" if dst.exists() else "转换后文件未生成"
    except Exception as exc:  # noqa: BLE001
        return str(exc)


def _libre_convert(src: Path, dst: Path) -> str:
    soffice = _find_soffice()
    if not soffice:
        return "找不到 LibreOffice"
    outdir = dst.parent
    try:
        r = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(outdir), str(src)],
            capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            return _classify_error(r.stderr)
        expected = outdir / (src.stem + ".pdf")
        if expected != dst and expected.exists():
            expected.rename(dst)
        return "" if dst.exists() else "转换后文件未生成"
    except Exception as exc:  # noqa: BLE001
        return str(exc)


def _find_soffice() -> Optional[str]:
    for name in ("soffice", "libreoffice"):
        try:
            r = subprocess.run(["which", name], capture_output=True, text=True)
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip()
        except Exception:  # noqa: BLE001
            pass
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.isfile(p):
            return p
    return None


def convert_one(src: Path, dst: Path) -> str:
    """把一份 Office/PDF 文档转成（或拷贝成）PDF。返回 "" 表示成功。"""
    if family_of(src) == "pdf":
        shutil.copy2(src, dst)
        return ""

    if sys.platform == "win32":
        err = _win_convert(src, dst)
        if not err:
            return ""
        return _libre_convert(src, dst)
    elif sys.platform == "darwin":
        err = _mac_convert(src, dst)
        if not err:
            return ""
        return _libre_convert(src, dst)
    else:
        return _libre_convert(src, dst)


def _classify_error(raw: str) -> str:
    raw_l = raw.lower()
    if any(k in raw_l for k in ("permission", "拒绝访问", "not authorized",
                                 "-10004", "0x80070005")):
        return "系统没放行「自动化控制」，请在 系统设置→隐私与安全性→自动化 中允许"
    if any(k in raw_l for k in ("busy", "正在使用", "0x800a0046")):
        return "文件被占用，请关闭后重试"
    if any(k in raw_l for k in ("not installed", "没有安装", "找不到",
                                 "class not registered", "0x80040154")):
        return "未检测到 Office，请安装 Microsoft Word / Excel"
    return raw[:120]


# ══════════════════════════════════════════════════════════════════════
# ④ PDF 操作
# ══════════════════════════════════════════════════════════════════════

def pdf_page_count(path: Path) -> int:
    try:
        doc = fitz.open(str(path))
        n = doc.page_count
        doc.close()
        return n
    except Exception:  # noqa: BLE001
        return 0


def pad_blank_page(src_pdf: Path, dst_pdf: Path) -> bool:
    """奇数页 PDF 末尾补一页空白 → 偶数页。返回是否补了白。"""
    doc = fitz.open(str(src_pdf))
    n = doc.page_count
    if n % 2 == 0:
        doc.save(str(dst_pdf))
        doc.close()
        return False
    last = doc[n - 1]
    w, h = last.rect.width, last.rect.height
    doc.new_page(width=w, height=h)
    doc.save(str(dst_pdf))
    doc.close()
    return True


def merge_pdfs(pdf_paths: Sequence[Path], out_path: Path) -> int:
    """把多份 PDF 拼成一份。返回总页数。"""
    merged = fitz.open()
    for p in pdf_paths:
        src = fitz.open(str(p))
        merged.insert_pdf(src)
        src.close()
    total = merged.page_count
    merged.save(str(out_path))
    merged.close()
    return total


# ══════════════════════════════════════════════════════════════════════
# ⑤ 主流程
# ══════════════════════════════════════════════════════════════════════

def collect_items(paths: Sequence[Path]) -> List[ConvItem]:
    return [ConvItem(src=p) for p in paths if is_supported(p)]


def run_pipeline(items: List[ConvItem], mode: PrintMode,
                 out_path: Path,
                 workdir: Optional[Path] = None,
                 progress: Optional[Callable[[int, int, str], None]] = None
                 ) -> Tuple[bool, str, MergePlan]:
    """转换 → 补白 → 拼接。返回 (成功?, 消息, MergePlan)。"""
    if not items:
        return False, "没有可处理的文件", MergePlan(mode=mode)

    plan = MergePlan(items=items, mode=mode)
    total = len(items)

    if workdir is None:
        workdir = Path(tempfile.mkdtemp(prefix="i4_merge_"))
    workdir.mkdir(parents=True, exist_ok=True)

    # ① 转换
    converted: List[Path] = []
    for i, it in enumerate(items):
        if progress:
            progress(i, total * 2, f"转换 {it.src.name} …")
        dst = workdir / f"{i:03d}_{it.src.stem}.pdf"
        err = convert_one(it.src, dst)
        if err:
            it.status = ConvStatus.FAIL
            it.error = err
            return False, f"转换失败：{it.src.name}\n{err}", plan
        it.status = ConvStatus.OK
        it.pdf_path = dst
        it.pages = pdf_page_count(dst)
        if it.pages == 0:
            it.status = ConvStatus.FAIL
            it.error = "转换后 PDF 为空"
            return False, f"转换失败：{it.src.name}（PDF 为空）", plan
        converted.append(dst)

    # ② 补白（双面模式才补）
    if mode == PrintMode.DUPLEX:
        padded_paths: List[Path] = []
        for i, it in enumerate(items):
            if progress:
                progress(total + i, total * 2, f"检查页数 {it.src.name} …")
            padded = workdir / f"{i:03d}_{it.src.stem}_pad.pdf"
            did_pad = pad_blank_page(it.pdf_path, padded)
            it.padded = did_pad
            if did_pad:
                plan.blank_added += 1
                it.pages += 1
            padded_paths.append(padded)
        converted = padded_paths

    # ③ 拼接
    if progress:
        progress(total * 2 - 1, total * 2, "拼接中 …")
    total_pages = merge_pdfs(converted, out_path)
    plan.total_pages = total_pages

    if progress:
        progress(total * 2, total * 2, "完成")

    return True, f"已生成：{out_path.name}（{total_pages} 页）", plan
