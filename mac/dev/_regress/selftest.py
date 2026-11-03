# -*- coding: utf-8 -*-
"""
selftest.py —— i4 核心引擎自检
================================
用合成 PDF 验证：补白逻辑 + 拼接逻辑 + 页数正确性。
"""

import os
import sys
import tempfile
from pathlib import Path

import os
import sys


def _project_root(_p):
    """向上找到含 app/paths.py 的那一级（不再按层数数）。"""
    while True:
        if os.path.isfile(os.path.join(_p, "app", "paths.py")):
            return _p
        _up = os.path.dirname(_p)
        if _up == _p:
            return _p
        _p = _up


_ROOT = _project_root(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_ROOT, "app"))
import paths  # noqa: E402
paths.install()

import pymupdf as fitz  # noqa: E402
from i4_core import (  # noqa: E402
    PrintMode, collect_items, merge_pdfs, pad_blank_page,
    pdf_page_count, run_pipeline,
)


def mk_pdf(path: Path, pages: int, w: int = 595, h: int = 842):
    doc = fitz.open()
    for _ in range(pages):
        doc.new_page(width=w, height=h)
    doc.save(str(path))
    doc.close()


def main():
    ok = fail = 0

    def chk(name, cond, extra=""):
        nonlocal ok, fail
        if cond:
            ok += 1
        else:
            fail += 1
            print(f"  !! {name}{'  -> ' + extra if extra else ''}")

    tmp = Path(tempfile.mkdtemp(prefix="i4_selftest_"))

    # ── 基础 PDF 操作 ──
    p1 = tmp / "a.pdf"; mk_pdf(p1, 3)   # 奇数页
    p2 = tmp / "b.pdf"; mk_pdf(p2, 2)   # 偶数页
    p3 = tmp / "c.pdf"; mk_pdf(p3, 1)   # 奇数页

    chk("pdf_page_count 3页", pdf_page_count(p1) == 3)
    chk("pdf_page_count 2页", pdf_page_count(p2) == 2)

    # ── 补白 ──
    r1 = tmp / "a_pad.pdf"
    did1 = pad_blank_page(p1, r1)
    chk("奇数页补白 -> True", did1)
    chk("补白后4页", pdf_page_count(r1) == 4)

    r2 = tmp / "b_pad.pdf"
    did2 = pad_blank_page(p2, r2)
    chk("偶数页不补白 -> False", not did2)
    chk("不补白仍2页", pdf_page_count(r2) == 2)

    # ── 拼接 ──
    merged = tmp / "merged.pdf"
    n = merge_pdfs([r1, r2], merged)
    chk("拼接 4+2=6页", n == 6, f"实际 {n}")

    # ── 双面全流程 ──
    items = collect_items([p1, p2, p3])
    ok1, msg1, plan1 = run_pipeline(items, PrintMode.DUPLEX,
                                     tmp / "duplex.pdf")
    chk("双面流程成功", ok1, msg1)
    chk("双面 3+2+1+2白=8页", plan1.total_pages == 8,
        f"实际 {plan1.total_pages}")
    chk("双面补了2白", plan1.blank_added == 2,
        f"实际 {plan1.blank_added}")

    # ── 单面全流程 ──
    items2 = collect_items([p1, p2, p3])
    ok2, msg2, plan2 = run_pipeline(items2, PrintMode.SIMPLEX,
                                      tmp / "simplex.pdf")
    chk("单面流程成功", ok2, msg2)
    chk("单面 3+2+1=6页", plan2.total_pages == 6,
        f"实际 {plan2.total_pages}")
    chk("单面不补白", plan2.blank_added == 0)

    # ── 全偶数（双面也不补） ──
    p4 = tmp / "d.pdf"; mk_pdf(p4, 4)
    items3 = collect_items([p2, p4])
    ok3, msg3, plan3 = run_pipeline(items3, PrintMode.DUPLEX,
                                      tmp / "all_even.pdf")
    chk("全偶数成功", ok3, msg3)
    chk("全偶数 2+4=6页", plan3.total_pages == 6)
    chk("全偶数不补白", plan3.blank_added == 0)

    print(f"\n自检：{ok} 项通过，{fail} 项失败")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
