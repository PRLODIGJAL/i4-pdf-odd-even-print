#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 i4 测试案例：Word/Excel 多页数作业文件，验证奇偶补白与拼接。"""
import os, sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook

OUT = Path.home() / "Desktop" / "i4测试案例"
OUT.mkdir(exist_ok=True)

def make_docx(path, pages, title, body_lines):
    """生成指定页数的 Word 文档（用分页符硬切，保证页数准确）"""
    doc = Document()
    # 设置页边距小一点，保证一页放得下内容
    for sec in doc.sections:
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(2.5)
        sec.right_margin = Cm(2.5)

    for p in range(pages):
        if p > 0:
            doc.add_page_break()
        h = doc.add_heading(f"{title} — 第 {p+1} 页 / 共 {pages} 页", level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("")
        for line in body_lines:
            doc.add_paragraph(line, style="List Bullet" if line.startswith("·") else None)
        doc.add_paragraph("")
        sig = doc.add_paragraph()
        sig.add_run(f"\n\n学生签名：____________    日期：____________\n（第 {p+1} 页）").font.size = Pt(11)

    doc.save(str(path))
    print(f"  ✓ {path.name}  ({pages} 页)")

def make_xlsx(path, rows, title):
    """生成 Excel 数据表（行数控制大概页数）"""
    wb = Workbook()
    ws = wb.active
    ws.title = "数据表"
    ws.append([f"{title}"])
    ws.append(["序号", "项目名称", "规格", "数量", "单位", "备注"])
    for i in range(1, rows + 1):
        ws.append([i, f"项目{i:03d}", f"规格A-{i}", (i * 3) % 50 + 1, "件", ""])
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 8
    ws.column_dimensions["F"].width = 12
    wb.save(str(path))
    print(f"  ✓ {path.name}  ({rows} 行数据)")

# ── Word 案例：覆盖 奇/偶/单页/多页 ──
print("生成 Word 案例…")
make_docx(
    OUT / "作业A_单页报告.docx", 1,
    "作业A · 实验报告",
    ["· 实验目的：验证基本测量方法", "· 实验步骤：按标准流程操作三次",
     "· 实验结果：数据在允许误差范围内", "· 结论：实验达到预期目标"],
)
make_docx(
    OUT / "作业B_双页总结.docx", 2,
    "作业B · 课程总结",
    ["· 本周学习了基础理论", "· 完成了课后练习 1-20 题",
     "· 与同学讨论了疑难问题", "· 整理了课堂笔记"],
)
make_docx(
    OUT / "作业C_三页方案.docx", 3,
    "作业C · 项目方案",
    ["· 背景分析：市场需求持续增长", "· 目标设定：三个月内完成原型",
     "· 实施计划：分三个阶段推进", "· 风险评估：技术风险可控",
     "· 资源需求：人力 3 人，预算 5 万"],
)
make_docx(
    OUT / "作业D_四页论文.docx", 4,
    "作业D · 小论文",
    ["· 引言：研究背景与意义", "· 文献综述：国内外研究现状",
     "· 研究方法：定量分析为主", "· 实验设计：对照实验",
     "· 数据分析：SPSS 统计", "· 结论与展望"],
)
make_docx(
    OUT / "作业E_五页手册.docx", 5,
    "作业E · 操作手册",
    ["· 概述：适用范围与目的", "· 准备工作：工具与材料清单",
     "· 操作步骤：详细分解流程", "· 注意事项：安全规范",
     "· 常见问题：故障排除", "· 维护保养：定期检查"],
)
make_docx(
    OUT / "作业F_七页汇编.docx", 7,
    "作业F · 资料汇编",
    ["· 第一部分：基础概念", "· 第二部分：进阶方法",
     "· 第三部分：案例分析", "· 第四部分：实操练习",
     "· 第五部分：自测题", "· 第六部分：参考答案",
     "· 第七部分：延伸阅读"],
)

# ── Excel 案例 ──
print("生成 Excel 案例…")
make_xlsx(OUT / "数据表G_清单.xlsx", 30, "物资清单")
make_xlsx(OUT / "数据表H_台账.xlsx", 80, "设备台账（内容较多，可能跨 2 页）")

# ── 预期结果说明 ──
readme = OUT / "★测试说明.txt"
readme.write_text("""PDF 奇偶拼接打印 — 测试案例说明
================================

本文件夹里的案例按「页数奇偶」覆盖了典型情况，用来验证双面/单面两种模式。

【文件清单】

  作业A_单页报告.docx      1 页（奇数）
  作业B_双页总结.docx      2 页（偶数）
  作业C_三页方案.docx      3 页（奇数）
  作业D_四页论文.docx      4 页（偶数）
  作业E_五页手册.docx      5 页（奇数）
  作业F_七页汇编.docx      7 页（奇数）
  数据表G_清单.xlsx       约 1 页（奇数，视打印设置可能不同）
  数据表H_台账.xlsx       约 2 页（偶数，视打印设置可能不同）

【怎么测】

1. 双击「启动PDF奇偶拼接打印.command」（Mac）或 .bat（Windows）
2. 全部选中上面这些文件，拖进窗口
3. 先选「双面打印」模式 → 点生成 → 看结果页数
4. 再选「单面打印」模式 → 点生成 → 对比页数

【预期结果（以 Word 件为准）】

双面模式（奇数页末尾自动补 1 张空白页，凑成偶数）：

  A(1) + B(2) + C(3) + D(4) + E(5) + F(7)
  = 1补1 + 2不补 + 3补1 + 4不补 + 5补1 + 7补1
  = 2 + 2 + 4 + 4 + 6 + 8
  = 总计 26 页（含 4 张空白页）

单面模式（不补白，直接拼接）：

  A(1) + B(2) + C(3) + D(4) + E(5) + F(7)
  = 总计 22 页（无空白页）

【为什么要补白】

双面打印机没有自动翻页时，如果一份作业是奇数页，
它的最后一页会印在下一份作业第一页的背面，两份作业串页。
补一张空白页垫在后面，下一份作业就从新的一页正面开始。

【Excel 的页数】

Excel 的页数取决于打印设置（缩放、页边距、纸张方向），
和 Word 的固定分页不同。如果你的 Excel 打印出来页数和预期不一样，
属于正常现象 —— 换成 Word 文件测就能得到精确页数。

【注意事项】

- 转换 Word/Excel 需要电脑装有 Microsoft Office 或 LibreOffice
- PDF 文件可以直接拼接，不需要 Office
- Windows 首次运行可能需要允许「自动化控制」权限
""", encoding="utf-8")
print(f"  ✓ {readme.name}")

print(f"\n全部生成完毕 → {OUT}")
print("可以开始实验了！")
