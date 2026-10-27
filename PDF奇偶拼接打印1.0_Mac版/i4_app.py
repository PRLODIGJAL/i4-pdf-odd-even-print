# -*- coding: utf-8 -*-
"""
i4_app.py —— PDF 奇偶拼接打印 · 界面
=====================================
UI 按 i5《浅色系UI风格指南》：三级底色 · 1px 细线 · 饱和色 ≤5%
"""

from __future__ import annotations

import json
import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from i4_core import (  # noqa: E402
    ALL_EXTS, ConvItem, ConvStatus, PrintMode,
    collect_items, is_supported, run_pipeline,
)

# ── 「把文件拖进来」（同项目3 enable_dnd 模式）─────────────────────────
try:
    import tkinterdnd2 as _dnd  # type: ignore
    from tkinterdnd2.TkinterDnD import require as _dnd_require
    DND_FILES = _dnd.DND_FILES
except Exception:
    _dnd = None
    _dnd_require = None
    DND_FILES = "DND_Files"

_dnd_ok = False


def enable_dnd(root: tk.Tk) -> bool:
    global _dnd_ok
    if _dnd_require is None:
        return False
    try:
        _dnd_require(root)
        _dnd_ok = True
        return True
    except Exception:
        return False


def drop_paths(data: str, splitlist=None) -> list[str]:
    if splitlist:
        return list(splitlist)
    result, token, depth, quote = [], "", 0, False
    for ch in data:
        if ch == '"':
            quote = not quote
            token += ch
        elif ch == "{" and not quote:
            depth += 1
            token += ch
        elif ch == "}" and not quote:
            depth = max(0, depth - 1)
            token += ch
        elif ch in " \t" and not quote and depth == 0:
            if token:
                result.append(token.strip('"'))
                token = ""
        else:
            token += ch
    if token:
        result.append(token.strip('"'))
    return result

# ── i5 色板 ─────────────────────────────────────────────────────────
BG = "#f5f6f8"
CARD = "#ffffff"
DZ_BG = "#f7f9fc"
DZ_BG_ON = "#eaf1ff"
STATUS_BG = "#eceef1"
BORDER = "#d9dee5"
TEXT = "#2a2f36"
TEXT2 = "#7a828c"
MUTED = "#9aa1a9"
DISABLED = "#c3c8ce"
PRIMARY = "#2f6fed"
PRIMARY_D = "#2257c4"
WARN = "#b26a00"
DANGER = "#c0392b"
DANGER_BG = "#fdf1ef"

FONT = "PingFang SC" if sys.platform == "darwin" else "Microsoft YaHei UI"

WIN_W = 640
WIN_H = 540
MIN_W = 540
MIN_H = 440

# ── 输出路径记忆（同项目3 config 模式）──────────────────────────────
CONFIG_PATH = Path.home() / ".pdf_odd_even_print.json"


def _default_out() -> str:
    return str(Path.home() / "Desktop" / "合集打印.pdf")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 奇偶拼接打印")
        self.configure(bg=BG)
        self.geometry(f"{WIN_W}x{WIN_H}")
        self.minsize(MIN_W, MIN_H)
        self._set_icon()

        self.items: list[ConvItem] = []
        self.mode = tk.StringVar(value="duplex")
        self.out_path = tk.StringVar(value=_default_out())
        self.busy = False
        self._stop = threading.Event()
        self._save_after = None

        self._setup_style()
        self._build_ui()
        self._load_config()
        self.mode.trace_add("write", lambda *_: self._schedule_save())
        enable_dnd(self)
        self._setup_dnd()

    def _set_icon(self):
        """给窗口设一个自定义图标，让 Dock / 任务栏不显示 Python 默认图标。"""
        try:
            img = tk.PhotoImage(width=32, height=32)
            # 画一个蓝色圆角方块 + 白色文档折角
            for y in range(32):
                for x in range(32):
                    dx, dy = x - 16, y - 16
                    if dx * dx + dy * dy < 220:
                        img.put("#2f6fed", (x, y))
            for y in range(6, 26):
                for x in range(10, 22):
                    img.put("#ffffff", (x, y))
            for y in range(6, 14):
                for x in range(16, 22):
                    if (x - 16) >= (y - 6):
                        img.put("#2f6fed", (x, y))
            self.iconphoto(True, img)
        except Exception:
            pass

    def _setup_style(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure(".", font=(FONT, 11), background=BG, foreground=TEXT)
        st.configure("TLabel", font=(FONT, 11), background=BG, foreground=TEXT)
        st.configure("TFrame", background=BG)

        st.configure("Primary.TButton",
                     font=(FONT, 11, "bold"), padding=(14, 5),
                     background=PRIMARY, foreground="#ffffff")
        st.map("Primary.TButton",
               background=[("active", PRIMARY_D), ("disabled", DISABLED)])

        st.configure("Small.TButton",
                     font=(FONT, 10), padding=(6, 2),
                     background=CARD, foreground=TEXT,
                     bordercolor=BORDER, borderwidth=1)
        st.map("Small.TButton",
               background=[("active", "#f0f2f5"), ("pressed", "#e6e9ee"),
                           ("disabled", BG)],
               foreground=[("disabled", DISABLED)])

        st.configure("Stop.TButton",
                     font=(FONT, 10), padding=(6, 2),
                     background=CARD, foreground=DANGER,
                     bordercolor=BORDER, borderwidth=1)
        st.map("Stop.TButton", background=[("active", DANGER_BG)])

        st.configure("Title.TLabel", font=(FONT, 12, "bold"), background=BG)
        st.configure("Sub.TLabel", font=(FONT, 10), background=BG, foreground=TEXT2)
        st.configure("Status.TLabel", font=(FONT, 10), background=STATUS_BG,
                     foreground=TEXT2)
        st.configure("i4.Horizontal.TProgressbar",
                     troughcolor="#e6e9ee", background=PRIMARY, thickness=6)

    def _build_ui(self):
        # 标题
        hdr = tk.Frame(self, bg=BG)
        hdr.pack(fill="x", padx=14, pady=(12, 4))
        ttk.Label(hdr, text="PDF 奇偶拼接打印", style="Title.TLabel").pack(side="left")
        ttk.Label(hdr, text="Word / Excel → 合集 PDF · 自动补空白页",
                  style="Sub.TLabel").pack(side="left", padx=(10, 0))

        # 输入区
        self.dz = tk.Frame(self, bg=DZ_BG, highlightthickness=1,
                           highlightbackground=BORDER)
        self.dz.pack(fill="both", expand=True, padx=14, pady=(4, 6))

        top_row = tk.Frame(self.dz, bg=DZ_BG)
        top_row.pack(fill="x", padx=10, pady=(10, 2))
        self.lbl_hint = tk.Label(
            top_row,
            text="⬇  拖入 Word / Excel 文件 · 按拖入顺序拼接",
            font=(FONT, 10), bg=DZ_BG, fg=MUTED, anchor="w")
        self.lbl_hint.pack(side="left", fill="x", expand=True)

        btn_row = tk.Frame(self.dz, bg=DZ_BG)
        btn_row.pack(fill="x", padx=10, pady=(2, 6))
        ttk.Button(btn_row, text="浏览…", style="Small.TButton",
                   command=self._on_browse).pack(side="left", padx=(0, 5))
        ttk.Button(btn_row, text="清空", style="Small.TButton",
                   command=self._on_clear).pack(side="left")

        list_frame = tk.Frame(self.dz, bg=DZ_BG)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.listbox = tk.Listbox(
            list_frame, font=(FONT, 10), bg=CARD, fg=TEXT,
            selectbackground=DZ_BG_ON, selectforeground=TEXT,
            highlightthickness=1, highlightbackground=BORDER,
            borderwidth=0, activestyle="none")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 模式
        mode_frame = tk.Frame(self, bg=BG)
        mode_frame.pack(fill="x", padx=14, pady=(0, 4))
        ttk.Label(mode_frame, text="打印模式：").pack(side="left")
        self._mk_radio(mode_frame, "双面打印（奇数页补空白）",
                       "duplex").pack(side="left", padx=(0, 14))
        self._mk_radio(mode_frame, "单面打印（不补白）",
                       "simplex").pack(side="left")

        # 操作行
        op_frame = tk.Frame(self, bg=BG)
        op_frame.pack(fill="x", padx=14, pady=(4, 6))
        self.btn_run = ttk.Button(op_frame, text="开始拼接",
                                  style="Primary.TButton", command=self._on_run)
        self.btn_run.pack(side="left", padx=(0, 6))
        self.btn_stop = ttk.Button(op_frame, text="停止",
                                   style="Stop.TButton", command=self._on_stop,
                                   state="disabled")
        self.btn_stop.pack(side="left", padx=(0, 6))

        out_row = tk.Frame(op_frame, bg=BG)
        out_row.pack(side="left", fill="x", expand=True, padx=(10, 0))
        ttk.Label(out_row, text="输出：", font=(FONT, 10),
                  foreground=TEXT2).pack(side="left")
        self.ent_out = tk.Entry(
            out_row, textvariable=self.out_path,
            font=(FONT, 10), bg=CARD, fg=TEXT,
            highlightthickness=1, highlightbackground=BORDER, borderwidth=0)
        self.ent_out.pack(side="left", fill="x", expand=True, padx=(4, 4))
        ttk.Button(out_row, text="选择…", style="Small.TButton",
                   command=self._on_pick_out).pack(side="left")

        # 进度条
        prog_frame = tk.Frame(self, bg=BG)
        prog_frame.pack(fill="x", padx=14, pady=(0, 4))
        self.progress = ttk.Progressbar(prog_frame, mode="determinate",
                                        style="i4.Horizontal.TProgressbar")
        self.progress.pack(fill="x")

        # 状态栏
        sf = tk.Frame(self, bg=STATUS_BG)
        sf.pack(fill="x", side="bottom")
        self.lbl_status = tk.Label(
            sf, text="把 Word / Excel 拖进上方那块（或点「浏览…」）",
            font=(FONT, 10), bg=STATUS_BG, fg=TEXT2, anchor="w",
            padx=14, pady=6)
        self.lbl_status.pack(fill="x")

    def _mk_radio(self, parent, text, value):
        return tk.Radiobutton(
            parent, text=text, variable=self.mode, value=value,
            font=(FONT, 10), bg=BG, fg=TEXT,
            activebackground=BG, activeforeground=TEXT,
            selectcolor=CARD, bd=0, highlightthickness=0, anchor="w")

    def _setup_dnd(self):
        if not _dnd_ok:
            return
        for w in (self.dz, self):
            if not hasattr(w, "drop_target_register"):
                continue
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    def _on_drop(self, event):
        self._add_files([Path(p) for p in drop_paths(
            getattr(event, "data", ""), getattr(event, "splitlist", None))])

    def _on_browse(self):
        exts = " ".join(f"*{e}" for e in sorted(ALL_EXTS))
        files = filedialog.askopenfilenames(
            title="选择 Word / Excel 文件",
            filetypes=[("Word / Excel", exts), ("所有文件", "*.*")])
        if files:
            self._add_files([Path(f) for f in files])

    def _add_files(self, paths: list[Path]):
        for p in paths:
            if is_supported(p) and not any(it.src == p for it in self.items):
                self.items.append(ConvItem(src=p))
        self._refresh_list()

    def _on_clear(self):
        self.items.clear()
        self._refresh_list()
        self._set_status("已清空。把 Word / Excel 拖进上方那块")

    def _refresh_list(self):
        self.listbox.delete(0, "end")
        for it in self.items:
            self.listbox.insert("end", it.src.name)
        n = len(self.items)
        if n:
            self._set_status(f"已选 {n} 份文件 · 选好模式后点「开始拼接」")

    def _on_pick_out(self):
        p = filedialog.asksaveasfilename(
            title="保存合集 PDF", defaultextension=".pdf",
            initialdir=str(Path.home() / "Desktop"),
            filetypes=[("PDF", "*.pdf")])
        if p:
            self.out_path.set(p)
            self._schedule_save()

    def _load_config(self):
        if not CONFIG_PATH.is_file():
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            return
        if not isinstance(data, dict):
            return
        out = data.get("out_path", "")
        if out and Path(out).parent.is_dir():
            self.out_path.set(out)
        mode = data.get("mode", "")
        if mode in ("duplex", "simplex"):
            self.mode.set(mode)

    def _save_config(self):
        data = {"out_path": self.out_path.get(), "mode": self.mode.get()}
        try:
            CONFIG_PATH.write_text(
                json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
        except Exception:
            pass

    def _schedule_save(self):
        if self._save_after:
            try:
                self.after_cancel(self._save_after)
            except Exception:
                pass
        self._save_after = self.after(600, self._save_config)

    def _on_run(self):
        if self.busy or not self.items:
            if not self.items:
                self._set_status("请先添加文件", warn=True)
            return
        out = self.out_path.get().strip()
        if not out:
            out = _default_out()
            self.out_path.set(out)
        self._save_config()
        mode = PrintMode.DUPLEX if self.mode.get() == "duplex" else PrintMode.SIMPLEX
        self.busy = True
        self._stop.clear()
        self.btn_run.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.progress["value"] = 0
        threading.Thread(target=self._run_worker,
                         args=(mode, Path(out)), daemon=True).start()

    def _run_worker(self, mode: PrintMode, out: Path):
        def progress(done, total, msg):
            if self._stop.is_set():
                return
            self.after(0, lambda: self._on_progress(done, total, msg))
        ok, msg, plan = run_pipeline(self.items, mode, out, progress=progress)
        self.after(0, lambda: self._on_done(ok, msg, plan))

    def _on_progress(self, done, total, msg):
        if total > 0:
            self.progress["value"] = (done / total) * 100
        self._set_status(msg)

    def _on_done(self, ok, msg, plan):
        self.busy = False
        self.btn_run.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self.progress["value"] = 100 if ok else 0
        if ok:
            pad = f" · 补了 {plan.blank_added} 页空白" if plan.blank_added else ""
            self._set_status(f"✓ {msg}{pad}")
        else:
            self._set_status(f"✗ {msg}", warn=True)

    def _on_stop(self):
        self._stop.set()
        self._set_status("正在停止…")

    def _set_status(self, text, warn=False):
        self.lbl_status.configure(text=text, fg=WARN if warn else TEXT2)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
