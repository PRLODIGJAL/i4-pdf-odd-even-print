#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF 奇偶拼接打印工具 v1.0 —— 环境自检与启动器（Windows / macOS 共用这一份）。

双击 .bat / .command 会调到这里，由它统一负责三件事：
  1. 检查工具自带的虚拟环境还能不能用；
  2. 不能用就现场重建（建 venv → 装依赖），全程打印中文进度；
  3. 就绪后用虚拟环境里的解释器拉起 i4_app.py。

**打印格式统一走 bootui.py**（与项目2/3 同一份，逐字节一致）：
标记只有 `[ OK ] [ !! ] [ -- ] [ .. ]` 四种，分节只用 `-- 名字 ----` 一行，
宽度固定 62 列。**别在这里直接 print** —— 一直接 print，几个项目的启动界面
就各长各的，那正是使用者抱怨的「文字部分千奇百怪」。

为什么必须有这个文件
--------------------
虚拟环境**不能跨机复制**。它里面记着创建时那台机器的 Python 绝对路径，
二进制也不通用（macOS 是 .so、Windows 是 .pyd）。把这个文件夹从一台电脑
拷到另一台，双击只会得到一句看不懂的报错。与其让使用者去查，不如让它自己修。

用法
----
    python setup_env.py                 自检 → 按需安装 → 启动（正常双击走这条）
    python setup_env.py --check         只自检并打印结论，不安装、不启动
    python setup_env.py --setup         只把环境装好，不启动界面
    python setup_env.py --rebuild       不管现状，强制重建环境（可配 --setup 一起用）

【软件更新后要不要重装】不用。装好的东西在本文件夹的 .venv/ 里，
更新只是换掉几个 .py —— 别连 .venv 一起删就行。
"""

import os
import shutil
import subprocess
import sys

import bootui as B

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VENV = os.path.join(ROOT, ".venv")
APP = os.path.join(ROOT, "app", "ui", "i4_app.py")

TOOL_NAME = "PDF 奇偶拼接打印工具"
TOOL_VERSION = "2.0"

# ── 依赖清单：**按平台分两张表** ──────────────────────────────────────────
#   pymupdf     —— 两边都要：PDF 页数 / 补白 / 拼接。
#   tkinterdnd2 —— 两边都要：「把文件拖进窗口」。
#                  拖放是 tkinter **自己没有**的能力，靠这个库补。
#                  ⚠️ 界面里它是**可选**的（没装上也能用，只是得点按钮挑文件）。
#   pywin32     —— **只有 Windows** 的「Office 转 PDF」才用得到。它提供
#                  `win32com` / `pythoncom` 两个模块，靠它们在本机拉起 Word / Excel
#                  把文档导出成 PDF。macOS 那边走 AppleScript，用不上它。
#
# ⚠️ 依赖清单改了，`check_imports()` 的 import 串也要跟着改 —— 两边一起看
#    `need_list()` 与 `check_imports()`：一个装、一个验，脱节了就会出现
#    「装完仍报不可用」。
BASE_NEED = ("pymupdf>=1.24", "tkinterdnd2>=0.3")
WIN_NEED = ("pywin32>=306",)
# tkinter 是标准库自带的界面，不属于 pip 装的，但也必须在自检里 import 一次。
STD_IMPORTS = ("tkinter",)
# 各平台额外要 import 的第三方模块（与上面的依赖表一一对应）
PLATFORM_IMPORTS = {
    "nt": ("pythoncom", "win32com.client"),
    "posix": (),
}


def need_list():
    """这台机器上要装的依赖 —— Windows 多一份 pywin32。"""
    return list(BASE_NEED) + (list(WIN_NEED) if os.name == "nt" else [])


def check_imports():
    """自检要 import 的东西，同样按平台拼。

    ⚠️ 往 BASE_NEED / WIN_NEED 加依赖，这里就要补一行 —— 一个装、一个验，
       脱节了就会出现「装完仍报不可用」。
    """
    key = "nt" if os.name == "nt" else "posix"
    mods = (["pymupdf", "tkinterdnd2"] + list(PLATFORM_IMPORTS.get(key, ()))
            + list(STD_IMPORTS))
    return "import " + ", ".join(mods)

# 国内直连 PyPI 常常很慢，先走清华镜像；失败再退回官方源。
MIRRORS = (
    ("清华镜像", "https://pypi.tuna.tsinghua.edu.cn/simple"),
    ("官方源", None),
)


def venv_python(venv_dir):
    """虚拟环境里的解释器路径（Windows 在 Scripts 下，其余在 bin 下）。"""
    if os.name == "nt":
        return os.path.join(venv_dir, "Scripts", "python.exe")
    return os.path.join(venv_dir, "bin", "python")


def run(args, **kw):
    """跑一个子进程并把输出接回来。返回 (成功?, 输出文本)。"""
    try:
        p = subprocess.run(args, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", **kw)
    except OSError as e:
        return False, str(e)
    out = (p.stdout or "") + (p.stderr or "")
    return p.returncode == 0, out


def probe(py):
    """这个解释器能不能跑起来、装的依赖全不全、它是不是本环境的解释器。

    返回 (可用?, 说明)。只认「能 import 全部依赖」且「sys.prefix 就是它自己
    所在的虚拟环境目录」—— 后者能挡住「从别的电脑拷过来的 venv」。
    """
    if not py or not os.path.lexists(py):
        return False, "解释器不存在"
    code = (check_imports() + ";import sys;print('PREFIX='+sys.prefix)")
    ok, out = run([py, "-c", code])
    if not ok:
        first = ""
        for ln in out.strip().splitlines():
            if ln.strip() and "PREFIX=" not in ln:
                first = ln.strip()
        return False, ("依赖缺失或无法运行" + ("（" + first[:110] + "）" if first else ""))
    prefix = ""
    for ln in out.splitlines():
        if ln.startswith("PREFIX="):
            prefix = ln[7:].strip()
    if os.path.normcase(os.path.realpath(prefix)) != os.path.normcase(os.path.realpath(VENV)):
        return False, "环境来自另一台电脑（内部路径指向 %s）" % (prefix or "?")
    return True, "就绪"


# ── 界面库（Tk）能不能**真的**起来 —— 单独一道探测 ────────────────────────
#
# ⭐ 为什么不能「import tkinter 成功」就算过：**import 成功 ≠ 能弹出窗口**。
#    Tcl 是按「可执行文件旁边的 ../lib/tclX.Y」去找 init.tcl 的，而 venv
#    **不会**把 Tcl/Tk 的 library 目录复制进去 —— 于是它跑去 .venv/lib/tcl9.0
#    里翻，那儿什么都没有，初始化当场失败。
# ⚠️ 探测必须用 `tkinter.Tk()`，**不能用 `tkinter.Tcl()`** —— 实测踩过：
#    `Tcl()` 只起 Tcl 解释器，压根不走 Tk 的那套库查找，明明弹不出窗口它
#    照样报「成功」，等于白探。
TK_PROBE = ("import tkinter\n"
            "r = tkinter.Tk()\n"
            "try:\n"
            "    r.attributes('-alpha', 0.0)\n"
            "except Exception:\n"
            "    pass\n"
            "r.withdraw()\n"
            "r.destroy()\n"
            "print('TKPROBE=ok')\n")


def probe_tk(py):
    """这个解释器能不能把 Tk 起起来。返回 (行不行, 说明)。"""
    ok, out = run([py, "-c", TK_PROBE])
    if ok and "TKPROBE=ok" in out:
        return True, "能起 Tk"
    for ln in out.strip().splitlines():
        ln = ln.strip()
        if "init.tcl" in ln or "TclError" in ln or "No module named" in ln:
            return False, ln[:130]
    return False, "Tk 起不来"


def fix_tk_library(py):
    """给 venv 补上 Tcl/Tk 的库目录（软链到**它自己那个** base 解释器那份）。"""
    if os.name == "nt":
        return False
    ok, out = run([py, "-c", "import sys;print('BASEPREFIX='+sys.base_prefix)"])
    if not ok:
        return False
    base = ""
    for ln in out.splitlines():
        if ln.startswith("BASEPREFIX="):
            base = ln[len("BASEPREFIX="):].strip()
    if not base:
        return False
    src_lib = os.path.join(base, "lib")
    dst_lib = os.path.join(VENV, "lib")
    if not (os.path.isdir(src_lib) and os.path.isdir(dst_lib)):
        return False
    try:
        names = sorted(os.listdir(src_lib))
    except OSError:
        return False
    made = []
    for name in names:
        if not (name.startswith("tcl") or name.startswith("tk")):
            continue
        src, dst = os.path.join(src_lib, name), os.path.join(dst_lib, name)
        if not os.path.isdir(src) or os.path.lexists(dst):
            continue
        try:
            os.symlink(src, dst)
            made.append(name)
        except OSError:
            continue
    if made:
        B.note("给虚拟环境补上 Tcl/Tk 库目录：%s（venv 不带 Tcl 库，不补界面就起不来）"
               % "、".join(made))
    return bool(made)


def drop_venv():
    """删掉现有的虚拟环境目录，返回「现在这个位置是不是干净的」。

    只删那种确实像 venv 的（有 pyvenv.cfg，或有 bin/Scripts 目录，
    或解释器路径存在——含断链）；万一是用户自己放的、恰好叫 .venv 的普通
    文件夹，就不动它，并如实说明，免得误删。
    """
    if not os.path.lexists(VENV):
        return True
    marks = (os.path.join(VENV, "pyvenv.cfg"),
             os.path.join(VENV, "bin"),
             os.path.join(VENV, "Scripts"),
             venv_python(VENV))
    if not any(os.path.lexists(m) for m in marks):
        B.item(B.FAIL, "旧环境", "它不像虚拟环境，没有删除")
        B.note("请手动删掉这个文件夹后重试：%s" % VENV)
        return False
    B.step("删除旧环境 .venv")
    try:
        shutil.rmtree(VENV)
    except OSError as e:
        B.item(B.FAIL, "旧环境", "删不掉：%s" % e)
        B.note("请手动删掉这个文件夹后重试：%s" % VENV)
        return False
    return True


def make_venv(base):
    B.step("创建虚拟环境 .venv（用 %s）" % base)
    ok, out = run([base, "-m", "venv", VENV])
    if not ok:
        B.item(B.FAIL, "虚拟环境", "创建失败")
        B.note(out.strip()[:600])
        return False
    return True


def install_deps(py):
    """装依赖。逐个镜像源试，任何一个成功即返回。

    ⚠️ 装什么由 `need_list()` 按平台决定（Windows 会多一个 pywin32）——
       这里别再写死包名，否则又一个「装的和验的不是同一份」的地方。
    """
    pkgs = need_list()
    for name, url in MIRRORS:
        args = [py, "-m", "pip", "install", "--disable-pip-version-check",
                "--no-input"] + pkgs
        if url:
            args += ["-i", url]
        B.step("正在安装依赖（%s）：%s" % (name, "、".join(pkgs)))
        B.note("首次约 1～3 分钟，请勿关窗口")
        ok, out = run(args)
        if ok:
            B.item(B.OK, "依赖", "安装完成（%s）" % name)
            return True
        tail = [ln for ln in out.strip().splitlines() if ln.strip()][-6:]
        B.item(B.FAIL, "依赖", "用「%s」没装上，末尾几行：" % name)
        for ln in tail:
            B.note(ln[:150])
        B.note("换一个源再试…")
    return False


def version_code():
    """`--check` 打印各依赖版本用的小脚本（包名从 need_list 派生，不写死）。"""
    names = [p.replace(">=", "=").split("=")[0] for p in need_list()]
    return ("import importlib.metadata as md\n"
            "for n in %r:\n"
            "    try:\n"
            "        print(f'{n} {md.version(n)}')\n"
            "    except Exception:\n"
            "        print(f'{n} （读不出版本）')" % (names,))


def dep_versions(py):
    """「依赖」那一行的值：`pymupdf 1.24.0 · tkinterdnd2 0.4.2`。"""
    ok, out = run([py, "-c", version_code()])
    if not ok:
        return "、".join(p.split(">=")[0] for p in need_list())
    items = []
    for ln in out.strip().splitlines():
        ln = ln.strip()
        if ln:
            items.append(ln.replace("（读不出版本）", "").strip())
    return " · ".join(items) or "（读不出版本）"


def ready_rows(py, tk_ok=True, tk_why="能起 Tk"):
    """环境就绪时的固定三行 —— 每条路都长这样，方便扫。"""
    B.item(B.OK, "虚拟环境", os.path.relpath(VENV, ROOT))
    B.item(B.OK, "依赖", dep_versions(py))
    B.item(B.OK if tk_ok else B.FAIL, "界面库", tk_why)


def ensure_ready(force=False):
    """确保有一个可用环境。返回可用解释器路径，失败返回 None。"""
    py = venv_python(VENV)

    if not force:
        ok, why = probe(py)
        if ok:
            tk_ok, tk_why = probe_tk(py)
            if tk_ok:
                ready_rows(py, tk_ok, tk_why)
                return py
            B.item(B.FAIL, "界面库", tk_why)
            if fix_tk_library(py):
                tk_ok, _ = probe_tk(py)
                if tk_ok:
                    B.item(B.OK, "界面库", "已补上 Tcl/Tk 库，可以用了")
                    return py
            B.note("补不上，需要重建一次环境（只重建本文件夹里的 .venv，不碰你的文件）。")
        elif os.path.lexists(py) or os.path.lexists(VENV):
            B.item(B.FAIL, "虚拟环境", why)
            B.note("需要重建一次（只重建本文件夹里的 .venv，不碰你的文件）。")
        else:
            B.item(B.SKIP, "虚拟环境", "还没装过，这一次先把环境备好")
        B.section("", "准备环境")

    base = sys.executable
    B.step("本机 Python %d.%d.%d（%s）" % (sys.version_info[0], sys.version_info[1],
                                          sys.version_info[2], base))
    if sys.version_info < (3, 9):
        B.item(B.FAIL, "本机 Python", "版本太低，需要 3.9 或更高")
        return None

    # 走到这里说明现有环境不可用（或用户要求强制重建）：**一律先清干净**。
    if not drop_venv():
        return None

    if not make_venv(base):
        return None

    py = venv_python(VENV)
    if not os.path.exists(py):
        B.item(B.FAIL, "虚拟环境", "建好了但找不到解释器")
        return None

    # pip 太旧会装不上新包；升级失败不致命，继续往下走。
    B.step("升级 pip")
    run([py, "-m", "pip", "install", "--upgrade", "pip",
         "--disable-pip-version-check", "--no-input", "-q"])

    if not install_deps(py):
        B.item(B.FAIL, "依赖", "没装上")
        B.note("常见原因：网络不通或被公司网络策略挡住 —— 换个网络再试；"
               "装的是 Microsoft Store 版 Python —— 建议到 python.org 重装")
        return None

    B.step("复核装好没有")
    ok, why = probe(py)
    if not ok:
        B.item(B.FAIL, "可用性", "装完仍然不可用 —— %s" % why)
        return None
    tk_ok, tk_why = probe_tk(py)
    ready_rows(py, tk_ok, tk_why)
    return py if tk_ok else None


def launch(py):
    """用虚拟环境的解释器拉起界面。"""
    if not os.path.exists(APP):
        B.end(B.FAIL, "找不到 i4_app.py",
              steps=["确认这个文件夹是完整的（i4_app.py 要和本文件在同一层）",
                     "重新拷一份完整的交付包"])
        return 1
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = "utf-8"       # 免得 Windows 控制台用 cp936 打印报错
    env["PYTHONUTF8"] = "1"
    B.end(B.OK, "环境就绪，正在打开界面…（关掉界面窗口即可退出）")
    try:
        return subprocess.run([py, APP], cwd=HERE, env=env).returncode
    except KeyboardInterrupt:
        return 0


def do_check():
    """`--check`：只看不动，把结论一次说完。"""
    py = venv_python(VENV)
    exist = os.path.isdir(VENV)
    B.section("", "运行环境")
    B.item(B.OK if exist else B.SKIP, "虚拟环境",
           "%s（%s）" % (os.path.relpath(VENV, ROOT), "存在" if exist else "不存在"))
    ok, why = probe(py)
    if not ok:
        B.item(B.FAIL, "可用性", "不可用 —— %s" % why)
        B.end(B.FAIL, "环境还没准备好",
              steps=["双击「运行前的环境准备.bat」，它会当场把环境备好"]
              if os.name == "nt" else ["重新双击启动器，它会自动装好"])
        return 1
    B.item(B.OK, "依赖", dep_versions(py))
    tk_ok, tk_why = probe_tk(py)
    B.item(B.OK if tk_ok else B.FAIL, "界面库", tk_why)
    B.item(B.OK, "可用性", "可用")
    B.end(B.OK, "可以正常启动")
    return 0


def main():
    args = sys.argv[1:]
    B.head(TOOL_NAME, "v" + TOOL_VERSION, "",
           [("系统", B.os_line()),
            ("Python", "%d.%d.%d" % sys.version_info[:3]),
            ("位置", ROOT)])

    if "--check" in args:
        return do_check()

    B.section("", "运行环境")
    py = ensure_ready(force="--rebuild" in args)
    if not py:
        B.end(B.FAIL, "没能准备好环境",
              steps=["把这一屏截图发给我，我按情况处理",
                     "想重来一次就双击「运行前的环境准备.bat」"
                     if os.name == "nt" else "想重来一次就重新双击启动器"])
        return 1

    if "--setup" in args:
        B.end(B.OK, "环境已就绪，以后双击启动器就能直接用")
        return 0

    return launch(py)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(0)
