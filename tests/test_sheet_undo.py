# -*- coding: utf-8 -*-
"""「在线填表」撤销 / 重做 的回归测试。

为什么这么写：被测的那十几个函数是**每次现从 index.html 里抽出来的**，
所以改了 index.html 之后跑这个脚本，测的就是最新代码，
不会出现"测试和实际代码各说各话"的情况。

用法（在项目根目录）：
    python tests/test_sheet_undo.py

退出码 0 = 全部通过，1 = 有失败项。
"""
import io
import os
import re
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
INDEX = os.path.join(ROOT, "index.html")
CASES = os.path.join(HERE, "sheet_undo_cases.js")

# 要抽出来测的函数（按依赖顺序无所谓，都是 function 声明）
WANTED = ("colName", "normSel", "paintSel", "updateSheetFoot", "renderSheet",
          "pushSheetUndo", "updateSheetUndoBtn", "applySheetSnap", "sheetUndo", "sheetRedo")


def extract(js, name):
    """按花括号配对，把某个 function 的完整源码抠出来。"""
    try:
        i = js.index("function %s(" % name)
    except ValueError:
        raise SystemExit("✗ index.html 里找不到函数 %s —— 是改名了吗？" % name)
    depth = 0
    for j in range(i, len(js)):
        if js[j] == "{":
            depth += 1
        elif js[j] == "}":
            depth -= 1
            if depth == 0:
                return js[i:j + 1]
    raise SystemExit("✗ 函数 %s 的花括号没配平" % name)


def find_node():
    import shutil
    p = shutil.which("node")
    if p:
        return p
    for pat in (r"C:\Users\Administrator\.workbuddy\binaries\node\versions",):
        if os.path.isdir(pat):
            for v in sorted(os.listdir(pat), reverse=True):
                exe = os.path.join(pat, v, "node.exe")
                if os.path.isfile(exe):
                    return exe
    raise SystemExit("✗ 找不到 node，跑不了这个测试（装了 Node.js 就行）")


# 最小运行环境：让抽出来的 DOM 代码能在 node 里跑起来
STUB = """
let SHEET_COLS = 4, SHEET_ROWS = 3, SEL = null, lastSelKey = "";
let SHEET = [["a","","",""],["b","","",""],["c","","",""]];
let SHEET_UNDO = [], SHEET_REDO = [], SHEET_EDIT_KEY = "";
const SHEET_UNDO_MAX = 100;
const SHEET_COL_CN = ["公司","岗位","收件邮箱","备注"];
const SHEET_COL_LETTER = ["A","B","C","D"];
const esc = s => String(s == null ? "" : s);
const stubEl = { innerHTML: "", textContent: "", value: "", disabled: false, style: {} };
const $ = () => stubEl;
const document = { querySelector: () => null, querySelectorAll: () => [], activeElement: null };
function queueSheetSave() {}
let lastToast = "";
function toast(m) { lastToast = m; }
"""


def main():
    src = io.open(INDEX, encoding="utf-8").read()
    blocks = re.findall(r"<script[^>]*>(.*?)</script>", src, re.S)
    if not blocks:
        raise SystemExit("✗ index.html 里没有 <script> 块")
    js = "\n".join(blocks)

    parts = [STUB] + [extract(js, n) for n in WANTED]
    cases = io.open(CASES, encoding="utf-8").read()

    tmp = os.path.join(tempfile.gettempdir(), "sheet_undo_regression.js")
    io.open(tmp, "w", encoding="utf-8").write("\n".join(parts) + "\n" + cases)

    print("从 index.html 抽了 %d 个函数（%s）\n" % (len(WANTED), "、".join(WANTED)))
    r = subprocess.run([find_node(), tmp], capture_output=True)
    sys.stdout.write(r.stdout.decode("utf-8", "replace"))
    err = r.stderr.decode("utf-8", "replace").strip()
    if err:
        sys.stderr.write(err + "\n")
    return r.returncode


if __name__ == "__main__":
    sys.exit(main())
