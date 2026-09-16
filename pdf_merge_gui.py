# -*- coding: utf-8 -*-
"""
PDF / 图片 合并小工具（本地图形界面）

用法：
    双击「合并PDF.bat」打开。
    点「添加文件…」把 PDF 和图片加进来 → 用「上移 / 下移」排好顺序 → 点「合并并保存…」。
    列表里的顺序，就是合并后 PDF 的页面顺序。

    选中图片后点「图片设置…」，可以调旋转、页面大小、缩放方式。
    点「预览…」可以按当前顺序把每一页都画出来，先看一眼再合并。

说明：
    纯本地处理，文件不上传任何服务器。
    依赖 pypdf + Pillow，PDF 缩略图用 pypdfium2。不改原文件，只是按顺序把内容拼成一个新 PDF。
"""
import os
import queue
import shutil
import sys
import tempfile
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "PDF / 图片 合并工具"
FONT = ("Microsoft YaHei UI", 10)
FONT_BOLD = ("Microsoft YaHei UI", 10, "bold")
FONT_BIG = ("Microsoft YaHei UI", 12, "bold")
GREY = "#7a7a7a"
BLUE = "#1a6fd4"
GREEN = "#1e8449"
RED = "#c0392b"

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif", ".tif", ".tiff"}
FILE_TYPES = [
    ("PDF 和图片", "*.pdf *.jpg *.jpeg *.png *.bmp *.webp *.gif *.tif *.tiff"),
    ("PDF 文件", "*.pdf"),
    ("图片文件", "*.jpg *.jpeg *.png *.bmp *.webp *.gif *.tif *.tiff"),
    ("全部文件", "*.*"),
]

# A4 纸的尺寸，单位「点」（1 点 = 1/72 英寸）
A4_W, A4_H = 595.28, 841.89
OUT_DPI = 144          # 图片放进 PDF 时的分辨率。144 够清晰，文件也不会太大

ROTATE_VALUES = [0, 90, 180, 270]
PAGE_VALUES = ["auto", "a4p", "a4l"]
PAGE_LABELS = {"auto": "跟随图片", "a4p": "A4 纵向", "a4l": "A4 横向"}
FIT_VALUES = ["contain", "cover"]
FIT_LABELS = {"contain": "完整显示（周围留白）", "cover": "铺满页面（边缘可能被裁掉）"}

# 预览窗口的参数
THUMB_W = 200        # 缩略图宽度（像素）
THUMB_H = 283        # 缩略图框的高度，按 A4 竖版比例来，这样网格能对齐
THUMB_COLS = 4       # 预览里每行放几张
ZOOM_W = 860         # 点开放大后，那一页渲染多宽
FONT_SMALL = ("Microsoft YaHei UI", 8)


# ==================== 小函数 ====================
def human_size(n):
    """把字节数变好看，比如 604101 -> 590.0 KB"""
    n = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def is_pdf(path):
    return os.path.splitext(path)[1].lower() == ".pdf"


def is_image(path):
    return os.path.splitext(path)[1].lower() in IMAGE_EXTS


def default_img_opts():
    return {"rotate": 0, "page": "auto", "fit": "contain"}


def read_pdf_pages(path):
    """读 PDF 页数。读不了（加密/损坏）返回 None。"""
    try:
        from pypdf import PdfReader
        reader = PdfReader(path)
        if getattr(reader, "is_encrypted", False):
            try:
                reader.decrypt("")
            except Exception:
                pass
        return len(reader.pages)
    except Exception:
        return None


def read_image_size(path):
    """读图片宽高（先按 EXIF 摆正）。读不了返回 None。"""
    try:
        from PIL import Image, ImageOps
        im = Image.open(path)
        im.load()
        fixed = ImageOps.exif_transpose(im)
        if fixed is not None:
            im = fixed
        return im.size   # (宽, 高)
    except Exception:
        return None


def rotate_image(im, degrees):
    """按顺时针方向旋转 90 的倍数。"""
    from PIL import Image
    d = degrees % 360
    if d == 90:
        return im.transpose(Image.ROTATE_270)     # PIL 的 270 是逆时针，正好等于顺时针 90
    if d == 180:
        return im.transpose(Image.ROTATE_180)
    if d == 270:
        return im.transpose(Image.ROTATE_90)
    return im


def render_image_page(path, opts):
    """
    把一张图片按设置画到一页白纸上。
    返回 (PIL 图片, 输出 DPI)。
    """
    from PIL import Image, ImageOps

    im = Image.open(path)
    im.load()

    # 1) 手机拍的图带方向信息，先按它摆正
    fixed = ImageOps.exif_transpose(im)
    if fixed is not None:
        im = fixed

    # 2) 原图自带的 DPI（「跟随图片」时要用），旋转缩放前先记下来
    try:
        src_dpi = float(im.info.get("dpi", (96, 96))[0])
        if not 20 < src_dpi < 1200:
            src_dpi = 96.0
    except Exception:
        src_dpi = 96.0

    # 3) 透明背景（PNG）先垫成白的
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")

    # 4) 旋转
    im = rotate_image(im, opts.get("rotate", 0))

    # 5) 这一页多大（单位：点）
    page = opts.get("page", "auto")
    if page == "a4p":
        pw, ph = A4_W, A4_H
    elif page == "a4l":
        pw, ph = A4_H, A4_W
    else:
        pw = im.width * 72.0 / src_dpi
        ph = im.height * 72.0 / src_dpi

    # 6) 换算成画布像素
    k = OUT_DPI / 72.0
    cw = max(1, int(round(pw * k)))
    ch = max(1, int(round(ph * k)))
    canvas = Image.new("RGB", (cw, ch), "white")

    # 7) 缩放：contain = 整张放得下；cover = 铺满，多出来的裁掉
    if opts.get("fit", "contain") == "cover":
        ratio = max(cw / im.width, ch / im.height)
    else:
        ratio = min(cw / im.width, ch / im.height)
    nw = max(1, int(round(im.width * ratio)))
    nh = max(1, int(round(im.height * ratio)))
    if (nw, nh) != im.size:
        im = im.resize((nw, nh), Image.LANCZOS)

    # 8) 贴到正中间
    canvas.paste(im, ((cw - nw) // 2, (ch - nh) // 2))

    # 9) 输出的 DPI 反推一下，让页面宽度精确等于目标值。
    #    直接用固定的 144 会因为取整让 A4 变成 595.5 点，差 0.2 点（约 0.07 毫米）。
    return canvas, cw * 72.0 / pw


# ==================== 图片设置窗口 ====================
class ImageSettingsDialog(tk.Toplevel):
    def __init__(self, master, app, paths):
        super().__init__(master)
        self.app = app
        self.paths = list(paths)
        self._photo = None          # 防止预览图被垃圾回收
        self._after_id = None

        self.title(f"图片设置 — 共 {len(self.paths)} 张")
        self.resizable(False, False)
        self.transient(master)

        rv = tk.IntVar(value=app.img_opts[self.paths[0]]["rotate"])
        pv = tk.StringVar(value=app.img_opts[self.paths[0]]["page"])
        fv = tk.StringVar(value=app.img_opts[self.paths[0]]["fit"])
        self.rv, self.pv, self.fv = rv, pv, fv

        body = tk.Frame(self)
        body.pack(fill="both", expand=True, padx=16, pady=16)

        # ---- 左：预览 ----
        left = tk.Frame(body)
        left.pack(side="left", fill="y")
        tk.Label(left, text="效果预览", font=FONT_BOLD).pack(anchor="w")
        self.preview = tk.Label(left, bg="#e9e9e9", width=34, height=22,
                                relief="solid", bd=1)
        self.preview.pack(pady=(6, 6))
        self.preview_note = tk.Label(left, text="", font=FONT, fg=GREY, wraplength=300,
                                     justify="left")
        self.preview_note.pack(anchor="w")

        # ---- 右：选项 ----
        right = tk.Frame(body)
        right.pack(side="left", fill="both", expand=True, padx=(20, 0))

        tk.Label(right, text="旋转", font=FONT_BOLD).pack(anchor="w")
        rf = tk.Frame(right)
        rf.pack(anchor="w", pady=(4, 12))
        for v in ROTATE_VALUES:
            tk.Radiobutton(rf, text=f"{v}°", font=FONT, variable=rv, value=v,
                           command=self.refresh).pack(side="left", padx=(0, 12))

        tk.Label(right, text="页面大小", font=FONT_BOLD).pack(anchor="w")
        pf = tk.Frame(right)
        pf.pack(anchor="w", pady=(4, 12))
        for v in PAGE_VALUES:
            tk.Radiobutton(pf, text=PAGE_LABELS[v], font=FONT, variable=pv, value=v,
                           command=self.refresh).pack(anchor="w")

        tk.Label(right, text="缩放方式", font=FONT_BOLD).pack(anchor="w")
        ff = tk.Frame(right)
        ff.pack(anchor="w", pady=(4, 12))
        for v in FIT_VALUES:
            tk.Radiobutton(ff, text=FIT_LABELS[v], font=FONT, variable=fv, value=v,
                           command=self.refresh).pack(anchor="w")

        # ---- 底部按钮 ----
        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=16, pady=(0, 16))
        ttk.Button(bottom, text="应用到所有图片", command=self.apply_all, width=16).pack(side="left")
        ttk.Button(bottom, text="取消", command=self.destroy, width=10).pack(side="right")
        ttk.Button(bottom, text="确定", command=self.confirm, width=10).pack(side="right", padx=(0, 8))

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.refresh()
        self._place_near(master)
        self.grab_set()

    def _place_near(self, master):
        self.update_idletasks()
        x = master.winfo_rootx() + 80
        y = master.winfo_rooty() + 60
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

    def current_opts(self):
        return {"rotate": self.rv.get(), "page": self.pv.get(), "fit": self.fv.get()}

    def refresh(self):
        """设置一变就重画预览。"""
        if self._after_id:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = self.after(120, self._do_preview)

    def _do_preview(self):
        self._after_id = None
        path = self.paths[0]
        opts = self.current_opts()
        try:
            from PIL import ImageTk
            canvas, _ = render_image_page(path, opts)
            box = 300
            canvas.thumbnail((box, box))
            self._photo = ImageTk.PhotoImage(canvas)
            self.preview.config(image=self._photo, width=canvas.width,
                                height=canvas.height, text="")
            w, h = canvas.size
            self.preview_note.config(
                text=f"{os.path.basename(path)}\n{IMAGE_SIZE_CACHE.get(path, ('?', '?'))[0]}"
                     f"×{IMAGE_SIZE_CACHE.get(path, ('?', '?'))[1]} 像素\n"
                     f"预览缩略到 {w}×{h}"
            )
        except Exception as exc:
            self._photo = None
            self.preview.config(image="", width=34, height=22,
                                text="预览失败")
            self.preview_note.config(text=f"预览画不出来：{exc}")

    def _apply_to(self, paths):
        opts = self.current_opts()
        for p in paths:
            self.app.img_opts[p] = dict(opts)
        self.app.refresh_rows()

    def confirm(self):
        self._apply_to(self.paths)
        self.destroy()

    def apply_all(self):
        all_imgs = [p for p in self.app.paths() if is_image(p)]
        self._apply_to(all_imgs)
        messagebox.showinfo("已应用", f"设置已套用到全部 {len(all_imgs)} 张图片。")
        self.destroy()


IMAGE_SIZE_CACHE = {}


def render_pdf_page(path, page_index, target_w):
    """
    把 PDF 的某一页画成图片。target_w = 画出来多宽（像素）。
    画不出来返回 None（加密、损坏之类的）。
    """
    try:
        import pypdfium2 as pdfium
        with pdfium.PdfDocument(path) as doc:
            if page_index >= len(doc):
                return None
            page = doc[page_index]
            w = page.get_width() or 595.0
            pil = page.render(scale=target_w / w).to_pil()
            page.close()
            return pil
    except Exception:
        return None


# ==================== 放大看某一页 ====================
class ZoomDialog(tk.Toplevel):
    """把某一页放大看清楚，可以左右翻页。"""

    def __init__(self, master, app, pages, index):
        super().__init__(master)
        self.app = app
        self.pages = pages
        self.index = index
        self._photo = None
        self._token = 0

        self.title("查看")
        self.geometry("920x880")
        self.minsize(420, 320)

        top = tk.Frame(self)
        top.pack(fill="x", padx=12, pady=(10, 6))
        self.lbl = tk.Label(top, text="", font=FONT_BOLD, anchor="w")
        self.lbl.pack(side="left")
        self.pagelbl = tk.Label(top, text="", font=FONT, fg=GREY)
        self.pagelbl.pack(side="right")

        mid = tk.Frame(self)
        mid.pack(fill="both", expand=True, padx=12)
        self.canvas = tk.Canvas(mid, bg="#e8e8e8", highlightthickness=0)
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=hsb.set)
        hsb.pack(fill="x", padx=12)

        bot = tk.Frame(self)
        bot.pack(fill="x", padx=12, pady=10)
        self.btn_prev = ttk.Button(bot, text="← 上一页", width=12, command=lambda: self.go(-1))
        self.btn_prev.pack(side="left")
        self.btn_next = ttk.Button(bot, text="下一页 →", width=12, command=lambda: self.go(1))
        self.btn_next.pack(side="left", padx=(8, 0))
        ttk.Button(bot, text="关闭", width=10, command=self.destroy).pack(side="right")

        for seq in ("<MouseWheel>",):
            self.canvas.bind(seq, self._on_wheel)
        self.canvas.bind("<Prior>", lambda e: self.go(-1))
        self.canvas.bind("<Next>", lambda e: self.go(1))

        self.bind("<Left>", lambda e: self.go(-1))
        self.bind("<Right>", lambda e: self.go(1))
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Control-Home>", lambda e: self.go(-len(self.pages)))
        self.bind("<Control-End>", lambda e: self.go(len(self.pages)))
        self.focus_set()
        self.show()

    def _on_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def go(self, delta):
        new = self.index + delta
        if 0 <= new < len(self.pages):
            self.index = new
            self.show()

    def show(self):
        info = self.pages[self.index]
        self.lbl.config(text=f"P{info['no']} · {info['name']}")
        self.pagelbl.config(text=f"第 {self.index + 1} / {len(self.pages)} 页")
        self.title(f"第 {info['no']} 页 — {info['name']}")
        self.btn_prev.config(state=("normal" if self.index > 0 else "disabled"))
        self.btn_next.config(state=("normal" if self.index < len(self.pages) - 1 else "disabled"))

        self._token += 1
        token = self._token
        self._photo = None
        self.canvas.delete("all")
        self.canvas.create_text(24, 24, anchor="nw", text="正在画这一页…",
                                font=FONT, fill=GREY)
        self.canvas.configure(scrollregion=(0, 0, 200, 120))
        self.after(10, lambda: self._render(token))

    def _render(self, token):
        if token != self._token or not self.winfo_exists():
            return
        from PIL import Image, ImageTk
        info = self.pages[self.index]
        pil = None
        try:
            if info["kind"] == "pdf":
                pil = render_pdf_page(info["path"], info["page_index"], ZOOM_W)
            else:
                opts = self.app.img_opts.get(info["path"], default_img_opts())
                pil = render_image_page(info["path"], opts)[0]
        except Exception:
            pil = None

        if token != self._token or not self.winfo_exists():
            return
        self.canvas.delete("all")
        if pil is None:
            self.canvas.create_text(24, 24, anchor="nw",
                                    text="这一页画不出来（文件加密或者已损坏）",
                                    font=FONT, fill=RED)
            self.canvas.configure(scrollregion=(0, 0, 400, 100))
            return

        if pil.width > ZOOM_W:
            ratio = ZOOM_W / pil.width
            pil = pil.resize((ZOOM_W, max(1, int(pil.height * ratio))), Image.LANCZOS)
        self._photo = ImageTk.PhotoImage(pil)
        self.canvas.create_image(0, 0, image=self._photo, anchor="nw")
        self.canvas.configure(scrollregion=(0, 0, pil.width, pil.height))
        self.canvas.yview_moveto(0)
        self.canvas.xview_moveto(0)


# ==================== 预览窗口 ====================
class PreviewDialog(tk.Toplevel):
    """按当前顺序，把合并后的每一页都画成缩略图。"""

    def __init__(self, master, app, start_no=1):
        super().__init__(master)
        self.app = app
        self.closed = False
        self.pages = []      # 每页的信息
        self.thumbs = []     # 缩略图引用，丢了图就不显示
        self.start_no = start_no
        self.error = None

        self.title("预览 — 合并后的样子")
        self.geometry("980x800")
        self.minsize(560, 420)
        self.transient(master)

        head = tk.Frame(self)
        head.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(head, text="合并后的样子", font=FONT_BIG).pack(anchor="w")
        self.info = tk.Label(head, text="正在生成缩略图…", font=FONT, fg=GREY)
        self.info.pack(anchor="w", pady=(3, 0))

        wrap = tk.Frame(self)
        wrap.pack(fill="both", expand=True, padx=16, pady=(8, 0))
        self.canvas = tk.Canvas(wrap, bg="#f2f2f2", highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        pad = (THUMB_W + 10) * THUMB_COLS + 16 * (THUMB_COLS + 1)
        self.inner = tk.Frame(self.canvas, bg="#f2f2f2", width=pad, height=10)
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        for w in (self.canvas, self.inner):
            w.bind("<MouseWheel>", self._on_wheel)

        bot = tk.Frame(self)
        bot.pack(fill="x", padx=16, pady=12)
        tk.Label(bot, text="点任意一张可以放大看，放大后能左右翻页。",
                 font=FONT, fg=GREY).pack(side="left")
        ttk.Button(bot, text="关闭", width=10, command=self.destroy).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda e: self.destroy())

        self.queue = queue.Queue()
        threading.Thread(target=self._worker, daemon=True).start()
        self.after(60, self._pump)

    def _on_wheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    # ---------- 后台线程：只画图，绝不碰界面 ----------
    def _worker(self):
        no = 0
        try:
            for path in list(self.app.entries):
                if self.closed:
                    return
                name = os.path.basename(path)
                if is_pdf(path):
                    import pypdfium2 as pdfium
                    with pdfium.PdfDocument(path) as doc:
                        for i in range(len(doc)):
                            if self.closed:
                                return
                            page = doc[i]
                            w = page.get_width() or 595.0
                            pil = page.render(scale=THUMB_W * 1.6 / w).to_pil()
                            page.close()
                            no += 1
                            self.queue.put(("page", (pil, {
                                "no": no, "kind": "pdf", "path": path,
                                "page_index": i, "name": name})))
                else:
                    opts = self.app.img_opts.get(path, default_img_opts())
                    pil = render_image_page(path, opts)[0]
                    no += 1
                    self.queue.put(("page", (pil, {
                        "no": no, "kind": "image", "path": path,
                        "page_index": 0, "name": name})))
            self.queue.put(("done", no))
        except Exception:
            self.queue.put(("error", traceback.format_exc()))

    # ---------- 主线程：定时把队列里的东西拿出来画到界面上 ----------
    def _pump(self):
        if self.closed:
            return
        for _ in range(64):
            try:
                kind, payload = self.queue.get_nowait()
            except queue.Empty:
                break
            if kind == "page":
                self._add(*payload)
            elif kind == "error":
                self._fail(payload)
                return
            elif kind == "done":
                self._done(payload)
                return
        try:
            self.after(50, self._pump)
        except Exception:
            pass

    def _add(self, pil, info):
        if self.closed:
            return
        from PIL import Image, ImageTk
        thumb = pil.copy()
        thumb.thumbnail((THUMB_W, THUMB_H), Image.LANCZOS)
        photo = ImageTk.PhotoImage(thumb)
        self.thumbs.append(photo)

        idx = len(self.pages)
        self.pages.append(info)
        r, c = divmod(idx, THUMB_COLS)

        # 每格做成一样大的卡片，横页竖页都能对齐
        cell = tk.Frame(self.inner, bg="#f2f2f2",
                        width=THUMB_W + 10, height=THUMB_H + 28)
        cell.grid(row=r, column=c, padx=8, pady=10)
        cell.grid_propagate(False)
        cell.columnconfigure(0, weight=1)
        cell.rowconfigure(0, weight=1)

        lab = tk.Label(cell, image=photo, bg="white", relief="solid", bd=1, cursor="hand2")
        lab.grid(row=0, column=0)
        cap = tk.Label(cell, text=f"P{info['no']} · {info['name']}", font=FONT_SMALL,
                       fg="#555555", bg="#f2f2f2", wraplength=THUMB_W + 6, justify="center")
        cap.grid(row=1, column=0, sticky="ew")

        for w in (lab, cap):
            w.bind("<Button-1>", lambda e, i=idx: self._zoom(i))
            w.bind("<MouseWheel>", self._on_wheel)

        self.info.config(text=f"正在生成缩略图… 已经画了 {len(self.pages)} 页")

    def _zoom(self, index):
        ZoomDialog(self, self.app, self.pages, index)

    def _done(self, n):
        if self.closed:
            return
        self.info.config(text=f"共 {n} 页。点任意一张可以放大看。", fg=GREEN)
        if self.start_no > 1:
            self._scroll_to(self.start_no)

    def _fail(self, msg):
        if self.closed:
            return
        self.error = msg
        last = [ln for ln in msg.strip().splitlines() if ln.strip()]
        brief = last[-1] if last else "未知错误"
        self.info.config(text=f"生成缩略图出错了：{brief}", fg=RED)

    def _scroll_to(self, page_no):
        """滚到第 N 页那儿，省得自己找。"""
        idx = page_no - 1
        kids = self.inner.winfo_children()
        if idx < 0 or idx >= len(kids):
            return
        self.inner.update_idletasks()
        y = kids[idx].winfo_y()
        total = max(1, self.inner.winfo_height())
        self.canvas.yview_moveto(max(0.0, (y - 12) / total))

    def destroy(self):
        self.closed = True
        self.thumbs.clear()
        super().destroy()


# ==================== 主窗口 ====================
class MergeApp:
    def __init__(self, root, out_dir=None, out_name=None):
        self.root = root
        # 从「简历批量投递助手」里打开时，会把附件库目录传进来，
        # 这样保存对话框直接落在附件库，合完就能勾选发送。
        self.out_dir = out_dir if (out_dir and os.path.isdir(out_dir)) else None
        self.out_name = out_name or ""
        self.entries = []      # 路径列表。顺序 = 合并顺序
        self.meta = {}         # 路径 -> dict(kind, c1, c2, c3)
        self.img_opts = {}     # 图片路径 -> dict(rotate, page, fit)
        self.items = {}        # 路径 -> Treeview 行 id
        self.busy = False
        self.merge_queue = queue.Queue()

        root.title(APP_TITLE + ("　—　保存到投递助手的附件库" if self.out_dir else ""))
        root.minsize(640, 460)
        self._center(780, 580)

        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Treeview", font=FONT, rowheight=28)
        style.configure("Treeview.Heading", font=FONT_BOLD)
        style.configure("Big.TButton", font=FONT_BOLD, padding=(14, 8))

        # ---------- 顶部 ----------
        head = tk.Frame(root)
        head.pack(fill="x", padx=16, pady=(16, 8))
        tk.Label(head, text="要合并的内容", font=FONT_BIG).pack(anchor="w")
        tk.Label(
            head,
            text="PDF 和图片都能加。下面的顺序＝合并后 PDF 的页面顺序，从上往下。"
                 "选中图片后点「图片设置…」可以旋转、改页面大小。",
            font=FONT, fg=GREY, justify="left",
        ).pack(anchor="w", pady=(3, 0))

        # ---------- 列表 ----------
        box = tk.Frame(root)
        box.pack(fill="both", expand=True, padx=16)

        self.tree = ttk.Treeview(
            box, columns=("c1", "c2", "c3"), show="tree headings", selectmode="extended"
        )
        self.tree.heading("#0", text="文件名", anchor="w")
        self.tree.heading("c1", text="页数 / 尺寸", anchor="center")
        self.tree.heading("c2", text="大小", anchor="center")
        self.tree.heading("c3", text="类型", anchor="center")
        self.tree.column("#0", width=420, minwidth=200, stretch=True, anchor="w")
        self.tree.column("c1", width=120, minwidth=100, stretch=False, anchor="center")
        self.tree.column("c2", width=90, minwidth=80, stretch=False, anchor="center")
        self.tree.column("c3", width=110, minwidth=90, stretch=False, anchor="center")

        vsb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<Double-1>", self._on_double_click)

        # ---------- 按钮 ----------
        bar = tk.Frame(root)
        bar.pack(fill="x", padx=16, pady=(12, 0))
        ttk.Button(bar, text="添加文件…", command=self.add_files, width=13).pack(side="left")
        ttk.Button(bar, text="上移", command=lambda: self.move(-1), width=7).pack(side="left", padx=(10, 0))
        ttk.Button(bar, text="下移", command=lambda: self.move(1), width=7).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="移除选中", command=self.remove_selected, width=10).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="清空", command=self.clear_all, width=7).pack(side="left", padx=(8, 0))
        self.btn_img = ttk.Button(bar, text="图片设置…", command=self.open_image_settings, width=12)
        self.btn_img.pack(side="left", padx=(20, 0))

        # ---------- 底部 ----------
        bottom = tk.Frame(root)
        bottom.pack(fill="x", padx=16, pady=16)
        self.btn_merge = ttk.Button(bottom, text="合并并保存…", style="Big.TButton",
                                    command=self.do_merge)
        self.btn_merge.pack(side="left")
        self.btn_preview = ttk.Button(bottom, text="预览…", style="Big.TButton",
                                      command=self.open_preview)
        self.btn_preview.pack(side="left", padx=(10, 0))
        self.status = tk.Label(bottom, text="还没添加文件", font=FONT, fg=GREY, anchor="w")
        self.status.pack(side="left", padx=(14, 0), fill="x", expand=True)

        root.bind("<Control-o>", lambda e: self.add_files())

    # ================= 通用 =================
    def _center(self, w, h):
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")

    def set_status(self, text, color=GREY):
        self.status.config(text=text, fg=color)

    def paths(self):
        return list(self.entries)

    def _default_dir(self):
        return os.path.dirname(self.entries[0]) if self.entries else os.path.expanduser("~")

    def _img_count(self):
        return sum(1 for p in self.entries if is_image(p))

    def _total_pages(self):
        """估算合并后总页数：PDF 按实际页数，图片一张算一页。"""
        n = 0
        for p in self.entries:
            m = self.meta.get(p)
            if not m:
                continue
            if m["kind"] == "pdf":
                n += m.get("pages") or 0
            else:
                n += 1
        return n

    # ================= 列表 =================
    def _make_meta(self, path):
        """算出这个文件在列表里怎么显示。搞不定返回 None。"""
        if is_pdf(path):
            pages = read_pdf_pages(path)
            if pages is None:
                return None
            return {"kind": "pdf", "pages": pages, "c1": f"{pages} 页",
                    "c2": human_size(os.path.getsize(path)), "c3": "PDF"}
        if is_image(path):
            size = read_image_size(path)
            if size is None:
                return None
            IMAGE_SIZE_CACHE[path] = size
            o = self.img_opts.setdefault(path, default_img_opts())
            return {"kind": "image", "size": size,
                    "c1": f"{size[0]}×{size[1]}",
                    "c2": human_size(os.path.getsize(path)),
                    "c3": self._img_type_text(o)}
        return None

    def _img_type_text(self, opts):
        parts = ["图片"]
        if opts.get("rotate"):
            parts.append(f"{opts['rotate']}°")
        if opts.get("page") == "a4p":
            parts.append("A4纵")
        elif opts.get("page") == "a4l":
            parts.append("A4横")
        if opts.get("fit") == "cover":
            parts.append("铺满")
        return " ".join(parts)

    def _add_one(self, path):
        path = os.path.abspath(path)
        if path in self.items:
            return False
        if not (is_pdf(path) or is_image(path)):
            return False
        meta = self._make_meta(path)
        if meta is None:
            return False
        self.entries.append(path)
        self.meta[path] = meta
        self._rebuild()
        return True

    def add_files(self):
        picked = filedialog.askopenfilenames(
            title="选要合并的 PDF 或图片（按住 Ctrl 可一次多选）",
            filetypes=FILE_TYPES,
            initialdir=self._default_dir(),
        )
        if not picked:
            return
        skipped = []
        for p in picked:
            if not self._add_one(p):
                skipped.append(os.path.basename(p))
        self._refresh_status()
        if skipped:
            messagebox.showwarning(
                "有文件没加进来",
                "下面这些跳过了（格式不支持、重复添加，或者打不开）：\n\n" + "\n".join(skipped),
            )

    def _rebuild(self):
        selected = self._selected_paths()
        self.tree.delete(*self.tree.get_children())
        self.items.clear()
        for i, p in enumerate(self.entries, 1):
            m = self.meta.get(p, {})
            item = self.tree.insert("", "end", text=f"{i}. {os.path.basename(p)}",
                                    values=(m.get("c1", "?"), m.get("c2", "?"), m.get("c3", "?")))
            self.items[p] = item
        keep = [self.items[p] for p in selected if p in self.items]
        if keep:
            self.tree.selection_set(keep)

    def refresh_rows(self):
        """图片设置改了之后，刷新列表里对应的文字。"""
        for p, m in self.meta.items():
            if m.get("kind") == "image":
                o = self.img_opts.get(p, default_img_opts())
                m["c3"] = self._img_type_text(o)
        self._rebuild()
        self._refresh_status()

    def _selected_paths(self):
        wanted = set(self.tree.selection())
        return [p for p in self.entries if self.items.get(p) in wanted]

    def _refresh_status(self):
        if not self.entries:
            self.set_status("还没添加文件")
            return
        n_img = self._img_count()
        n_pdf = len(self.entries) - n_img
        txt = f"已选 {len(self.entries)} 个：PDF {n_pdf} 个、图片 {n_img} 张，合计约 {self._total_pages()} 页"
        self.set_status(txt)

    def move(self, delta):
        sel = set(self._selected_paths())
        if not sel:
            self.set_status("先点一行选中它，再点上移/下移", RED)
            return
        order = range(len(self.entries) - 1, -1, -1) if delta > 0 else range(len(self.entries))
        moved = False
        for idx in order:
            if self.entries[idx] not in sel:
                continue
            new = idx + delta
            if new < 0 or new >= len(self.entries) or self.entries[new] in sel:
                continue
            self.entries[idx], self.entries[new] = self.entries[new], self.entries[idx]
            moved = True
        if moved:
            self._rebuild()
        self._refresh_status()

    def remove_selected(self):
        sel = self._selected_paths()
        if not sel:
            return
        for p in sel:
            if p in self.entries:
                self.entries.remove(p)
            self.meta.pop(p, None)
            self.img_opts.pop(p, None)
            self.items.pop(p, None)
        self._rebuild()
        self._refresh_status()

    def clear_all(self):
        if not self.entries:
            return
        self.entries.clear()
        self.meta.clear()
        self.img_opts.clear()
        self.items.clear()
        self.tree.delete(*self.tree.get_children())
        self._refresh_status()

    def _on_double_click(self, event):
        """双击图片行 = 打开图片设置；双击 PDF 行 = 预览，直接跳到它那一页。"""
        row = self.tree.identify_row(event.y)
        if not row:
            return
        for p, i in self.items.items():
            if i == row:
                if is_image(p):
                    self.open_image_settings()
                else:
                    self.open_preview(start_no=self._page_no_of(p))
                break

    # ================= 预览 =================
    def _page_no_of(self, path):
        """这个文件的第一页，排在合并结果的第几页（从 1 开始数）。"""
        n = 0
        for p in self.entries:
            if p == path:
                break
            m = self.meta.get(p, {})
            n += (m.get("pages") or 1) if m.get("kind") == "pdf" else 1
        return n + 1

    def open_preview(self, start_no=1):
        if not self.entries:
            messagebox.showinfo("提示", "还没添加任何文件。\n\n请点「添加文件…」。")
            return
        PreviewDialog(self.root, self, start_no=start_no)

    # ================= 图片设置 =================
    def open_image_settings(self):
        sel = [p for p in self._selected_paths() if is_image(p)]
        if not sel:
            messagebox.showinfo("提示", "先在列表里点一行选中图片，再点这个按钮。")
            return
        ImageSettingsDialog(self.root, self, sel)

    # ================= 合并 =================
    def do_merge(self):
        if self.busy:
            return
        if not self.entries:
            messagebox.showinfo("提示", "还没添加任何文件。\n\n请点「添加文件…」。")
            return
        if len(self.entries) == 1 and is_pdf(self.entries[0]):
            if not messagebox.askyesno("提示", "只有 1 个 PDF，合并出来和原文件一样。要继续吗？"):
                return

        first = self.entries[0]
        if self.out_name:
            default_name = os.path.splitext(os.path.basename(self.out_name))[0]
        else:
            default_name = os.path.splitext(os.path.basename(first))[0] + "-合并"
        out = filedialog.asksaveasfilename(
            title="合并后的 PDF 存到哪",
            initialdir=self.out_dir or os.path.dirname(first),
            initialfile=default_name + ".pdf",
            defaultextension=".pdf",
            filetypes=[("PDF 文件", "*.pdf")],
        )
        if not out:
            return
        out = os.path.abspath(out)

        for p in self.entries:
            if os.path.normcase(p) == os.path.normcase(out):
                if not messagebox.askyesno(
                    "小心覆盖",
                    f"保存的位置和这个源文件重名了：\n\n{os.path.basename(p)}\n\n"
                    "继续的话它会被覆盖。真的要继续吗？",
                ):
                    return
                break

        self.busy = True
        self.btn_merge.config(state="disabled")
        self.btn_img.config(state="disabled")
        self.btn_preview.config(state="disabled")
        self.set_status("正在合并，图片多的话会慢一点，稍等…", BLUE)
        self.merge_queue = queue.Queue()
        threading.Thread(target=self._merge_worker, args=(out,), daemon=True).start()
        self.root.after(80, self._pump_merge)

    def _pump_merge(self):
        """主线程里盯着后台合并的结果。"""
        try:
            out, pages, err = self.merge_queue.get_nowait()
        except queue.Empty:
            self.root.after(80, self._pump_merge)
            return
        self._merge_done(out, pages, err)

    def _merge_worker(self, out):
        """后台干活：PDF 直接搬页，图片先画成 A4 再搬。"""
        tmpdir = None
        try:
            from pypdf import PdfReader, PdfWriter
            from PIL import Image

            writer = PdfWriter()
            pages = 0
            tmpdir = tempfile.mkdtemp(prefix="pdfmerge_")

            for idx, p in enumerate(list(self.entries)):
                if is_pdf(p):
                    reader = PdfReader(p)
                    if getattr(reader, "is_encrypted", False):
                        try:
                            reader.decrypt("")
                        except Exception:
                            pass
                    for page in reader.pages:
                        writer.add_page(page)
                    pages += len(reader.pages)
                    continue

                # 图片：画到白纸上 → 存成临时单页 PDF → 再搬页
                opts = self.img_opts.get(p, default_img_opts())
                canvas, dpi = render_image_page(p, opts)
                tmp_pdf = os.path.join(tmpdir, f"img_{idx}.pdf")
                canvas.save(tmp_pdf, "PDF", resolution=dpi)
                canvas.close()
                reader = PdfReader(tmp_pdf)
                for page in reader.pages:
                    writer.add_page(page)
                pages += len(reader.pages)

            if pages == 0:
                raise RuntimeError("没拼出任何页面，请检查添加的文件。")

            tmp_out = out + ".part"
            with open(tmp_out, "wb") as f:
                writer.write(f)
            os.replace(tmp_out, out)
            self.merge_queue.put((out, pages, None))
        except Exception as exc:
            self.merge_queue.put((out, 0, f"{exc}\n\n{traceback.format_exc()}"))
        finally:
            if tmpdir:
                shutil.rmtree(tmpdir, ignore_errors=True)

    def _merge_done(self, out, pages, err):
        self.busy = False
        self.btn_merge.config(state="normal")
        self.btn_img.config(state="normal")
        self.btn_preview.config(state="normal")
        if err:
            self.set_status("合并失败", RED)
            messagebox.showerror("合并失败", "出错了：\n\n" + err)
            return
        self.set_status(f"完成：{pages} 页 → {os.path.basename(out)}", GREEN)
        if messagebox.askyesno(
            "合并完成",
            f"搞定，一共 {pages} 页。\n\n文件在：\n{out}\n\n要现在打开看看吗？",
        ):
            try:
                os.startfile(out)
            except Exception:
                pass


def main():
    # 命令行参数：--outdir / --outname 由「简历批量投递助手」传入；
    # 其余参数当作要预加载的文件（支持把文件拖到 bat 图标上打开）。
    out_dir = None
    out_name = None
    files = []
    args = list(sys.argv[1:])
    i = 0
    while i < len(args):
        if args[i] == "--outdir" and i + 1 < len(args):
            out_dir = args[i + 1]
            i += 2
            continue
        if args[i] == "--outname" and i + 1 < len(args):
            out_name = args[i + 1]
            i += 2
            continue
        files.append(args[i])
        i += 1

    root = tk.Tk()
    app = MergeApp(root, out_dir=out_dir, out_name=out_name)

    for p in files:
        if os.path.isfile(p):
            app._add_one(p)
    app._refresh_status()

    def hook(exc_type, exc_value, exc_tb):
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        try:
            messagebox.showerror("程序出错了", msg)
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc_value, exc_tb)

    sys.excepthook = hook
    root.mainloop()


if __name__ == "__main__":
    main()
