# -*- coding: utf-8 -*-
"""
简历批量投递助手 - 本地服务
功能：投递清单管理（增删改查 / Excel 导入导出）、邮件模板（变量替换）、
     证明材料合并成一个 PDF（PDF + 图片混拼，图片自动摆正 / 压到 A4）、
     QQ 邮箱 SMTP 批量发送（间隔控制 / 每日上限 / 失败记录 / 中途停止）
说明：所有数据只保存在本机 data/ 与 attachments/ 目录中，不经过任何第三方服务器。
"""
import base64
import io
import json
import os
import shutil
import smtplib
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from datetime import datetime
from email.header import Header
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, urlparse

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "data")
ATTACH_DIR = os.path.join(BASE, "attachments")
INDEX_HTML = os.path.join(BASE, "index.html")

ATT_NOTE_FILE = "attachment_notes.json"   # {文件名: 备注文字}，备注只存在本机，不会随邮件发出

# ---------------------------------------------------------------- 材料合并
# 暂存区：用户丢进来的证明材料先放这儿，排好顺序再合成。
# 顺序直接写进文件名开头的序号（001__、002__…），所以关掉软件再打开顺序也不会乱。
MERGE_DIR = os.path.join(DATA_DIR, "_merge_staging")

A4_W, A4_H = 595.28, 841.89   # A4 纸尺寸，单位「点」（1 点 = 1/72 英寸）
MERGE_DPI = 144               # 图片放进 PDF 的分辨率。144 够 HR 看清，文件也不会太大
MERGE_MARGIN = 0.04           # 页面四周留 4% 白边，证书照片贴着纸边不好看、打印也容易被切

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".gif", ".tif", ".tiff"}

SMTP_HOST = "smtp.qq.com"
SMTP_PORT = 465

DEFAULT_CONFIG = {
    "email": "",          # QQ 邮箱地址
    "auth_code": "",      # QQ 邮箱授权码（不是登录密码）
    "display_name": "",   # 发件人显示名
    "name": "",           # 我的姓名（模板变量 {姓名}）
    "school": "",         # 我的学校（模板变量 {学校}）
    "major": "",          # 我的专业（模板变量 {专业}）
    "interval": 40,       # 每封邮件间隔（秒）
    "daily_limit": 50,    # 每日发送上限
}

DEFAULT_TEMPLATE = {
    "subject": "应聘【{岗位}】- {姓名}（{学校}）",
    "body": (
        "尊敬的HR：\n\n"
        "您好！\n\n"
        "我是{学校}{姓名}，从招聘信息中了解到贵司正在招聘{岗位}，"
        "我对该岗位非常感兴趣，现将我的简历附上，期待能有机会参加面试。\n\n"
        "感谢您在百忙之中查阅我的邮件，期待您的回复！\n\n"
        "祝工作顺利！\n\n"
        "{姓名}\n"
        "{学校}\n"
    ),
}


# ---------------------------------------------------------------- 数据存取
def _path(name):
    return os.path.join(DATA_DIR, name)


def load_json(name, default):
    try:
        with open(_path(name), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(name, obj):
    tmp = _path(name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, _path(name))


def init_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(ATTACH_DIR, exist_ok=True)
    os.makedirs(MERGE_DIR, exist_ok=True)
    if not os.path.exists(_path("config.json")):
        save_json("config.json", DEFAULT_CONFIG)
    if not os.path.exists(_path("jobs.json")):
        save_json("jobs.json", [])
    if not os.path.exists(_path("template.json")):
        save_json("template.json", DEFAULT_TEMPLATE)
    if not os.path.exists(_path("sent_log.json")):
        save_json("sent_log.json", [])
    if not os.path.exists(_path(ATT_NOTE_FILE)):
        save_json(ATT_NOTE_FILE, {})


def get_config():
    cfg = dict(DEFAULT_CONFIG)
    cfg.update(load_json("config.json", {}))
    return cfg


def get_template():
    tpl = dict(DEFAULT_TEMPLATE)
    tpl.update(load_json("template.json", {}))
    return tpl


def get_jobs():
    return load_json("jobs.json", [])


def save_jobs(jobs):
    save_json("jobs.json", jobs)


def today_sent_count():
    today = datetime.now().strftime("%Y-%m-%d")
    logs = load_json("sent_log.json", [])
    return sum(1 for e in logs if e.get("sent_at", "").startswith(today) and e.get("ok"))


def append_sent_log(entry):
    logs = load_json("sent_log.json", [])
    logs.append(entry)
    save_json("sent_log.json", logs)


# ---------------------------------------------------------------- 变量渲染
def render(tpl_text, job, cfg):
    mapping = {
        "{公司}": job.get("company", ""),
        "{岗位}": job.get("position", ""),
        "{备注}": job.get("note", ""),
        "{姓名}": cfg.get("name", ""),
        "{学校}": cfg.get("school", ""),
        "{专业}": cfg.get("major", ""),
    }
    out = tpl_text or ""
    for k, v in mapping.items():
        out = out.replace(k, v)
    return out


# ---------------------------------------------------------------- 邮件发送
def get_att_notes():
    d = load_json(ATT_NOTE_FILE, {})
    return d if isinstance(d, dict) else {}


def set_att_note(name, note):
    """写入/清除某个附件的备注（备注为空等于删掉备注）。备注只存本机，不会写进邮件。"""
    d = get_att_notes()
    note = (note or "").strip()
    if note:
        d[name] = note
    else:
        d.pop(name, None)
    save_json(ATT_NOTE_FILE, d)
    return d


def drop_att_note(name):
    d = get_att_notes()
    if name in d:
        d.pop(name)
        save_json(ATT_NOTE_FILE, d)


def list_attachments():
    notes = get_att_notes()
    items = []
    for fn in sorted(os.listdir(ATTACH_DIR)):
        p = os.path.join(ATTACH_DIR, fn)
        if os.path.isfile(p):
            items.append({"name": fn, "size": os.path.getsize(p), "note": notes.get(fn, "")})
    return items


# ---------------------------------------------------------------- 材料合并
def safe_fname(name):
    """只保留文件名本身（去掉路径），再把 Windows 不认的字符换成下划线。"""
    name = os.path.basename(str(name or "")).strip()
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, "_")
    return name or "未命名"


def is_pdf_name(name):
    return os.path.splitext(name)[1].lower() == ".pdf"


def is_image_name(name):
    return os.path.splitext(name)[1].lower() in IMAGE_EXTS


def pdf_page_count(path):
    """读 PDF 页数；读不了（加密/损坏）返回 None。"""
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


def _staged_files():
    """暂存区里所有文件，按文件名开头的序号排好。"""
    if not os.path.isdir(MERGE_DIR):
        return []
    out = []
    for fn in sorted(os.listdir(MERGE_DIR)):
        p = os.path.join(MERGE_DIR, fn)
        if os.path.isfile(p) and "__" in fn:
            out.append((fn, fn.split("__", 1)[1]))
    return out


def merge_list():
    items = []
    for stored, shown in _staged_files():
        p = os.path.join(MERGE_DIR, stored)
        kind = "pdf" if is_pdf_name(shown) else ("image" if is_image_name(shown) else "other")
        items.append({
            "id": stored,
            "name": shown,
            "size": os.path.getsize(p),
            "kind": kind,
            "ok": kind in ("pdf", "image"),
            "pages": pdf_page_count(p) if kind == "pdf" else (1 if kind == "image" else 0),
        })
    return items


def merge_add(name, data):
    """把上传的文件放进暂存区，序号续在最后。"""
    os.makedirs(MERGE_DIR, exist_ok=True)
    shown = safe_fname(name)
    used = [int(f.split("__", 1)[0]) for f, _ in _staged_files() if f.split("__", 1)[0].isdigit()]
    seq = (max(used) + 1) if used else 1
    while True:
        target = os.path.join(MERGE_DIR, "%03d__%s" % (seq, shown))
        if not os.path.exists(target):
            break
        seq += 1
    with open(target, "wb") as f:
        f.write(data)
    return target


def merge_move(fid, direction):
    """把某个文件往上/往下挪一格：交换两个文件的序号。"""
    ids = [stored for stored, _ in _staged_files()]
    if fid not in ids:
        return
    i = ids.index(fid)
    j = i + (-1 if direction < 0 else 1)
    if j < 0 or j >= len(ids):
        return
    a, b = ids[i], ids[j]
    seq_a, name_a = a.split("__", 1)
    seq_b, name_b = b.split("__", 1)
    tmp = os.path.join(MERGE_DIR, "000__swap.tmp")
    os.replace(os.path.join(MERGE_DIR, a), tmp)
    os.replace(os.path.join(MERGE_DIR, b), os.path.join(MERGE_DIR, seq_a + "__" + name_b))
    os.replace(tmp, os.path.join(MERGE_DIR, seq_b + "__" + name_a))


def merge_remove(fid):
    p = os.path.join(MERGE_DIR, os.path.basename(fid))
    if os.path.isfile(p):
        os.remove(p)
    # 删掉一个以后，把剩下的序号重新排成 001、002… 保持连续
    rest = _staged_files()
    for new_i, (stored, shown) in enumerate(rest, 1):
        want = "%03d__%s" % (new_i, shown)
        if want != stored:
            os.replace(os.path.join(MERGE_DIR, stored), os.path.join(MERGE_DIR, want))


def merge_clear():
    for stored, _ in _staged_files():
        try:
            os.remove(os.path.join(MERGE_DIR, stored))
        except OSError:
            pass


def render_image_page(path):
    """把一张图片画到一页 A4 白纸上：按比例缩到放得下，居中，四周留白。

    手机拍的照片常带"躺倒"的方向信息，先按它摆正。
    图是横的就用 A4 横向，是竖的就用 A4 纵向，免得横版证书被压得很小。
    返回 (PIL 图片, 输出 DPI)。
    """
    from PIL import Image, ImageOps

    im = Image.open(path)
    im.load()
    fixed = ImageOps.exif_transpose(im)
    if fixed is not None:
        im = fixed

    # 透明背景（PNG）先垫成白的，不然 PDF 里会变黑块
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")

    pw, ph = (A4_H, A4_W) if im.width > im.height else (A4_W, A4_H)
    k = MERGE_DPI / 72.0
    cw = max(1, int(round(pw * k)))
    ch = max(1, int(round(ph * k)))
    canvas = Image.new("RGB", (cw, ch), "white")

    mx = int(cw * MERGE_MARGIN)
    my = int(ch * MERGE_MARGIN)
    box_w = max(1, cw - mx * 2)
    box_h = max(1, ch - my * 2)
    ratio = min(box_w / im.width, box_h / im.height)
    nw = max(1, int(round(im.width * ratio)))
    nh = max(1, int(round(im.height * ratio)))
    if (nw, nh) != im.size:
        im = im.resize((nw, nh), Image.LANCZOS)
    canvas.paste(im, ((cw - nw) // 2, (ch - nh) // 2))

    # 用页面宽度反推 DPI，让 PDF 里的页面尺寸正好等于 A4（避免取整误差）
    return canvas, cw * 72.0 / pw


_GUI_PY = {"exe": None}


def find_gui_python():
    """找一个同时装了 tkinter 和 pypdf 的 Python，用来跑独立合并工具。

    坑：本软件自己跑在精简环境里，那个 Python 只有 pypdf、没有 tkinter，
    所以不能直接用 sys.executable（会报 No module named 'tkinter'，窗口根本开不出来）。
    这里按顺序挨个试，逻辑和「合并PDF.bat」保持一致，探测结果缓存起来。
    """
    if _GUI_PY["exe"]:
        return _GUI_PY["exe"]

    here = sys.executable or ""
    cands = []
    for c in (
        os.path.join(os.path.dirname(here), "pythonw.exe") if here else "",
        here,
        "pythonw",
        "python",
    ):
        if c and c not in cands:
            cands.append(c)

    flags = 0x08000000 if os.name == "nt" else 0
    for exe in cands:
        try:
            r = subprocess.run([exe, "-c", "import tkinter, pypdf"],
                               capture_output=True, timeout=25, creationflags=flags)
        except Exception:
            continue
        if r.returncode == 0:
            _GUI_PY["exe"] = exe
            return exe
    return None


def open_merge_tool(out_name=""):
    """把独立的「PDF / 图片 合并工具」作为单独窗口打开。

    它功能比网页版全（预览每一页、单张旋转、逐张调页面大小），所以主推这个。
    通过 --outdir 告诉它默认存到附件库，合完的 PDF 直接就能勾选发送。
    """
    script = os.path.join(BASE, "pdf_merge_gui.py")
    if not os.path.isfile(script):
        raise ValueError("找不到独立合并工具 pdf_merge_gui.py，请确认它还在软件目录里")

    exe = find_gui_python()
    if not exe:
        raise ValueError("没找到能开窗口的 Python（需要同时装了 tkinter 和 pypdf）。"
                         "可以双击软件目录里的「合并PDF.bat」手动打开。")

    cmd = [exe, script, "--outdir", ATTACH_DIR]
    if out_name:
        cmd += ["--outname", safe_fname(out_name)]
    kwargs = {"cwd": BASE}
    if os.name == "nt":
        # 不额外弹一个黑色控制台窗口（工具自己的窗口照常显示）
        kwargs["creationflags"] = 0x08000000
    subprocess.Popen(cmd, **kwargs)


def merge_run(out_name):
    """按暂存区顺序合成一个 PDF，直接放进附件库。返回 (文件名, 页数, 字节数)。"""
    from pypdf import PdfReader, PdfWriter

    picked = [it for it in merge_list() if it["ok"]]
    if not picked:
        raise ValueError("暂存区里还没有可合并的文件（支持 PDF 和图片）")

    writer = PdfWriter()
    pages = 0
    tmpdir = tempfile.mkdtemp(prefix="merge_")
    try:
        for it in picked:
            src = os.path.join(MERGE_DIR, it["id"])
            if it["kind"] == "pdf":
                reader = PdfReader(src)
                if getattr(reader, "is_encrypted", False):
                    try:
                        reader.decrypt("")
                    except Exception:
                        pass
                for page in reader.pages:
                    writer.add_page(page)
                pages += len(reader.pages)
                continue

            # 图片：先画到 A4 白纸上，存成临时单页 PDF，再把这页搬过来
            canvas, dpi = render_image_page(src)
            tmp_pdf = os.path.join(tmpdir, "img_%03d.pdf" % pages)
            canvas.save(tmp_pdf, "PDF", resolution=dpi)
            canvas.close()
            one = PdfReader(tmp_pdf)
            for page in one.pages:
                writer.add_page(page)
            pages += len(one.pages)

        if pages == 0:
            raise ValueError("没拼出任何页面，请检查添加的文件")

        out_name = safe_fname(out_name or "证明材料")
        if not out_name.lower().endswith(".pdf"):
            out_name += ".pdf"
        target = os.path.join(ATTACH_DIR, out_name)
        base, ext = os.path.splitext(out_name)
        i = 1
        while os.path.exists(target):
            target = os.path.join(ATTACH_DIR, "%s(%d)%s" % (base, i, ext))
            i += 1

        tmp_out = target + ".part"
        with open(tmp_out, "wb") as f:
            writer.write(f)
        os.replace(tmp_out, target)
        return os.path.basename(target), pages, os.path.getsize(target)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def resolve_attachments(job):
    """按投递目标解析附件：atts 为空 -> 发全部附件；否则只发勾选的（按勾选顺序）。"""
    chosen = job.get("atts") or []
    all_atts = list_attachments()
    if not chosen:
        return all_atts
    keep = set(chosen)
    picked = [a for a in all_atts if a["name"] in keep]
    order = {n: i for i, n in enumerate(chosen)}
    picked.sort(key=lambda a: order.get(a["name"], 999))
    return picked


def build_message(job, cfg, tpl):
    """构建邮件。优先使用该投递目标的单独定制（subject_override / body_override / atts），
    未定制的部分回落到全局模板。"""
    subject = render(job.get("subject_override") or tpl.get("subject", ""), job, cfg)
    body = render(job.get("body_override") or tpl.get("body", ""), job, cfg)
    display = cfg.get("display_name") or cfg.get("name") or "求职者"

    msg = MIMEMultipart()
    msg["From"] = formataddr((str(Header(display, "utf-8")), cfg["email"]))
    msg["To"] = job["email"]
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for att in resolve_attachments(job):
        with open(os.path.join(ATTACH_DIR, att["name"]), "rb") as f:
            data = f.read()
        part = MIMEApplication(data)
        part.add_header("Content-Disposition", "attachment",
                        filename=("utf-8", "", att["name"]))
        msg.attach(part)
    return msg, subject


def send_one(job, cfg, tpl):
    try:
        msg, subject = build_message(job, cfg, tpl)
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            s.login(cfg["email"], cfg["auth_code"])
            s.send_message(msg)
        return True, "", subject
    except smtplib.SMTPAuthenticationError:
        return False, "登录失败：邮箱地址或授权码不正确（授权码不是QQ密码，需在QQ邮箱设置里生成）", ""
    except (smtplib.SMTPException, socket.timeout, OSError) as e:
        return False, "发送失败：%s" % e, ""


# ---------------------------------------------------------------- 发送线程
SEND_STATE = {
    "running": False,
    "stop": False,
    "current": "",
    "done": 0,
    "total": 0,
    "log": [],  # {time, text, level}
}


def log_add(text, level="info"):
    SEND_STATE["log"].append({
        "time": datetime.now().strftime("%H:%M:%S"),
        "text": text,
        "level": level,
    })
    if len(SEND_STATE["log"]) > 200:
        SEND_STATE["log"] = SEND_STATE["log"][-200:]


def _interruptible_sleep(seconds):
    for _ in range(int(seconds * 2)):
        if SEND_STATE["stop"]:
            return
        time.sleep(0.5)


def sender_worker(job_ids):
    cfg = get_config()
    tpl = get_template()
    interval = max(5, int(cfg.get("interval", 40)))
    daily_limit = int(cfg.get("daily_limit", 50))
    jobs = get_jobs()
    by_id = {j["id"]: j for j in jobs}

    SEND_STATE["running"] = True
    SEND_STATE["stop"] = False
    SEND_STATE["done"] = 0
    SEND_STATE["total"] = len(job_ids)
    log_add("开始投递，共 %d 封，每封间隔 %d 秒" % (len(job_ids), interval))

    try:
        for idx, jid in enumerate(job_ids):
            if SEND_STATE["stop"]:
                log_add("已手动停止", "err")
                break
            job = by_id.get(jid)
            if not job:
                continue
            if today_sent_count() >= daily_limit:
                log_add("已达今日发送上限（%d 封），剩余目标已跳过，明天再继续" % daily_limit, "err")
                job["status"] = "待发"
                break

            job["status"] = "发送中"
            job["error"] = ""
            SEND_STATE["current"] = "%s · %s" % (job.get("company", ""), job.get("position", ""))
            save_jobs(jobs)

            ok, err, subject = send_one(job, cfg, tpl)
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if ok:
                job["status"] = "已发送"
                job["sent_at"] = now
                job["error"] = ""
                log_add("已发送：%s → %s（%s）" % (job.get("company"), job["email"], subject), "ok")
            else:
                job["status"] = "失败"
                job["error"] = err
                log_add("%s：%s" % (job.get("company"), err), "err")
            append_sent_log({
                "id": job["id"], "company": job.get("company"), "email": job["email"],
                "ok": ok, "error": err, "sent_at": now,
            })
            save_jobs(jobs)
            SEND_STATE["done"] += 1

            if idx < len(job_ids) - 1 and not SEND_STATE["stop"]:
                _interruptible_sleep(interval)
    finally:
        SEND_STATE["running"] = False
        SEND_STATE["current"] = ""
        log_add("本轮投递结束：成功处理 %d / %d" % (SEND_STATE["done"], SEND_STATE["total"]))


# ---------------------------------------------------------------- Excel
def excel_template_bytes():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "投递清单"
    headers = ["公司", "岗位", "收件邮箱", "备注"]
    fills = PatternFill("solid", fgColor="185FA5")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fills
    samples = [
        ["示例科技", "后端开发", "hr@example.com", "官网看到的，要求主题写姓名+学校"],
        ["示例工程局", "2027届校园招聘", "join@example2.com", "没写具体岗位就填大类；要求特殊的话回软件里点编辑单独改"],
    ]
    for r, row in enumerate(samples, 2):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
    widths = [24, 20, 34, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.cell(row=5, column=1, value="说明：表头请保留这四列（备注可留空），从第2行开始填，填完保存后在本软件中导入。")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def excel_export_bytes():
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "投递记录"
    headers = ["公司", "岗位", "收件邮箱", "备注", "状态", "发送时间", "失败原因",
               "我的进展", "重要日期", "我的笔记"]
    fills = PatternFill("solid", fgColor="185FA5")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fills
    for r, j in enumerate(get_jobs(), 2):
        row = [j.get("company"), j.get("position"), j.get("email"), j.get("note"),
               j.get("status"), j.get("sent_at", ""), j.get("error", ""),
               j.get("progress", "未回音"), j.get("event_date", ""), j.get("mynote", "")]
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
    widths = [24, 20, 34, 28, 10, 20, 36, 12, 14, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def excel_import(b64_data):
    from openpyxl import load_workbook
    raw = base64.b64decode(b64_data)
    wb = load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb.active

    header_map = {}
    for cell in ws[1]:
        v = str(cell.value or "").strip()
        if not v:
            continue
        if "公司" in v or "企业" in v:
            header_map["company"] = cell.column
        elif "岗位" in v or "职位" in v:
            header_map["position"] = cell.column
        elif "邮箱" in v or "mail" in v.lower():
            header_map["email"] = cell.column
        elif "备注" in v or "说明" in v:
            header_map["note"] = cell.column
        elif "进展" in v:
            header_map["progress"] = cell.column
        elif "日期" in v:
            header_map["event_date"] = cell.column
        elif "笔记" in v:
            header_map["mynote"] = cell.column
    if "email" not in header_map or ("company" not in header_map and "position" not in header_map):
        raise ValueError("表头无法识别：需要包含「公司/岗位/收件邮箱」列，请使用下载的模板填写")

    jobs = get_jobs()
    existing = {(j.get("email"), j.get("position"), j.get("company")) for j in jobs}
    added, skipped, bad = 0, 0, 0
    for row in ws.iter_rows(min_row=2):
        vals = {k: (str(row[c - 1].value or "").strip() if c - 1 < len(row) else "")
                for k, c in header_map.items()}
        company = vals.get("company", "")
        position = vals.get("position", "")
        email = vals.get("email", "")
        note = vals.get("note", "")
        if not email and not company and not position:
            continue
        if "@" not in email:
            bad += 1
            continue
        key = (email, position, company)
        if key in existing:
            skipped += 1
            continue
        existing.add(key)
        progress = vals.get("progress", "").strip()
        jobs.append({
            "id": "j%d" % int(time.time() * 1000) + str(added),
            "company": company, "position": position, "email": email, "note": note,
            "subject_override": "", "body_override": "", "atts": [],
            "progress": progress if progress in ("未回音", "笔试", "面试", "Offer", "挂了", "我放弃") else "未回音",
            "event_date": vals.get("event_date", "").strip(),
            "mynote": vals.get("mynote", ""),
            "status": "待发", "error": "", "sent_at": "",
        })
        added += 1
    save_jobs(jobs)
    return {"added": added, "skipped": skipped, "invalid": bad}


# ---------------------------------------------------------------- HTTP 服务
class Handler(BaseHTTPRequestHandler):
    server_version = "ResumeSender/1.0"

    def log_message(self, *args):
        pass

    # ---------- 基础 ----------
    def _json(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _file(self, path, ctype="text/html; charset=utf-8", download_name=None):
        try:
            with open(path, "rb") as f:
                data = f.read()
        except Exception:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        if download_name:
            self.send_header("Content-Disposition",
                             "attachment; filename*=UTF-8''%s" % quote(download_name))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        length = int(self.headers.get("Content-Length", 0) or 0)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            return json.loads(raw.decode("utf-8") or "{}")
        except Exception:
            return {}

    # ---------- GET ----------
    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path == "/" or path == "/index.html":
                self._file(INDEX_HTML)
            elif path == "/api/state":
                cfg = get_config()
                masked = dict(cfg)
                masked["auth_code"] = "******" if cfg.get("auth_code") else ""
                self._json({
                    "ok": True,
                    "config": masked,
                    "template": get_template(),
                    "send": {k: SEND_STATE[k] for k in ("running", "stop", "current", "done", "total")},
                    "today_sent": today_sent_count(),
                    "daily_limit": cfg.get("daily_limit", 50),
                    "attachments": list_attachments(),
                    "merge_files": merge_list(),
                })
            elif path == "/api/jobs":
                self._json({"ok": True, "jobs": get_jobs()})
            elif path == "/api/jobs/template":
                self._file_bytes(excel_template_bytes(), "简历投递清单模板.xlsx")
            elif path == "/api/jobs/export":
                self._file_bytes(excel_export_bytes(), "投递记录导出.xlsx")
            elif path == "/api/status":
                self._json({
                    "ok": True,
                    "running": SEND_STATE["running"],
                    "stop": SEND_STATE["stop"],
                    "current": SEND_STATE["current"],
                    "done": SEND_STATE["done"],
                    "total": SEND_STATE["total"],
                    "log": SEND_STATE["log"][-30:],
                    "today_sent": today_sent_count(),
                })
            else:
                self.send_error(404)
        except Exception as e:
            self._json({"ok": False, "error": str(e)}, 500)

    def _file_bytes(self, data, name):
        self.send_response(200)
        self.send_header("Content-Type",
                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition",
                         "attachment; filename*=UTF-8''%s" % quote(name))
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # ---------- POST ----------
    def do_POST(self):
        path = urlparse(self.path).path
        body = self._body()
        try:
            if path == "/api/config":
                cfg = get_config()
                for k in ("email", "auth_code", "display_name", "name",
                          "school", "major", "interval", "daily_limit"):
                    if k in body and body[k] != "":
                        v = body[k]
                        if k in ("interval", "daily_limit"):
                            v = max(5, int(v))
                        if k == "auth_code" and v == "******":
                            continue
                        cfg[k] = v
                save_json("config.json", cfg)
                self._json({"ok": True})
            elif path == "/api/template":
                tpl = get_template()
                if "subject" in body:
                    tpl["subject"] = str(body["subject"])
                if "body" in body:
                    tpl["body"] = str(body["body"])
                save_json("template.json", tpl)
                self._json({"ok": True})
            elif path == "/api/preview":
                job_id = body.get("job_id", "")
                jobs = get_jobs()
                job = next((j for j in jobs if j["id"] == job_id), None)
                if not job:
                    job = jobs[0] if jobs else {"company": "示例科技", "position": "后端开发",
                                                "email": "hr@example.com", "note": "",
                                                "subject_override": "", "body_override": "", "atts": []}
                cfg = get_config()
                tpl = get_template()
                display = cfg.get("display_name") or cfg.get("name") or "求职者"
                atts = resolve_attachments(job)
                self._json({
                    "ok": True,
                    "from": "%s <%s>" % (display, cfg.get("email", "未配置邮箱")),
                    "to": job.get("email", ""),
                    "subject": render(job.get("subject_override") or tpl.get("subject", ""), job, cfg),
                    "body": render(job.get("body_override") or tpl.get("body", ""), job, cfg),
                    "attachments": atts,
                    "custom": bool(job.get("subject_override") or job.get("body_override") or job.get("atts")),
                    "job": job,
                })
            elif path == "/api/jobs/add":
                jobs = get_jobs()
                job = {
                    "id": "j%d" % int(time.time() * 1000),
                    "company": str(body.get("company", "")).strip(),
                    "position": str(body.get("position", "")).strip(),
                    "email": str(body.get("email", "")).strip(),
                    "note": str(body.get("note", "")).strip(),
                    "subject_override": str(body.get("subject_override", "")).strip(),
                    "body_override": str(body.get("body_override", "")),
                    "atts": [str(a) for a in (body.get("atts") or [])],
                    "progress": str(body.get("progress", "")).strip() or "未回音",
                    "event_date": str(body.get("event_date", "")).strip(),
                    "mynote": str(body.get("mynote", "")),
                    "status": "待发", "error": "", "sent_at": "",
                }
                if "@" not in job["email"]:
                    self._json({"ok": False, "error": "收件邮箱格式不正确"})
                    return
                jobs.append(job)
                save_jobs(jobs)
                self._json({"ok": True, "job": job})
            elif path == "/api/jobs/update":
                jobs = get_jobs()
                jid = body.get("id")
                for j in jobs:
                    if j["id"] == jid:
                        for k in ("company", "position", "email", "note", "subject_override",
                                  "body_override", "mynote", "event_date"):
                            if k in body:
                                j[k] = str(body[k]).strip()
                        if "progress" in body:
                            v = str(body["progress"]).strip()
                            j["progress"] = v if v in ("未回音", "笔试", "面试", "Offer", "挂了", "我放弃") else "未回音"
                        if "atts" in body:
                            j["atts"] = [str(a) for a in (body.get("atts") or [])]
                        break
                save_jobs(jobs)
                self._json({"ok": True})
            elif path == "/api/jobs/delete":
                ids = set(body.get("ids", []))
                jobs = [j for j in get_jobs() if j["id"] not in ids]
                save_jobs(jobs)
                self._json({"ok": True})
            elif path == "/api/jobs/reset":
                jobs = get_jobs()
                for j in jobs:
                    if j["status"] in ("失败", "已发送"):
                        j["status"] = "待发"
                        j["error"] = ""
                save_jobs(jobs)
                self._json({"ok": True})
            elif path == "/api/jobs/import":
                result = excel_import(body.get("data", ""))
                self._json({"ok": True, **result})
            elif path == "/api/attachments":
                name = os.path.basename(str(body.get("name", "附件")))
                data = base64.b64decode(body.get("data", ""))
                if len(data) > 15 * 1024 * 1024:
                    self._json({"ok": False, "error": "附件超过 15MB，建议只发 PDF 简历"})
                    return
                target = os.path.join(ATTACH_DIR, name)
                i = 1
                base_name, ext = os.path.splitext(name)
                while os.path.exists(target):
                    target = os.path.join(ATTACH_DIR, "%s(%d)%s" % (base_name, i, ext))
                    i += 1
                with open(target, "wb") as f:
                    f.write(data)
                self._json({"ok": True, "attachments": list_attachments()})
            elif path == "/api/attachments/delete":
                name = os.path.basename(str(body.get("name", "")))
                p = os.path.join(ATTACH_DIR, name)
                if os.path.exists(p):
                    os.remove(p)
                drop_att_note(name)
                self._json({"ok": True, "attachments": list_attachments()})
            elif path == "/api/attachments/note":
                name = os.path.basename(str(body.get("name", "")))
                if name and os.path.isfile(os.path.join(ATTACH_DIR, name)):
                    set_att_note(name, str(body.get("note", "")))
                self._json({"ok": True, "attachments": list_attachments()})
            elif path == "/api/merge/add":
                data = base64.b64decode(body.get("data", ""))
                if not data:
                    self._json({"ok": False, "error": "文件是空的，没收到内容"})
                    return
                if len(data) > 30 * 1024 * 1024:
                    self._json({"ok": False, "error": "单个文件超过 30MB，请先压缩一下"})
                    return
                merge_add(body.get("name", ""), data)
                self._json({"ok": True, "files": merge_list()})
            elif path == "/api/merge/move":
                merge_move(str(body.get("id", "")), int(body.get("dir", 0) or 0))
                self._json({"ok": True, "files": merge_list()})
            elif path == "/api/merge/remove":
                merge_remove(str(body.get("id", "")))
                self._json({"ok": True, "files": merge_list()})
            elif path == "/api/merge/clear":
                merge_clear()
                self._json({"ok": True, "files": merge_list()})
            elif path == "/api/merge/open-tool":
                try:
                    open_merge_tool(str(body.get("name", "")))
                except ValueError as e:
                    self._json({"ok": False, "error": str(e)})
                    return
                self._json({"ok": True})
            elif path == "/api/merge/run":
                try:
                    out, pages, size = merge_run(str(body.get("name", "")))
                except ValueError as e:
                    self._json({"ok": False, "error": str(e)})
                    return
                merge_clear()
                self._json({"ok": True, "output": out, "pages": pages, "size": size,
                            "attachments": list_attachments(), "files": merge_list()})
            elif path == "/api/start":
                if SEND_STATE["running"]:
                    self._json({"ok": False, "error": "已有投递任务在进行中"})
                    return
                cfg = get_config()
                if not cfg.get("email") or not cfg.get("auth_code"):
                    self._json({"ok": False, "error": "请先在「发送设置」中填写 QQ 邮箱和授权码"})
                    return
                tpl = get_template()
                if not tpl.get("subject") or not tpl.get("body"):
                    self._json({"ok": False, "error": "请先在「邮件模板」中填写主题和正文"})
                    return
                ids = body.get("ids", [])
                jobs = get_jobs()
                valid = [j["id"] for j in jobs if j["id"] in set(ids) and j["status"] != "已发送"]
                if not valid:
                    self._json({"ok": False, "error": "没有可投递的目标（已发送的不会重复投递）"})
                    return
                for j in jobs:
                    if j["id"] in set(valid) and j["status"] == "失败":
                        j["status"] = "待发"
                        j["error"] = ""
                save_jobs(jobs)
                threading.Thread(target=sender_worker, args=(valid,), daemon=True).start()
                self._json({"ok": True, "count": len(valid)})
            elif path == "/api/stop":
                SEND_STATE["stop"] = True
                self._json({"ok": True})
            else:
                self.send_error(404)
        except Exception as e:
            self._json({"ok": False, "error": str(e)}, 500)


def main():
    init_dirs()
    server = None
    port = 8765
    for p in (8765, 8766, 8767, 8768):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", p), Handler)
            port = p
            break
        except OSError:
            continue
    if server is None:
        print("错误：8765-8768 端口都被占用，无法启动")
        return
    url = "http://127.0.0.1:%d" % port
    print("=" * 50)
    print("  简历批量投递助手 已启动")
    print("  请在浏览器访问：%s" % url)
    print("  关闭本窗口即退出软件（数据都保存在本地）")
    print("=" * 50)
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
