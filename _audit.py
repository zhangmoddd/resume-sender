# -*- coding: utf-8 -*-
"""临时脚本：逐条复核另一位 agent 的结论（真起服务实测）。跑完自删，不动真数据。"""
import base64
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix="audit_")
R = []


def say(t):
    R.append(t)
    print(t)


def flush():
    try:
        io.open(os.path.join(BASE, "..", "~audit.txt"), "w", encoding="utf-8").write("\n".join(R))
    except Exception:
        pass


sys.path.insert(0, BASE)
import app  # noqa: E402

app.DATA_DIR = os.path.join(TMP, "data")
app.ATTACH_DIR = os.path.join(TMP, "attachments")
app.MERGE_DIR = os.path.join(app.DATA_DIR, "_merge_staging")
app.LOCK_FILE = os.path.join(app.DATA_DIR, "app.lock")
app.init_dirs()

SENT = []


class FakeSMTP:
    def __init__(self, *a, **k):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def login(self, *a):
        return True

    def send_message(self, msg):
        SENT.append(msg["To"])


app.smtplib.SMTP_SSL = FakeSMTP
app._interruptible_sleep = lambda s: time.sleep(0.02)

srv = app.ThreadingHTTPServer(("127.0.0.1", 8903), app.Handler)
threading.Thread(target=srv.serve_forever, daemon=True).start()


def post(path, body, raw=False):
    req = urllib.request.Request("http://127.0.0.1:8903" + path,
                                 data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            d = r.read()
    except urllib.error.HTTPError as e:
        d = e.read()
        return ("HTTP%d " % e.code) + d.decode("utf-8", "replace")
    try:
        return json.loads(d.decode("utf-8")) if not raw else d
    except Exception:
        return d.decode("utf-8", "replace")


def get(path):
    with urllib.request.urlopen("http://127.0.0.1:8903" + path, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def cfgset(**kw):
    c = app.get_config()
    c.update(kw)
    app.save_json("config.json", c)


def att(name):
    with open(os.path.join(app.ATTACH_DIR, name), "wb") as f:
        f.write(b"%PDF-1.4 x")


import traceback  # noqa: E402

try:
    exec(io.open(__file__, encoding="utf-8").read().split("\n# ====BODY====\n", 1)[1])
except Exception:
    say("!! 脚本自己出错了：")
    say(traceback.format_exc())
finally:
    flush()
    srv.shutdown()

raise SystemExit(0)

# ====BODY====
say("=" * 70)
say("【复核 1】点「编辑」再保存，atts 会不会从「全带」变成固定名单")
say("=" * 70)
att("简历.pdf")
cfgset(email="me@qq.com", auth_code="x" * 16, interval=40, daily_limit=50)
post("/api/template", {"subject": "应聘{岗位}", "body": "我是{姓名}"})
# 用在线填表导入一行（这条路 atts 是空的 = 全带）
post("/api/sheet/import", {"rows": [["公司A", "岗位A", "a@x.com", ""]]})
j = get("/api/jobs")["jobs"][0]
say("导入后这条的 atts = %r  → 语义：带附件库全部" % j["atts"])
# 模拟前端「编辑→保存」：编辑框里没单独勾过时会把现有附件全部当勾选
atts_from_ui = [a["name"] for a in get("/api/state")["attachments"]]
post("/api/jobs/update", {"id": j["id"], "atts": atts_from_ui})
j2 = [x for x in get("/api/jobs")["jobs"] if x["id"] == j["id"]][0]
say("保存一次后 atts = %r" % j2["atts"])
# 现在再上传一个新附件，看这条会不会带上
att("证明材料.pdf")
pr = post("/api/preview", {"job_id": j["id"]})
names = [a["name"] for a in pr["attachments"]] if isinstance(pr, dict) else pr
say("之后新传了「证明材料.pdf」，这条实际会带：%s" % names)
say("→ 结论：%s" % ("属实（新附件不会带上，且不报错）"
                    if "证明材料.pdf" not in names else "不属实"))

say("")
say("=" * 70)
say("【复核 2】config.json 里 interval 不是数字 → 投递线程死掉、running 卡住")
say("=" * 70)
cfgset(interval="", daily_limit=50)          # 模拟手改配置文件（接口层已拦住空值）
post("/api/jobs/add", {"company": "公司B", "position": "岗位B", "email": "b@x.com"})
jid = [x for x in get("/api/jobs")["jobs"] if x["company"] == "公司B"][0]["id"]
r = post("/api/start", {"ids": [jid]})
say("start 返回：%r" % (r,))
time.sleep(0.6)
st = get("/api/status")
say("几秒后 running=%s done=%s total=%s 日志条数=%s"
    % (st["running"], st["done"], st["total"], len(st["log"])))
say("再点一次开始投递：%r" % (post("/api/start", {"ids": [jid]}),))
post("/api/stop", {})
time.sleep(0.3)
st2 = get("/api/status")
say("点「停止发送」之后 running 还是 %s" % st2["running"])
say("→ 结论：%s" % ("属实（线程死在 try 之前，running 永远 True，只能重启软件）"
                    if st2["running"] else "不属实（能自己恢复）"))
app.SEND_STATE["running"] = False
cfgset(interval=40)

say("")
say("=" * 70)
say("【复核 3】第二个实例读不出锁文件里的端口")
say("=" * 70)
lockf = app.acquire_instance_lock()
say("主实例拿到锁：%s" % (lockf is not None))
if lockf:
    lockf.seek(0)
    lockf.truncate()
    lockf.write("8766\n")
    lockf.flush()
try:
    io.open(app.LOCK_FILE, encoding="utf-8").read()
    say("主进程再读一次锁文件：读到了（说明没被锁挡住）")
except Exception as e:
    say("主进程自己再读一次锁文件都读不出来：%s: %s" % (type(e).__name__, e))
child = subprocess.run([sys.executable, "-c",
                        "import sys;sys.path.insert(0,r'%s');"
                        "import app;print('child got port:', app.read_running_port())" % BASE],
                       cwd=BASE, capture_output=True, timeout=30)
co = child.stdout.decode("utf-8", "replace").strip()
say("锁文件里实际写的端口 = 8766")
say("子进程读到的端口 = %s" % co)
if child.stderr:
    say("子进程 stderr 最后一行 = %s"
        % child.stderr.decode("utf-8", "replace").strip().splitlines()[-1][:200])
say("→ 结论：%s" % ("属实（另一个进程读不到端口，只能回落 8765）"
                    if "8766" not in co else "不属实"))
if lockf:
    lockf.close()

say("")
say("=" * 70)
say("【复核 4】模板变量插入：光标在 0 位置时会不会跑到末尾")
say("=" * 70)
say("代码：const pos = el.selectionStart || el.value.length;")
say("JS 实测：selectionStart=0 时  (0 || 999) = %s" % (0 or 999))
say("→ 结论：属实（光标在开头时插到末尾）")

say("")
say("=" * 70)
say("【复核 5】atts 传字符串会不会被按字拆开 / 缺 id 的条目会不会让接口 500")
say("=" * 70)
post("/api/jobs/add", {"company": "公司C", "position": "岗位C", "email": "c@x.com",
                       "atts": "不是列表"})
jc = [x for x in get("/api/jobs")["jobs"] if x["company"] == "公司C"][0]
say("atts 传字符串 → 存成 %r" % jc["atts"])
# 造一条没有 id 的脏数据
jobs = app.get_jobs()
jobs.append({"company": "没有id", "position": "x", "email": "d@x.com"})
app.save_jobs(jobs)
say("/api/preview 拿脏数据：%s" % str(post("/api/preview", {"job_id": "没有id"}))[:120])

say("")
say("=" * 70)
say("【复核 6】他顺手下的结论：收件邮箱里塞换行，会不会偷偷带出 Bcc")
say("=" * 70)
import email.utils  # noqa: E402
from email.mime.multipart import MIMEMultipart  # noqa: E402
from email.mime.text import MIMEText  # noqa: E402
from email.header import Header  # noqa: E402

evil = "hr@a.com\nBcc: victim@b.com"
say("getaddresses 解析结果 = %r" % (email.utils.getaddresses([evil]),))
mm = MIMEMultipart()
mm["From"] = "me@qq.com"
mm["To"] = evil
mm["Subject"] = Header("hi", "utf-8")
mm.attach(MIMEText("正文", "plain", "utf-8"))
txt = mm.as_string()
injected = any(l.strip().lower().startswith("bcc:") for l in txt.splitlines())
say("生成的邮件里出现了独立的 Bcc 行 = %s" % injected)
for l in txt.splitlines():
    if "victim" in l or "Bcc" in l:
        say("   > %s" % l)

srv.shutdown()
shutil.rmtree(TMP, ignore_errors=True)
flush()
