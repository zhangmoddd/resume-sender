# -*- coding: utf-8 -*-
"""临时脚本：验证这一轮 6 处修复。跑完自删，不动真数据。"""
import base64
import io
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import traceback
import urllib.error
import urllib.request

BASE = os.path.dirname(os.path.abspath(__file__))
TMP = tempfile.mkdtemp(prefix="verify2_")
R = []


def say(t):
    R.append(str(t))
    print(t)


def flush():
    try:
        io.open(os.path.join(BASE, "..", "~verify2.txt"), "w", encoding="utf-8").write("\n".join(R))
    except Exception:
        pass


sys.path.insert(0, BASE)
import app  # noqa: E402

app.DATA_DIR = os.path.join(TMP, "data")
app.ATTACH_DIR = os.path.join(TMP, "attachments")
app.MERGE_DIR = os.path.join(app.DATA_DIR, "_merge_staging")
app.LOCK_FILE = os.path.join(app.DATA_DIR, "app.lock")
app.init_dirs()

PORT = 8905
app.APP_PORTS = (PORT,)          # 让"找自己"的函数只探这一个端口
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

srv = app.ThreadingHTTPServer(("127.0.0.1", PORT), app.Handler)
threading.Thread(target=srv.serve_forever, daemon=True).start()
time.sleep(0.2)


def post(path, body):
    req = urllib.request.Request("http://127.0.0.1:%d%s" % (PORT, path),
                                 data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            d = r.read()
    except urllib.error.HTTPError as e:
        return {"_http": e.code, "_body": e.read().decode("utf-8", "replace")[:200]}
    try:
        return json.loads(d.decode("utf-8"))
    except Exception:
        return {"_raw": d.decode("utf-8", "replace")[:200]}


def get(path):
    with urllib.request.urlopen("http://127.0.0.1:%d%s" % (PORT, path), timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


def cfgset(**kw):
    c = app.get_config()
    c.update(kw)
    app.save_json("config.json", c)


def att(name):
    with open(os.path.join(app.ATTACH_DIR, name), "wb") as f:
        f.write(b"%PDF-1.4 x")


def wait_done(t=15):
    t0 = time.time()
    while time.time() - t0 < t:
        if not app.SEND_STATE["running"]:
            return True
        time.sleep(0.05)
    return False


try:
    att("简历.pdf")
    cfgset(email="me@qq.com", auth_code="x" * 16, interval=40, daily_limit=50)
    post("/api/template", {"subject": "应聘{岗位}", "body": "我是{姓名}"})

    say("=" * 62)
    say("修复1：点编辑再保存，附件还会不会从「跟着附件库走」变固定名单")
    say("=" * 62)
    post("/api/sheet/import", {"rows": [["公司A", "岗位A", "a@x.com", ""]]})
    j = get("/api/jobs")["jobs"][0]
    say("导入后 atts=%r（空=跟着附件库走）" % j["atts"])
    # 前端勾着「跟着附件库走」时提交的就是空名单
    post("/api/jobs/update", {"id": j["id"], "atts": [], "note": "改了个备注"})
    j2 = [x for x in get("/api/jobs")["jobs"] if x["id"] == j["id"]][0]
    att("证明材料.pdf")            # 之后才上传的新附件
    pr = post("/api/preview", {"job_id": j["id"]})
    names = [a["name"] for a in pr.get("attachments", [])]
    say("保存后 atts=%r；新传「证明材料.pdf」后实际会带：%s" % (j2["atts"], names))
    say("→ %s" % ("已修好：新附件照样带上" if "证明材料.pdf" in names else "!! 还是没带上"))

    # 反向：明确只勾两个的固定名单，新附件不该自动带上
    post("/api/jobs/update", {"id": j["id"], "atts": ["简历.pdf"]})
    pr = post("/api/preview", {"job_id": j["id"]})
    names2 = [a["name"] for a in pr.get("attachments", [])]
    say("改成固定只发「简历.pdf」后：%s" % names2)
    say("→ %s" % ("固定名单仍然有效（只发勾的）" if names2 == ["简历.pdf"] else "!! 固定名单失效"))

    say("")
    say("=" * 62)
    say("修复2：config 里 interval 变成空字符串，会不会卡死")
    say("=" * 62)
    cfgset(interval="", daily_limit=50)
    post("/api/jobs/add", {"company": "公司B", "position": "岗位B", "email": "b@x.com"})
    jid = [x for x in get("/api/jobs")["jobs"] if x["company"] == "公司B"][0]["id"]
    r = post("/api/start", {"ids": [jid]})
    say("start 返回：%r" % (r,))
    done = wait_done()
    st = get("/api/status")
    say("几秒后 running=%s done=%s total=%s 日志=%d 条"
        % (st["running"], st["done"], st["total"], len(st["log"])))
    say("日志内容：%s" % [l["text"] for l in st["log"]][:3])
    say("→ %s" % ("已修好：按默认 40 秒继续发，没卡死" if done else "!! 还是卡住了"))
    cfgset(interval=40)

    say("")
    say("=" * 62)
    say("修复3：第二个实例能不能找到正在跑的那个端口")
    say("=" * 62)
    found = app.find_running_instance()
    say("本软件跑在 %d，探测结果 = %s" % (PORT, found))
    say("→ %s" % ("已修好：能认出自己的端口" if found == PORT else "!! 还是认不出"))
    app.APP_PORTS = (8911, 8912)      # 换成没人在跑的端口
    say("换成没人在跑的端口后探测 = %s（应该是 None）" % app.find_running_instance())
    app.APP_PORTS = (PORT,)

    say("")
    say("=" * 62)
    say("修复5：atts 传字符串 / 缺 id 的脏数据")
    say("=" * 62)
    post("/api/jobs/add", {"company": "公司C", "position": "岗位C", "email": "c@x.com",
                           "atts": "不是列表"})
    jc = [x for x in get("/api/jobs")["jobs"] if x["company"] == "公司C"][0]
    say("atts 传字符串 → 存成 %r" % jc["atts"])
    say("→ %s" % ("已修好：不再被按字拆开" if jc["atts"] == ["不是列表"] else "!! 还是被拆了"))
    jobs = app.get_jobs()
    jobs.append({"company": "没有id", "position": "x", "email": "d@x.com"})
    app.save_jobs(jobs)
    r = post("/api/preview", {"job_id": "没有id"})
    say("/api/preview 拿缺 id 的脏数据 → %s" % ("ok=%s" % r.get("ok") if "ok" in r else r))
    say("→ %s" % ("已修好：不再 500" if r.get("ok") else "!! 还是报错"))

    say("")
    say("=" * 62)
    say("顺带：邮箱里带换行/怪字符会不会漏进去")
    say("=" * 62)
    r = post("/api/jobs/add", {"company": "公司D", "position": "岗位D",
                               "email": "hr@a.com\nBcc: victim@b.com"})
    say("带换行+冒号的邮箱 → %s" % (r,))
    say("→ %s" % ("已挡住：当场拒绝并说清哪里不对"
                  if r.get("ok") is False and "不该出现" in (r.get("error") or "") else "!! 没挡住"))
    r = post("/api/jobs/add", {"company": "公司D2", "position": "岗位D2", "email": "hr@normal.com"})
    say("正常邮箱仍然能加进来 = %s" % (r.get("ok") is True))

    say("")
    say("=" * 62)
    say("修复6：SMTP 超时")
    say("=" * 62)
    src = io.open(os.path.join(BASE, "app.py"), encoding="utf-8").read()
    say("send_one 里的超时 = %s" % ("90 秒" if "SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=90)" in src
                                    else "?? 没找到 90"))
    say("sender_worker 开头在 try 里 = %s" % ("是" if "    try:\n        cfg = get_config()" in src
                                               else "不是"))

    say("")
    say("=" * 62)
    say("回归：之前修过的还灵不灵")
    say("=" * 62)
    def find(company):
        for x in get("/api/jobs")["jobs"]:
            if x.get("company") == company:
                return x
        return {}

    SENT.clear()
    # 邮箱校验
    r = post("/api/jobs/add", {"company": "公司X", "position": "岗", "email": "hr@a.comBcc:v@b.com"})
    say("邮箱里带冒号 → 拒绝并说清楚：%s" % (r,))
    r2 = post("/api/sheet/import", {"rows": [["公司Y", "岗", "a@b.com\nBcc: v@b.com", ""]]})
    say("在线导入这种行 → 计入无效条数：%s" % (r2,))

    # 附件缺失拦截（用一条全新的待发行）
    post("/api/jobs/add", {"company": "公司G", "position": "岗位G", "email": "g@x.com",
                           "atts": ["根本没这个.pdf"]})
    jg = find("公司G").get("id")
    r = post("/api/start", {"ids": [jg]})
    say("附件对不上仍然拦下 = %s" % (r.get("ok") is False and "找不到" in (r.get("error") or "")))

    # 发信期间改清单不被覆盖
    post("/api/jobs/add", {"company": "公司E", "position": "岗位E", "email": "e@x.com"})
    je = find("公司E").get("id")
    post("/api/start", {"ids": [je]})
    time.sleep(0.1)
    app.mutate_job(je, lambda x: x.update({"progress": "面试"}))
    wait_done()
    time.sleep(0.2)
    fin = find("公司E")
    say("发信期间改的进展没被覆盖 = %s（%s）" % (fin.get("progress") == "面试", fin.get("progress")))

    # 连点两次
    SENT.clear()
    post("/api/jobs/add", {"company": "公司F", "position": "岗位F", "email": "f@x.com"})
    jf = find("公司F").get("id")
    out = {}


    def hit(k):
        out[k] = post("/api/start", {"ids": [jf]})


    ts = [threading.Thread(target=hit, args=(i,)) for i in (1, 2)]
    [t.start() for t in ts]
    [t.join() for t in ts]
    say("并发两次开始投递，只接受一个 = %s" % (sum(1 for v in out.values() if v.get("ok")) == 1))
    wait_done()
    say("这轮总共真发出去的邮件 = %s" % SENT)
except Exception:
    say("!! 脚本出错：")
    say(traceback.format_exc())
finally:
    flush()
    try:
        srv.shutdown()
    except Exception:
        pass
    shutil.rmtree(TMP, ignore_errors=True)

raise SystemExit(0)
