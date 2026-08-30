# -*- coding: utf-8 -*-
"""log_store.py —— 转换日志库(服务端专属, 2026-08-30)。

SQLite 单文件零依赖; 每次转换请求记录一条(时间/文件/耗时/通道/各 API 消耗),
保留 30 天(purge 自动清理, 每小时至多执行一次)。线程安全(全局锁)。

表结构 logs: ts(ISO 请求时间), filename, size, status(ok/fail), stage(失败阶段),
mode(通道: rule/glm-ocr/glm-ocr+baidu/ocr-sandwich), duration_s, rows,
glm_calls/glm_prompt/glm_completion(GLM-OCR token), vlm_calls/vlm_prompt/
vlm_completion(视觉兜底 token), baidu_calls(百度点数=25/次), error。
"""
import json
import os
import sqlite3
import threading
import time
from datetime import datetime, timedelta

RETENTION_DAYS = 30
_lock = threading.Lock()
_conn = None
_last_purge = [0.0]

COLUMNS = ("id", "ts", "filename", "size", "status", "stage", "mode",
           "duration_s", "rows", "glm_calls", "glm_prompt", "glm_completion",
           "vlm_calls", "vlm_prompt", "vlm_completion", "baidu_calls", "error")

DDL = """CREATE TABLE IF NOT EXISTS logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  ts TEXT NOT NULL,
  filename TEXT,
  size INTEGER DEFAULT 0,
  status TEXT,
  stage TEXT,
  mode TEXT,
  duration_s REAL DEFAULT 0,
  rows INTEGER,
  glm_calls INTEGER DEFAULT 0,
  glm_prompt INTEGER DEFAULT 0,
  glm_completion INTEGER DEFAULT 0,
  vlm_calls INTEGER DEFAULT 0,
  vlm_prompt INTEGER DEFAULT 0,
  vlm_completion INTEGER DEFAULT 0,
  baidu_calls INTEGER DEFAULT 0,
  error TEXT
)"""


def _connect():
    global _conn
    if _conn is None:
        path = os.environ.get("LOG_DB_PATH",
                              os.path.join(os.path.dirname(
                                  os.path.abspath(__file__)), "logs", "convert.db"))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        _conn = sqlite3.connect(path, check_same_thread=False)
        _conn.execute(DDL)
        _conn.commit()
    return _conn


def add_entry(filename, size, status, stage, mode, duration_s, rows=None,
              usage=None, error=None):
    """记录一条转换日志。usage: {"glm_calls","glm_prompt","glm_completion",
    "vlm_calls","vlm_prompt","vlm_completion","baidu_calls"}(缺省补 0)。"""
    u = usage or {}
    try:
        with _lock:
            conn = _connect()
            conn.execute(
                "INSERT INTO logs (ts, filename, size, status, stage, mode,"
                " duration_s, rows, glm_calls, glm_prompt, glm_completion,"
                " vlm_calls, vlm_prompt, vlm_completion, baidu_calls, error)"
                " VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (datetime.now().isoformat(timespec="seconds"),
                 str(filename or "")[:200], int(size or 0), status, stage,
                 mode, round(float(duration_s or 0), 2), rows,
                 int(u.get("glm_calls", 0)), int(u.get("glm_prompt", 0)),
                 int(u.get("glm_completion", 0)), int(u.get("vlm_calls", 0)),
                 int(u.get("vlm_prompt", 0)), int(u.get("vlm_completion", 0)),
                 int(u.get("baidu_calls", 0)),
                 str(error or "")[:300]))
            conn.commit()
        maybe_purge()
    except Exception as e:  # noqa: BLE001 (日志失败绝不影响转换)
        print(f"[提示] 日志写入失败: {e}", flush=True)


def maybe_purge(force=False):
    """清理超过保留期的日志(每小时至多一次; force 强制)。"""
    now = time.time()
    if not force and now - _last_purge[0] < 3600:
        return
    with _lock:
        if not force and now - _last_purge[0] < 3600:
            return
        _last_purge[0] = now
        try:
            conn = _connect()
            cutoff = (datetime.now() - timedelta(days=RETENTION_DAYS)
                      ).isoformat(timespec="seconds")
            conn.execute("DELETE FROM logs WHERE ts < ?", (cutoff,))
            conn.commit()
        except Exception as e:  # noqa: BLE001
            print(f"[提示] 日志清理失败: {e}", flush=True)


def query(days=30, status=None, limit=500, offset=0):
    """查询日志(新→旧)。"""
    with _lock:
        conn = _connect()
        cutoff = (datetime.now() - timedelta(days=days)
                  ).isoformat(timespec="seconds")
        sql = "SELECT * FROM logs WHERE ts >= ?"
        args = [cutoff]
        if status in ("ok", "fail"):
            sql += " AND status = ?"
            args.append(status)
        sql += " ORDER BY id DESC LIMIT ? OFFSET ?"
        args += [int(limit), int(offset)]
        # 列名硬编码: 容器内 sqlite3 cursor.description 会返回全 None(踩坑)
        return [dict(zip(COLUMNS, row)) for row in conn.execute(sql, args)]


def stats(days=30):
    """汇总统计: 总量/成功/失败/耗时合计/各 API 消耗/按日分布。"""
    with _lock:
        conn = _connect()
        cutoff = (datetime.now() - timedelta(days=days)
                  ).isoformat(timespec="seconds")

    def one(sql, args=()):
        with _lock:
            row = _connect().execute(sql, args).fetchone()
        return row

    total, ok, fail = one(
        "SELECT COUNT(*), SUM(status='ok'), SUM(status='fail')"
        " FROM logs WHERE ts >= ?", (cutoff,))
    dur, glc, glp, glm_cpt, vc, vp, vcp, bc = one(
        "SELECT IFNULL(SUM(duration_s),0), IFNULL(SUM(glm_calls),0),"
        " IFNULL(SUM(glm_prompt),0), IFNULL(SUM(glm_completion),0),"
        " IFNULL(SUM(vlm_calls),0), IFNULL(SUM(vlm_prompt),0),"
        " IFNULL(SUM(vlm_completion),0), IFNULL(SUM(baidu_calls),0)"
        " FROM logs WHERE ts >= ?", (cutoff,))
    daily = []
    with _lock:
        rows = _connect().execute(
            "SELECT substr(ts,1,10) d, COUNT(*), SUM(status='ok')"
            " FROM logs WHERE ts >= ? GROUP BY d ORDER BY d", (cutoff,))
        daily = [{"day": r[0], "total": r[1], "ok": r[2] or 0}
                 for r in rows]
    return {
        "total": total or 0, "ok": ok or 0, "fail": fail or 0,
        "duration_sum": round(dur or 0, 1),
        "glm": {"calls": glc or 0, "prompt": glp or 0,
                "completion": glm_cpt or 0},
        "vlm": {"calls": vc or 0, "prompt": vp or 0,
                "completion": vcp or 0},
        "baidu_calls": bc or 0,
        "baidu_points": (bc or 0) * 25,
        "daily": daily[-14:],
    }


def export_json(days=30):
    return json.dumps(query(days=days, limit=100000), ensure_ascii=False,
                      indent=1)
