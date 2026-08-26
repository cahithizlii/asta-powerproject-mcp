"""p6_mcp sunucusunu gercek bir MCP stdio istemcisi gibi surer.

Import testi degil: sureci baslatir, JSON-RPC ile initialize + tools/list +
tools/call yapar. Boylece .claude.json kaydinin calisacagi kanitlanir.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading

PY = os.environ.get("P6_PYTHON", sys.executable)
SERVER = os.path.join(os.environ.get("P6_REPO", os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))), "p6_mcp_core.py")


def main() -> int:
    env = dict(os.environ)
    if os.environ.get("JAVA_HOME"):
        env["JAVA_HOME"] = os.environ["JAVA_HOME"]
    env["PYTHONIOENCODING"] = "utf-8"

    proc = subprocess.Popen(
        [PY, SERVER], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, env=env, text=True, encoding="utf-8", bufsize=1)

    errbuf: list[str] = []
    threading.Thread(target=lambda: [errbuf.append(l) for l in proc.stderr],
                     daemon=True).start()

    def send(obj) -> None:
        proc.stdin.write(json.dumps(obj) + "\n")
        proc.stdin.flush()

    def recv(want_id):
        while True:
            line = proc.stdout.readline()
            if not line:
                raise SystemExit("sunucu kapandi. stderr:\n" + "".join(errbuf[-25:]))
            line = line.strip()
            if not line:
                continue
            msg = json.loads(line)
            if msg.get("id") == want_id:
                return msg

    failures: list[str] = []

    send({"jsonrpc": "2.0", "id": 1, "method": "initialize",
          "params": {"protocolVersion": "2024-11-05",
                     "capabilities": {},
                     "clientInfo": {"name": "smoke", "version": "1"}}})
    init = recv(1)
    info = init.get("result", {}).get("serverInfo", {})
    print("1) initialize -> %s %s" % (info.get("name"), info.get("version")))
    if not info.get("name"):
        failures.append("initialize serverInfo bos")

    send({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})

    send({"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}})
    tools = recv(2).get("result", {}).get("tools", [])
    names = [t["name"] for t in tools]
    print("2) tools/list -> %s" % names)
    for want in ("p6_job", "p6_query"):
        if want not in names:
            failures.append(want + " tool listede yok")

    send({"jsonrpc": "2.0", "id": 3, "method": "tools/call",
          "params": {"name": "p6_job",
                     "arguments": {"params": {"action": "service_health"}}}})
    res = recv(3)
    text = ""
    for item in res.get("result", {}).get("content", []):
        text += item.get("text", "")
    print("3) tools/call service_health ->")
    for line in text.splitlines()[:14]:
        print("     " + line)
    try:
        payload = json.loads(text)
    except Exception:
        payload = {}
        failures.append("service_health JSON degil")
    if payload.get("driver") != "SQLServer":
        failures.append("driver SQLServer degil: %r" % payload.get("driver"))
    if not payload.get("service", {}).get("running"):
        failures.append("PrmJobSv calismiyor")

    send({"jsonrpc": "2.0", "id": 4, "method": "tools/call",
          "params": {"name": "p6_job",
                     "arguments": {"params": {"action": "job_data",
                                              "proj_id": 368}}}})
    res4 = recv(4)
    t4 = "".join(i.get("text", "") for i in res4.get("result", {}).get("content", []))
    expected = "(0||(Default Project|368)((0||Schedule Projects()((0||368()())))))"
    print("4) tools/call job_data -> %s" % ("BIREBIR" if expected in t4 else "FARKLI"))
    if expected not in t4:
        failures.append("job_data blob referansla eslesmiyor")

    proc.stdin.close()
    try:
        proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()

    print()
    if failures:
        print("BASARISIZ:")
        for f in failures:
            print("   - " + f)
        if errbuf:
            print("stderr kuyrugu:")
            for l in errbuf[-10:]:
                print("   " + l.rstrip())
        return 1
    print("MCP STDIO TESTI GECTI")
    return 0


if __name__ == "__main__":
    sys.exit(main())
