import json, sys, urllib.request, urllib.error
token = open('/workspace/.at_script').read().strip()
sid = "12Wc3kc0n6vRm9DHvlcimQk0K8N4hnC1nSPEF6wKRx8K44LZfTzL2Su_x"
H = {"Authorization": "Bearer " + token, "Content-Type": "application/json"}

def call(method, url, payload=None):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers=H)
    try:
        return json.loads(urllib.request.urlopen(req, timeout=90).read().decode() or "{}")
    except urllib.error.HTTPError as e:
        return {"error": e.code, "body": e.read().decode()[:600]}

v = call("POST", f"https://script.googleapis.com/v1/projects/{sid}/versions", {})
print("VER", v.get("versionNumber"), str(v)[:160])
fn = sys.argv[1]; params = sys.argv[2:]
res = call("POST", f"https://script.googleapis.com/v1/scripts/{sid}:run",
           {"function": fn, "devMode": False, "parameters": [{"string": p} for p in params]})
print(json.dumps(res, ensure_ascii=False)[:1500])
