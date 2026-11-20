import json, urllib.request, urllib.error
token = open('/workspace/.at_script').read().strip()
sid = "12Wc3kc0n6vRm9DHvlcimQk0K8N4hnC1nSPEF6wKRx8K44LZfTzL2Su_x"
H = {"Authorization": "Bearer " + token, "Content-Type": "application/json"}

def call(method, url, payload=None):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers=H)
    try:
        r = urllib.request.urlopen(req, timeout=60)
        return json.loads(r.read().decode() or "{}")
    except urllib.error.HTTPError as e:
        return {"error": e.code, "body": e.read()[:800].decode()}

ver = call("POST", f"https://script.googleapis.com/v1/projects/{sid}/versions", {})
print("VERSION", json.dumps(ver, ensure_ascii=False)[:300])
vnum = ver.get("versionNumber")
if vnum:
    dep = call("POST", f"https://script.googleapis.com/v1/projects/{sid}/deployments",
               {"versionNumber": vnum, "manifestFileName": "appsscript", "description": "SellerFact intake v1"})
    print("DEPLOYMENT", json.dumps(dep, ensure_ascii=False)[:900])
