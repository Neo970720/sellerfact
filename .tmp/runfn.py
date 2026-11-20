import json, sys, urllib.request, urllib.error
token = open('/workspace/.at_script').read().strip()
sid = "12Wc3kc0n6vRm9DHvlcimQk0K8N4hnC1nSPEF6wKRx8K44LZfTzL2Su_x"
fn = sys.argv[1]
params = sys.argv[2:]
payload = {"devMode": True, "runtimeVersion": "V8"}
req = urllib.request.Request(f"https://script.googleapis.com/v1/scripts/{sid}:run",
    data=json.dumps(payload).encode(), method="POST",
    headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"})
# functionName goes in the URL path form too
url = f"https://script.googleapis.com/v1/scripts/{sid}:run"
payload = {"function": fn, "devMode": True, "parameters": [{"string": p} for p in params]}
req = urllib.request.Request(url, data=json.dumps(payload).encode(), method="POST",
    headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"})
try:
    r = urllib.request.urlopen(req, timeout=90)
    print(r.read().decode()[:1500])
except urllib.error.HTTPError as e:
    print("ERR", e.code, e.read().decode()[:1200])
