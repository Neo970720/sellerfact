import json, urllib.request, urllib.error, sys
token = open('/workspace/.at_script').read().strip()
sid = "12Wc3kc0n6vRm9DHvlcimQk0K8N4hnC1nSPEF6wKRx8K44LZfTzL2Su_x"
code = open('/workspace/work/sellerfact/intake/Code.gs', encoding='utf-8').read()
manifest = {
  "timeZone": "Asia/Seoul",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {"executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS"},
  "oauthScopes": [
    "https://www.googleapis.com/auth/script.send_mail",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email"
  ]
}
body = json.dumps({"files": [
  {"name": "appsscript", "type": "JSON", "source": json.dumps(manifest, ensure_ascii=False)},
  {"name": "Code", "type": "SERVER_JS", "source": code},
]}).encode()
req = urllib.request.Request(f"https://script.googleapis.com/v1/projects/{sid}/content", data=body, method="PUT",
    headers={"Authorization": "Bearer " + token, "Content-Type": "application/json"})
try:
    r = urllib.request.urlopen(req, timeout=60)
    print(r.status, r.read()[:400].decode())
except urllib.error.HTTPError as e:
    print("ERR", e.code, e.read()[:1000].decode())
