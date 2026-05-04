import urllib.request, json

base = "http://localhost:5000"
video_id = "5127cde4"
video_filename = "video_5127cde4.mp4"

req = urllib.request.Request(
    f"{base}/api/generate-final-video",
    data=json.dumps({"video_id": video_id, "video_filename": video_filename}).encode(),
    headers={"Content-Type": "application/json"},
    method="POST"
)
try:
    r = urllib.request.urlopen(req, timeout=300)
    result = json.loads(r.read())
    print("Final video:", json.dumps(result, indent=2))
except urllib.error.HTTPError as e:
    print("Error:", e.code, json.loads(e.read()))
