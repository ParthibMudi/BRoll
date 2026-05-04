import urllib.request, json

base = "http://localhost:5000"
video_id = "5127cde4"

# Test 1: health
r = urllib.request.urlopen(f"{base}/api/health")
print("Health:", json.loads(r.read()))

# Test 2: generate-ppt (reuse existing plan)
req = urllib.request.Request(
    f"{base}/api/generate-ppt",
    data=json.dumps({"video_id": video_id}).encode(),
    headers={"Content-Type": "application/json"},
    method="POST"
)
r = urllib.request.urlopen(req)
print("PPT:", json.loads(r.read()))

# Test 3: generate-broll-clips
req = urllib.request.Request(
    f"{base}/api/generate-broll-clips",
    data=json.dumps({"video_id": video_id}).encode(),
    headers={"Content-Type": "application/json"},
    method="POST"
)
r = urllib.request.urlopen(req)
result = json.loads(r.read())
print("Clips:", json.dumps(result, indent=2))
