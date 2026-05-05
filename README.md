# BRollAI — Smart Video Explainer POC

Convert talking-head videos into professional explainer videos with AI-generated infographic B-roll.

## Tech Stack (All Free)

| Component | Tool |
|-----------|------|
| AI / LLM | Groq API (free tier) — llama-3.3-70b |
| Backend | Python + Flask |
| PPT Generation | python-pptx |
| Slide → Image | LibreOffice headless |
| Image → Video | ffmpeg |
| Video Merging | ffmpeg |
| Frontend | Vanilla HTML/CSS/JS |

## Prerequisites

1. **Python 3.10+**
2. **ffmpeg** — `sudo apt install ffmpeg` or `brew install ffmpeg`
3. **LibreOffice** — `sudo apt install libreoffice` or `brew install libreoffice`
4. **Groq API Key** — Free at https://console.groq.com

## Installation

```bash
# Clone / extract the project
cd broll_app

# Install Python dependencies
pip install -r requirements.txt

# runing
python app.py
```

## Running

```bash
./start.sh
# OR
python app.py
```

Open http://localhost:5000 in your browser.

## Pipeline Steps

### Step 1 — Upload Video
Upload any MP4/MOV/AVI video. The app supports full-length videos; the B-roll will only be inserted at selected moments.

### Step 2 — Timestamped Transcript
Paste a JSON transcript with segment timestamps. Use the **Load Sample Transcript** button to test immediately without a real video.

Format:
```json
{
  "video_id": "my_video",
  "segments": [
    {"id": "seg_001", "start": 2.0, "end": 8.0, "text": "..."},
    {"id": "seg_002", "start": 8.5, "end": 16.0, "text": "..."}
  ]
}
```

### Step 3 — AI B-Roll Planner
Groq LLM (Llama 3.3 70B) reads the transcript and identifies 1–3 moments where visual explanation adds value. It selects a template and generates slide content.

### Step 4 — Generate PPT Slides
python-pptx creates branded infographic slides from the B-roll plan.

Templates supported:
- **4-Point List** — for reasons, benefits, steps
- **Decision Tree** — for yes/no decisions
- **Problem → Solution** — for problem/solution explanations

### Step 5 — Generate B-Roll Clips
LibreOffice converts each PPTX to PNG, then ffmpeg renders a 4–10 second MP4 clip.

### Step 6 — Merge Final Video
ffmpeg cuts the source video at each B-roll timestamp, inserts the infographic clips, and concatenates everything into a final MP4.

## Project Structure

```
broll_app/
├── app.py              # Flask backend (all logic)
├── requirements.txt
├── start.sh
├── static/
│   └── index.html      # Full frontend
├── input/              # Uploaded videos + transcripts
└── output/             # Generated slides, clips, final video
```

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | /api/upload-video | Upload source video |
| POST | /api/upload-transcript | Submit transcript JSON |
| POST | /api/generate-broll-plan | AI B-roll selection via Groq |
| POST | /api/generate-ppt | Create PPTX slides |
| POST | /api/generate-broll-clips | PPTX → PNG → MP4 |
| POST | /api/generate-final-video | Merge everything |
| GET  | /api/output/<filename> | Download any output file |
| GET  | /api/outputs | List all output files |

## Notes

- For demo purposes, you can use the sample transcript without uploading a real video (steps 1–5 still work, step 6 requires a real video)
- LibreOffice headless conversion may take 5–15 seconds per slide
- ffmpeg merging time depends on video length
- The Groq free tier allows ~14,400 requests/day which is more than enough

## Output Files

After running the full pipeline:
- `broll_plan_<id>.json` — AI-selected B-roll moments
- `broll_001.pptx` — Generated slide (PPTX)
- `broll_001.png` — Slide rendered as image
- `broll_001.mp4` — 4–10 second video clip
- `final_video_<id>.mp4` — Complete merged video
