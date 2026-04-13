## Job tracker (Gmail → Excel + reminders)

### What it does
- Scans your Gmail (from `2026-01-01` onward, configurable in `config.json`)
- Exports an Excel tracker to `output/applications.xlsx`
- Creates calendar reminders for **Action needed** in `output/reminders.ics`
- Optional: **local AI parsing** using **Ollama** (free, runs on your Mac)

### One-time setup
1. Install dependencies

```bash
cd /Users/anand/Desktop/job-tracker
python3.13 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

2. Gmail OAuth
- Put `credentials.json` in this folder (same level as `sync.py`)
- Run once; browser opens; sign in with `@umd.edu`

### Run

```bash
cd /Users/anand/Desktop/job-tracker
source .venv/bin/activate
python sync.py --source gmail --out output/applications.xlsx --ics output/reminders.ics
```

### Local AI (recommended)
This reduces rule brittleness and improves extraction (company, role, action type).

1. Install Ollama: `https://ollama.com`
2. Start it (it runs a local server at `http://localhost:11434`)
3. Pull a model (example):

```bash
ollama pull qwen2.5:7b-instruct
```

4. Run normally (AI is the default). If Ollama isn’t running, the script falls back to non-AI mode.

To disable AI explicitly:

```bash
python sync.py --ai none --source gmail --out output/applications.xlsx --ics output/reminders.ics
```

### Notes
- AI results are cached in `output/ai-cache.json` to avoid re-parsing the same email.
- Only **Action needed** rows generate reminders in the `.ics`.

