# autoscale-resume-intelligence

Resume → CSV extraction pipeline. Parses PDF/DOCX/TXT resumes and emits one row per bullet point with employer, role, technologies, and dates.

## Output Schema

One row per bullet:

| Column | Description |
|---|---|
| `employee_name` | Candidate name |
| `employer` | Company name |
| `role` | Job title |
| `job_start` | Job start date (`YYYY-MM` or `YYYY`) |
| `job_end` | Job end date (null = Present) |
| `bullet_text` | Raw bullet text |
| `technologies` | Semicolon-joined tech list |
| `project_start` | Bullet/project start (falls back to `job_start`) |
| `project_end` | Bullet/project end (falls back to `job_end`) |
| `source_file` | Original filename |

## Stack

- Python 3.11+
- `pdfplumber` (PDFs), `python-docx` (Word), stdlib for `.txt`
- `anthropic` SDK with tool use for structured output
- `pydantic` v2 for schema validation
- `pandas` for CSV output
- `asyncio` + `anthropic.AsyncAnthropic` for parallelism

## Project Structure

```
resume_extractor/
├── main.py              # CLI entrypoint
├── ingest.py            # file → raw text
├── models.py            # Pydantic schemas
├── extract.py           # Claude API + tool-use call
├── flatten.py           # nested model → CSV rows
├── prompts/
│   └── extraction.md    # system prompt
├── resumes/             # input dir
└── output/
    ├── resumes.csv
    └── manifest.json    # processed-file hashes
```

## Pydantic Models (`models.py`)

```python
from pydantic import BaseModel

class Bullet(BaseModel):
    text: str
    technologies: list[str] = []
    project_start: str | None = None  # "YYYY-MM" or "YYYY"
    project_end: str | None = None

class Job(BaseModel):
    employer: str
    role: str
    start_date: str | None
    end_date: str | None  # null = Present
    bullets: list[Bullet]

class Resume(BaseModel):
    employee_name: str
    jobs: list[Job]
```

## Extraction Flow

1. Register `Resume` as an Anthropic tool via `model_json_schema()`.
2. Force tool use with `tool_choice={"type": "tool", "name": "extract_resume"}`.
3. Claude returns a `tool_use` block → validate into `Resume`.
4. On `ValidationError`, retry once with the error fed back as a correction message.

Model: `claude-sonnet-4-5` is the sweet spot for cost/accuracy. Use Opus for unusually messy resumes.

## Flatten Step

For each resume → job → bullet, emit one row. Fill `project_start` / `project_end` from the bullet if present, else fall back to the parent job's dates. Join `technologies` with `;`.

## Main Loop

1. Scan `resumes/` for new/changed files (SHA-256 vs `manifest.json`).
2. Ingest in parallel (thread pool — I/O bound).
3. Extract in parallel (asyncio — API bound, semaphore at ~5 concurrent).
4. Flatten all rows.
5. Append to `resumes.csv`, update manifest.
6. Log per-file status + any validation failures to `output/errors.log`.

## CLI

```bash
python main.py --input resumes/ --output output/resumes.csv [--force]
```

## Edge Cases

- **Scanned PDFs** (no text layer) → detect empty extraction, log for OCR follow-up
- **"Present" / "Current" / "Now"** → null `end_date`
- **Date ranges like "2018-2020"** on a bullet → split into start/end
- **Bullets with no tech mentions** → empty list, not null
- **Multi-column PDFs** → `pdfplumber` with layout mode or fall back to raw text

## Build Order

| Step | Component | Est. Time |
|---|---|---|
| 1 | `models.py` + `prompts/extraction.md` | 45 min |
| 2 | `ingest.py` (PDF, DOCX, TXT) | 1 hr |
| 3 | `extract.py` single-file happy path | 1 hr |
| 4 | `flatten.py` + CSV writer | 30 min |
| 5 | `main.py` wiring + manifest + async | 1.5 hrs |
| 6 | Test on 5–10 real resumes, tune prompt | ongoing |

**Total: ~5 hours to a working v1.**

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install anthropic pydantic pdfplumber python-docx pandas
export ANTHROPIC_API_KEY=sk-ant-...
```

## License

Proprietary — AutoScaleWorks.ai
