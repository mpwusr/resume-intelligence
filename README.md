# resume-intelligence

Resume / CV → Excel pipeline. Reads a directory of PDF / DOCX / TXT resumes and
emits a single `.xlsx` workbook with one analysed row per candidate, modelled
after the `CM_Parser` schema.

## Output schema

A single sheet (`Foglio1`) with thirteen columns, one row per candidate:

| # | Column | Type | Source |
|---|---|---|---|
| 1 | Analysis Date | datetime | extraction run timestamp |
| 2 | CV Link | `=HYPERLINK(...)` formula | local `file://` URI to the source resume |
| 3 | First Name | text | model |
| 4 | Last Name | text | model |
| 5 | Current Role | text | most recent job title |
| 6 | Current Employer | text | most recent employer |
| 7 | Prior Employers | comma-joined list | most recent first |
| 8 | Years Experience | text (`"X+ years"` / `"X years"`) | model |
| 9 | Cloud Skills | comma-joined or `"N/A"` | constrained to `{AWS, Azure, GCP, OCI, IBM Cloud, Alibaba Cloud}` |
| 10 | Strategic Summary | English prose, 3–4 sentences | model-generated narrative |
| 11 | Current Tech | comma-joined list | tools/platforms in current role |
| 12 | Historical Tech Map | `EMPLOYER: tech, tech; EMPLOYER: tech, tech` | flattened from a `dict[employer → list[str]]` |
| 13 | System ID | 16-char hex | first 16 chars of SHA-256 of the source file |

## How it works

1. **Ingest** — `pdfplumber` for PDFs, `python-docx` for DOCX, `pathlib.read_text`
   for TXT. Each file becomes a single text blob.
2. **Extract** — one Claude API call per resume, using
   [tool use](https://docs.anthropic.com/en/docs/build-with-claude/tool-use)
   forced via `tool_choice={"type": "tool", "name": "extract_candidate"}`. The
   tool's `input_schema` is the `Candidate` Pydantic model's
   `model_json_schema()`, so the model emits a JSON object that exactly matches
   the schema.
3. **Validate** — `Candidate.model_validate(tool_use.input)`. On
   `ValidationError`, the script retries once with the validation error sent
   back as a `tool_result` correction message.
4. **Write** — `openpyxl` builds the workbook in memory and saves it. The
   `CV Link` column is written as a `=HYPERLINK(...)` formula so the cell is
   clickable in Excel / Numbers / Google Sheets.

Model: `claude-sonnet-4-6`, `max_tokens=8192`. Strategic summary is generated
in English regardless of the resume's source language.

## Pydantic model

```python
class Candidate(BaseModel):
    first_name: str
    last_name: str
    current_role: str
    current_employer: str
    prior_employers: list[str] = []
    years_experience: str            # "12+ years"
    cloud_skills: list[str] = []     # subset of CLOUD_VOCAB
    strategic_summary: str           # 3–4 sentence English narrative
    current_tech: list[str] = []
    historical_tech: dict[str, list[str]] = {}
```

## Project layout

```
.
├── README.md
├── extract_resumes.py    # ingestion + extraction + workbook writer (single script)
├── requirements.txt
├── .env.example          # template — copy to .env and fill in
├── .gitignore            # excludes resumes/, output/, .env, .venv/
├── resumes/              # input directory (gitignored, contains PII)
└── output/               # generated xlsx (gitignored)
```

## Setup

Requires Python 3.11+.

```bash
git clone https://github.com/mpwusr/autoscale-resume-intelligence.git
cd autoscale-resume-intelligence

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cp .env.example .env
# edit .env and paste your real ANTHROPIC_API_KEY
```

`.env` is gitignored; verify with `git check-ignore -v .env`.

## Usage

Drop resumes (PDF / DOCX / TXT) into `resumes/`, then:

```bash
python extract_resumes.py --input resumes/ --output output/CV_Analysis.xlsx
```

Per-file status is printed to stdout (`[ok]`, `[skip]`, `[fail]`) and the final
workbook is written to `--output`. Existing output files are overwritten.

### `.doc` (legacy Word format)

`python-docx` does not read the old `.doc` binary format. Convert first:

```bash
textutil -convert docx resumes/SomeOne.doc   # built into macOS
rm resumes/SomeOne.doc
```

### Scanned PDFs

PDFs without a text layer extract to an empty string and are skipped with a
`[skip] empty extraction` message. OCR is out of scope; pre-process those CVs
with a tool like `ocrmypdf` if needed.

## Configuration

`.env` keys:

| Key | Purpose |
|---|---|
| `ANTHROPIC_API_KEY` | required — your Anthropic console key |

`load_dotenv(override=True)` is called at import, so values in `.env` always
win over any pre-existing shell exports. Useful when an old / placeholder key
is still in your shell environment.

## Cost

Roughly per resume at `claude-sonnet-4-6` rates:

- Input: ~8–15K tokens (a typical resume PDF after text extraction)
- Output: ~500–1500 tokens (the structured tool call)

Budget ~ \$0.05 per CV. A 100-CV batch is ≈ \$5.

## Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| `command not found: python` | macOS provides `python3`, not `python` | use `python3`, or activate the venv (then plain `python` works) |
| `ModuleNotFoundError: No module named 'pdfplumber'` | the `python3` you ran is not the interpreter `pip` installed into | activate the venv (`source .venv/bin/activate`), or run `python3 -m pip install -r requirements.txt` with the same interpreter |
| `error: externally-managed-environment` | Homebrew/PEP 668 blocks system-wide pip installs | use a venv (recommended) |
| `ANTHROPIC_API_KEY is not set` | `.env` not loaded *or* an empty/stale value is set in the shell | the script uses `load_dotenv(override=True)` to defeat shell shadowing — verify `.env` exists in the repo root and contains a non-empty value |
| `401 invalid x-api-key` | placeholder or wrong key in `.env` | replace with a real key from <https://console.anthropic.com/settings/keys> |
| `400 Your credit balance is too low` | account out of API credits | top up at <https://console.anthropic.com/settings/billing> |

## Limitations / not-yet-built

- **Synchronous** — one resume at a time. For batches > ~50, add async with an
  `asyncio.Semaphore`.
- **No manifest / dedup** — every run re-extracts every file. SHA-256 hashes
  are written as `System ID` but not consulted as a cache.
- **No OCR fallback** for scanned PDFs.
- **`CV Link` points at local `file://` URIs.** If the workbook is shared,
  swap in Google Drive HYPERLINKs (e.g. via a `--drive-map` CSV).
- **Strategic summary tone** is loosely specified — the prompt says "no fluff,
  no first-person", but you should eyeball a few outputs and tighten if needed.

## License

[The Unlicense](LICENSE) — public domain. No rights reserved.
