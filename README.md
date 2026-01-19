# course-kit

Utilities and templates used to build the J4734 course materials (calendar, PPTX
slides, and templates).

## Quickstart

- Prerequisites: Python 3.8+ (3.10 recommended), pip.
- Install minimal dependencies:

```bash
pip install openpyxl python-pptx
```

- Export the semester calendar (example):

```bash
python3 scripts/import_calendar_xls.py \
  --xls "/path/to/sp2026-dates.csv" \
  --populate-from-db \
  --semester SP2026 \
  --populate-base "/path/to/4734-master.xltx" \
  --populate-output "/path/to/SP2026_calendar.xlsx"
```

- Many other helper scripts exist in the repo; see the `scripts/` and top-level
  Python files for available commands.

## Templates & Assets

- The repository contains templates such as `4734-master.xltx` (template source).
- Large or local-only assets are intentionally ignored. The `assets/` directory
  is present locally but is excluded from the repository to keep the history
  small. If you need to publish large binary assets, use Git LFS or external
  hosting.

## Secrets / Configuration

- A redacted example config is provided: `canvas_config.example.json` — copy it
  to `canvas_config.json` and fill in your real values when running tools that
  need Canvas access.

  **Important:** `canvas_config.json` is gitignored. **Do not** commit real API
  keys. If you previously committed any keys, rotate/revoke them immediately.

Example `canvas_config.example.json` contents:

```json
{
  "canvas_url": "https://your-institution.instructure.com",
  "api_key": "REPLACE_WITH_YOUR_CANVAS_API_KEY",
  "course_id": "COURSE_ID"
}
```

## Repo cleanup & history note

- Local-only output files (PPTX, large images) were removed from the index and
  added to `.gitignore` to keep the repo small.
- A backup branch named `backup-before-filter` was created before the history
  rewrite. If you need to recover any removed binaries, check that branch (do
  not push it publicly if it contains large or sensitive files).

## Tests

- Some tests exist in the repo; run them with `pytest` if you want to validate
  behavior locally.

```bash
pytest -q
```

## Contributing

- Create feature branches off `main` and open pull requests against `main`.
- Keep binary assets out of the repo; prefer Git LFS or external storage.

If you want, I can add a short `requirements.txt` and a developer `Makefile` to
simplify common tasks.
