# Contributing

Thanks for improving `summarize-ppt-notes`.

## Local Setup

```bash
python3 -m venv .venv
. .venv/bin/activate
python -m pip install --upgrade pip setuptools
python -m pip install -e ".[dev,render]"
python scripts/doctor.py
python -m pytest
```

## Pull Request Standard

- Keep the Codex Skill valid: `SKILL.md` must keep only `name` and `description` in frontmatter.
- Add or update tests for extraction, quality reports, or DOCX generation changes.
- Do not commit real courseware, proprietary slides, generated work directories, or personal documents.
- Prefer deterministic local extraction. Model calls should stay outside the core scripts unless they are optional and clearly documented.
