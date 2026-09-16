# Repository Guidelines

## Project Structure & Module Organization
This repository is a small FastAPI app for generating Word documents from templates.

- `main.py`: application entrypoint, API routes, and background generation jobs.
- `run_usdoc.py`: local launcher for `uvicorn`.
- `templates/`: Jinja2 HTML templates, including `index.html`.
- `static/`: CSS and browser assets.
- `uploads/`, `outputs/`: runtime files created by the app.
- `build/`, `dist/`: packaging artifacts from PyInstaller.
- `US.docx`, `US_tmp_read.docx`: sample/template documents.
- `generation_records.json`: persisted job status for background generation.

## Build, Test, and Development Commands
Use the virtual environment in `.venv` if present.

- `pip install -r requirements.txt`: install runtime dependencies.
- `python run_usdoc.py`: start the app locally on `http://localhost:8010`.
- `uvicorn main:app --reload --port 8010`: run the API with auto-reload during development.
- `pyinstaller usdoc.spec`: build the packaged Windows executable.

There is no dedicated automated test suite in the repository yet.

## Coding Style & Naming Conventions
The codebase follows standard Python style with 4-space indentation and clear, descriptive function names. Prefer `snake_case` for functions, variables, and filenames, and keep route handlers small and explicit. There is no formatter or linter config checked in, so match the surrounding style in `main.py`.

When adding files, use lowercase names with underscores when needed, for example `generation_records.json` or `run_usdoc.py`. Keep generated outputs out of source directories.

## Testing Guidelines
No formal test framework is configured. If you add tests, place them in a `tests/` directory and name them `test_*.py`. Focus on template validation, filename sanitization, and document generation paths.

For manual verification, launch the app and check that uploads, generated `.docx` files, Excel batch imports, and `generation_records.json` are created correctly. Also confirm invalid template placeholders are rejected.

## Commit & Pull Request Guidelines
Git history uses short, focused commits, often starting with `feat:` or `chore:`. Keep commit messages imperative and scoped to one change.

Pull requests should include a short summary, any setup or verification steps, and screenshots for UI changes in `templates/index.html` or `static/style.css`. Mention any packaging impact if `usdoc.spec` or build outputs change.

## Security & Configuration Tips
Do not commit secrets such as API keys or generated documents. Runtime files belong in `uploads/` and `outputs/`, and local logs like `server.out.log` should stay untracked. The app also honors `PORT` when launched through `run_usdoc.py`. When changing document generation, verify filename sanitization and template validation still reject unsafe input.
