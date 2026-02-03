# Gemini Developer Guidelines

## 1. Session Initialization
- **CRITICAL**: At the start of every session, **read `dev_log.md`** to understand the latest project status, recent refactoring, and outstanding tasks.
- Check `requirements.txt` to understand the dependency tree.

## 2. Development Standards
- **Code Style**:
  - Follow PEP 8 guidelines for Python code.
  - **Type Hinting**: enforce strict type hints (`typing` module) for all new functions and class methods.
- **Security**:
  - **NEVER** commit `service_key.json` or any other secret keys.
  - Always verify `.gitignore` protects sensitive files before adding to git.

## 3. Workflow & Verification
- **Testing**:
  - Always dry-run scripts after modification.
  - For `mokil_high_school_results_gen.py`, use piped input (e.g., `echo "" | python ...`) to bypass interactive prompts during automated testing.
- **Git Operations**:
  - Run `git status` frequently to track file states.
  - Use semantic commit messages (e.g., `Refactor: ...`, `Docs: ...`, `Feat: ...`).
  - Push to `main` only after local verification passes.

## 4. Project Structure
- `generators/`: Contains core logic scripts.
- `reports/`: Output directory for generated HTML/Excel files.
- `service_key.json`: Required for `generate_dashboard.py` (must be present locally but ignored by git).
