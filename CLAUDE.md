# The Declared List Generator

## Project Overview
The Declared List Generator (DLG) is a tool for generating a Declared List Document. The tool will read a database that contains selecable items. These items will be presented to the user in table form. The user is then able to select all items that are part of a specific declared list. The selections is a configuration that will be stored in a session or project file that can be loaded into a later session. It is foreseeable that with time a database of these session/project files will be created for re-use. When the user is done selecting items, the tool will have a function for generating a word document containing the declared list as a word table. The document template will be provided as an empty word file


### Project Direction
* **GUI**: The GUI will run in the web browser (Edge, Firefox, Chrome)
* **Language:** Python 3.11+ (use type hints throughout).
* **Dependencies:** Stick strictly to the Python standard library and widely adopted, well-known modules (e.g., `requests`, `pydantic`, `pytest`). Ask before adding niche third-party packages.
* **Code Documentation:** Every module, class, and function must include clear Google-style docstrings. Add concise inline comments explaining *why* complex logic exists, not *what* obvious lines do.


### Planning & State Tracking (Surviving /clear)
* **Source of Truth:** Maintain a `PLAN.md` file in the root directory to track architecture decisions, implementation phases, and completed steps.
* **Plan Before Coding:** Before writing or modifying code for any feature, write out a step-by-step implementation plan in `PLAN.md` and ask for my approval.
* **Update After Every Step:** Immediately after completing and testing a task, check it off in `PLAN.md`, log any key design decisions or gotchas, and note the exact next step.
* **Session Resume:** Whenever a new session starts (or after `/clear`), read `PLAN.md` first before taking any action.




### Documentation Layout
* `PLAN.md` – short (about one screen): status, current release plan with checkboxes, next step, links. Finished releases move to `docs/archive/`.
* `docs/decisions.md` – design decisions and gotchas (newest first). `docs/backlog.md` – ideas and open items. `CHANGELOG.md` – per-release summary. `README.md` – user-facing.
* `my_notes.md` is the user's private, git-ignored scratch file: copy its decisions into the docs above, never commit it.
* Versioning: SemVer, annotated tags `vX.Y.Z` on `main`, version in `pyproject.toml`; session `schema_version` is separate.
* Platform: develop on Linux, must also work on Windows (no POSIX-only paths/APIs).
