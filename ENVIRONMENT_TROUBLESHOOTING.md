# Environment Troubleshooting

This project now assumes a workspace-local Python virtual environment at `.venv/`, especially for Windows launcher scripts and VS Code settings.

## Expected layout

```text
twse-foreign-daily-report/
  .venv/
    Scripts/
      python.exe
      Activate.ps1
  .vscode/
    settings.json
    launch.json
    profile.ps1
  run_latest.bat
  run_with_date.bat
```

## What changed after the last commit

- Windows batch launchers now call the interpreter from `.venv\Scripts\python.exe` instead of relying on `python` from `PATH`.
- VS Code launch settings now use `${workspaceFolder}`-relative interpreter paths, so the repo can be moved without breaking debugging.
- VS Code terminal settings now load `.vscode/profile.ps1`, which activates the project virtual environment on startup.

## If the batch files fail

Recreate the virtual environment and reinstall dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Then confirm the interpreter exists:

```powershell
Test-Path .\.venv\Scripts\python.exe
```

## If VS Code still uses the wrong interpreter

- Reopen the workspace so `${workspaceFolder}` is resolved again.
- Run `Python: Select Interpreter` and choose `.venv\Scripts\python.exe` if VS Code cached an older path.
- Open a new integrated terminal so `Project PowerShell` reloads `.vscode/profile.ps1`.

## Why this setup is safer

- It avoids accidental execution against a global Python installation.
- It removes machine-specific absolute paths from the shared workspace settings.
- It makes launcher behavior consistent between command line and VS Code.
