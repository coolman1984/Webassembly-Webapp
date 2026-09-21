# Packaging a Python tool so a non-technical user never installs anything

## The bar

Someone with no dev tools, no admin rights, and no idea what Python is, copies the project folder (or receives
it as a zip) to a PC that has never run anything like this before, double-clicks one `.bat` file, and it works -
today and in six months, regardless of what else is or isn't installed on that machine. The only two things
assumed present are Windows itself and desktop Microsoft Excel (which cannot be bundled - it's licensed
software, not a library).

## Bundle a private, trimmed Python runtime inside the project

Rather than the official "embeddable" zip (which needs manual `pip`/`.pth` surgery to become usable with
third-party packages), the reliable approach is to build the bundle from a real, working Python installation
that already has the needed packages, and copy only what's used:

1. Copy the interpreter and its core DLLs (`python.exe`, `pythonw.exe`, `python3XX.dll`, `vcruntime*.dll`) next
   to the project.
2. Copy `Lib/` (the standard library) *excluding* what the tool doesn't use: test suites, `idlelib`, `tkinter`
   (unless a GUI needs it), `turtledemo`, `ensurepip`, `lib2to3`, `pydoc_data`, `__pycache__`.
3. Copy only the specific third-party packages actually imported (e.g. just the `pywin32` pieces: `win32`,
   `win32com`, `win32comext`, `pywin32_system32`, plus its loader `.pth`/bootstrap files) into
   `runtime/Lib/site-packages/` - not the whole of some development environment's `site-packages`.
4. Precompile: `python.exe -I -m compileall -q Lib` so the very first run on the target machine is fast even
   though the folder might be on a slow or read-only share.

A build like this for a pywin32-based tool comes out around 40-50 MB - small enough to check into source
control or ship as a zip, and entirely self-contained.

## The launcher must be immune to the target machine's own Python (or lack of one)

```bat
"runtime\python.exe" -I -X utf8 server.py
```

`-I` (isolated mode) makes Python ignore the user's `PYTHONPATH`, their own `site-packages`, and any
environment-variable configuration entirely - the bundled runtime behaves identically whether the target
machine has no Python at all, a different Python version already on `PATH`, or a broken/half-configured one.
Never invoke a bare `python`/`python3` and hope it resolves correctly; always use the full path to the bundled
interpreter.

## Verify it for real, not by assumption

Run the bundled interpreter with a scrubbed environment - only bare Windows system directories on `PATH`, no
Python-related environment variables at all - and confirm the tool still starts and does everything it needs to
(including COM automation, since `pywin32`'s registration can be sensitive to environment). This catches "works
on my dev machine because my dev machine's own Python happens to be on PATH too" bugs that are otherwise
invisible until the tool reaches an actual clean target machine.

## A maintainer-only build script, run once, checked in as output

Building the trimmed runtime is a one-time (or occasional) maintainer task, not something the end user or even
every future contributor needs to redo - write it as a script (e.g. a PowerShell script using `robocopy` for the
selective copies) so it's reproducible and re-runnable when a dependency changes, but ship the *built* `runtime/`
folder itself as part of what the end user receives.

## Startup latency matters more than you'd think

A service meant to be double-clicked and immediately used should report itself ready in a couple of seconds, not
tens of seconds. If a launcher self-checks "is an instance already running?" by pinging a fixed port range, keep
that probe narrow (the last-used port plus one default, not a wide scan) - probing many closed ports one by one
can itself take several seconds on a machine running local security/endpoint software that intercepts loopback
connection attempts.
