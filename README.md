# Script Debugger

A GUI debugger for Windows batch files (`.bat`, `.cmd`) and PowerShell scripts (`.ps1`).
Open a script, set breakpoints and capture points, then run or step through it while
watching every variable. While paused you can edit the script and apply the change to the
running script, and save any variable's complete value to a file. Code that is not at a
breakpoint or capture point runs at full speed.

## Requirements

- Windows 10/11
- Python 3.10+ with tkinter (the standard python.org / Microsoft Store install includes it)
- Windows PowerShell 5.1 (built in) or PowerShell 7 (`pwsh`) for `.ps1` scripts

No third-party packages are needed.

## Running

Double-click `Script Debugger.cmd`, or from a terminal:

```
python script_debugger.py
python script_debugger.py samples\demo.cmd
```

## Using it

1. **Open** a script (toolbar button, `Ctrl+O`, or pass the path on the command line).
   The script is only loaded. Nothing executes until you press Run or Step.
2. **Set breakpoints** by clicking the gutter to the left of a line, or putting the cursor on
   a line and pressing `F9` (red dot: execution pauses there). **Capture points** are set with
   Ctrl+click in the gutter or `Ctrl+F9` (blue diamond): the variables are recorded there
   without pausing, so you can watch a loop without stopping it. Lines that can never be
   stopped on (blank lines, labels, comments, block closers) are shown in grey; a point placed
   there moves to the next executable line. Both kinds are remembered per file between
   sessions and follow their lines when you edit the text.
3. **Run** (`F5`) executes until the first breakpoint. **Step Into** (`F11`) starts the script
   and pauses before its first line.
4. While paused:
   - The yellow line is the one that will execute next.
   - **Variables** shows every variable as it exists right before that line runs.
     Red rows changed on the previous line, green rows are newly created. Double-click a row
     to see a long value in full. The filter box narrows by name.
   - **Save value...** (or right-click a variable) writes the variable's *complete* value to a
     file. PowerShell values can be saved as text (`.txt`, console-style formatting), JSON
     (`.json`, up to 10 levels deep) or raw bytes (`.bin`, for `byte[]` values). The value is
     fetched live from the paused script, so it is never truncated. After the script has
     finished, the captured value is saved instead (PowerShell previews are cut at 4000
     characters, and the debugger warns you when that happens).
   - **Variable history** logs every change seen between consecutive captures (pauses and
     capture points), in execution order. The *Lines* column shows the range the change
     happened in, e.g. `9→20` means "somewhere after line 9 ran, before line 20". Click a row
     to jump to that line.
   - **Track vars** (toolbar) selects the tracking mode. *Pauses + capture points* is the
     default and costs nothing on other lines. *Every line* records a value snapshot before
     every line, which gives an exact line for every change but makes long loops slow
     (about 4 ms per line for batch, 10 ms or more per line for PowerShell).
   - `F5` Continue, `F11` Step Into, `F10` Step Over, `Shift+F11` Step Out (PowerShell only),
     `Shift+F5` Stop. Right-click a line for *Run to this line*.
5. **Editing while paused.** The editor is live whenever the script is loaded or paused.
   Change, insert or delete lines, or right-click to comment/uncomment a selection. Then press
   **Apply edits** (`Ctrl+Enter`), or just Continue / Step: pending edits are applied to the
   running script before it resumes. The title bar shows `[edits not yet applied]` until then.
   `Ctrl+S` saves the editor contents back to the original file (the title shows `*` while the
   editor differs from disk). Running a script always uses the editor contents, saved or not.
6. **Pause** breaks a free-running script. Batch scripts pause at the next line. PowerShell
   scripts pause at the next breakpoint, capture point or top-level statement; a loop that
   contains none of those cannot be interrupted (use Stop), because interrupting anywhere
   would require a breakpoint on every line and that is exactly what makes scripts slow.
7. Script output appears in the dark console pane. Scripts that read from standard input
   (`pause`, `set /p`, `Read-Host`) receive what you type into *Send to script stdin*.
8. **Arguments** are passed to the script exactly as typed, e.g. `-Name Bob` or `C:\logs 5`.
9. **Handing off to a native debugger** (Debug > Native debugger settings). Two tools for
   scripts that load DLLs or spawn processes:
   - *Attach before the first line*: Run starts the script process, launches your debugger
     command against its PID (presets for WinDbg, x64dbg, dnSpy, or a custom command with
     `{pid}`), and waits until you press Continue. With WinDbg the preset enables
     `.childdbg 1` and `sxe cpr`, so every child process the script creates, including one
     created suspended for hollowing, is debugged and breaks on creation. *Attach debugger*
     on the toolbar launches the same command at any later pause.
   - *Pause on assembly load* (PowerShell): the script pauses the instant it loads a .NET
     assembly, before any of that assembly's code runs, including assemblies loaded from a
     byte array that never touch disk (which a native debugger's module-load events cannot
     see). The status bar names the assembly and the PID; attach dnSpy or WinDbg with SOS
     there. Assemblies from `$PSHOME`, the Windows folder and the GAC are ignored. A load
     performed by a background thread cannot pause instantly: it is reported and the script
     pauses at its next statement. `samplessmload.ps1` demonstrates both.

## How it works

**Batch files** have no native debugger, so the program writes an instrumented copy next to
the original (`~dbg_<name>.cmd`, deleted when the run ends). A `call :__dbg__ N` is inserted
before every executable line; that subroutine dumps the environment (`set`), the current
`ERRORLEVEL` and directory, prints a marker, and waits on stdin until the debugger lets it
continue. `ERRORLEVEL` is preserved across the call. The copy lives beside the original so
`%~dp0` and relative paths behave the same as a normal run.

cmd.exe reads a batch file by byte offset as it executes. Live editing uses that: while the
script is paused the debugger rewrites the instrumented copy, keeping the byte position cmd
will resume at exactly where it was (a padding comment before the body absorbs the size
difference). Edits apply at the next line, including lines revisited by `goto` loops.
Text inserted *above* the paused line can grow by up to 4 KB per run.

**PowerShell scripts** use PowerShell's own debugger. A harness parses the script and runs it
as one compiled chunk from a shadow file next to the original (`~dbg_<name>.ps1`, line numbers
identical to the editor) with a tiny `__DbgCheck` call spliced onto the end of every top-level
statement. Line breakpoints with actions do the capturing: the action collects the variables
of every scope from the current function up to the script scope (plus globals the script
created), writes them as JSON, and waits on stdin when it has to pause.

Performance: PowerShell's cost per executed statement grows with the number of breakpoints
bound to the running code, so in run mode the harness keeps breakpoints only on the lines that
need them (breakpoints, capture points, run-to-line target). Stepping, Pause and the
*Every line* tracking mode temporarily put a breakpoint on every line, and Continue removes
them again, even inside a loop that is already running. Value previews are bounded (long
strings, byte arrays and big collections are summarised), so a multi-megabyte blob costs
nothing to display; *Save value...* still writes it in full.

Applying an edit re-parses the edited text (syntax errors are rejected with the line number).
At the next statement boundary the harness exports the script's variables and functions,
unwinds, and starts a new chunk built from the edited text at that statement with everything
re-injected. Scripts with `begin/process/end` blocks or `using` statements run without live
editing.

Toggle *Show inherited environment* / *Show automatic variables* to also see the variables that
were already there before the script started (PATH, `$PSVersionTable`, ...).

   - *Process guard*: the debugger launches the script's process as its own debuggee
     (Windows `DEBUG_PROCESS`), which automatically covers every descendant: children,
     grandchildren, and processes created suspended and resumed later. Each new process can
     be frozen at creation, at its first native instruction, and/or the moment it loads the
     .NET runtime (`clr.dll`/`coreclr.dll`) — independently selectable, since they serve
     different payloads:
     - *At creation* / *at first instruction* catch a native payload before it runs. A pop-up
       offers *Hand off to debugger* (detach, keep it frozen, launch a **native** debugger
       command — WinDbg or x64dbg — against its PID, then *Resume* once your
       breakpoints are set), *Resume*, or *Kill*.
     - *At .NET/CLR load* catches a **managed (.NET) payload** instead, and is the one to use
       for "most PowerShell malware ends up executing .NET payloads". It freezes on the CLR's
       own first-chance "notification" exception (`0x04242420`), which the runtime raises once
       its debugging data structures are valid, specifically so a native debugger knows it is
       safe to hand over to a managed one — the same signal WinDbg's SOS and Visual Studio's
       mixed-mode debugger wait for. That is later than `clr.dll`/`coreclr.dll` simply being
       mapped (dnSpy's own Attach to Process list won't show a process that young) but still
       well before any managed code, including the payload's own assemblies, runs — confirmed
       by tracing every debug event against a real `MSBuild.exe` launched suspended. It still
       fires even when a hollowing loader redirects the thread's entry point to skip the native
       "first instruction" trap entirely, since reaching this point only requires running
       enough of the CLR, not the native loader thunk. The pop-up's hand-off here launches the
       **managed** debugger configured for assembly-load pause / attach-at-start (dnSpy etc.) —
       this is the combination that makes dnSpy useful against an externally hollowed process,
       not just against PowerShell's own `Assembly.Load`.
     Images listed under *Never freeze* (default `conhost.exe`) are only logged. While the
     guard is on, no other debugger can attach to the script process itself, so it is an
     alternative to *Attach before the first line*, not an addition to it. The script process
     sees `IsDebuggerPresent()` as true. `samples\spawn.ps1` demonstrates a native child
     created suspended, resumed, and spawning a grandchild; `samples\spawn_clr.ps1`
     demonstrates the same pattern with a .NET target, for the CLR-load stage.

## Analysing hostile scripts

The debugger runs the script for real, in a normal cmd.exe or powershell.exe, with your
account's rights. Use an isolated analysis VM with no network, and remember that real-time
antivirus will quarantine a malicious DLL long before the debugger sees it. Anti-debugging
checks inside a loaded DLL (`IsDebuggerPresent`, timing) see the native debugger you attach,
not this tool.

## Known limitations

- Batch `FOR` variables (`%%i`) are not environment variables and cannot be listed.
- Batch lines inside parenthesised blocks are stopped on, but `%var%` inside such a block was
  expanded when the block was parsed, exactly as cmd does normally. Edits inside a block that
  cmd has already parsed take effect the next time the block is entered.
- Step Over / Step Out need call-depth information, which batch files do not provide; for batch
  scripts Step Over behaves like Step Into.
- A batch script that `exit`s (without `/b`) ends the cmd process immediately, as it normally would.
- PowerShell edits take effect at the next *top-level* statement. A loop that is currently
  running keeps its old body until it completes. Functions defined before an edit keep their
  original line numbers until the next run. Classes defined by the script do not survive an
  edit restart.
- In the default tracking mode the history only shows what changed between captures; put a
  capture point inside a loop (or switch to *Every line*) when you need per-iteration detail.
- Variable values are captured before each line, so a change made by the very last line of a
  script shows only in the final Variables view, not in the history.
