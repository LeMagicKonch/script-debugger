#!/usr/bin/env python3
"""Script Debugger - a GUI debugger for Windows batch (.bat/.cmd) and PowerShell (.ps1) scripts.

Load a script, place breakpoints in the gutter, then Run / Step. The script does not
execute until you press Run or Step. Variables are captured at every pause and at
"capture points" (or on every line in the slow tracking mode), so assignments are logged
in the Variable History pane. While paused you can edit the script in place and apply the
change to the running script, and save any variable's full value to a file.
"""
import ctypes
import difflib
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Script Debugger"
STOP_MARKER = "__DBG_STOP__"
TRACE_MARKER = "__DBG_TRACE__"
MSG_MARKER = "__DBG_MSG__"
BATCH_EXTS = (".bat", ".cmd")
PS_EXTS = (".ps1",)
CREATE_NO_WINDOW = 0x08000000
TRUNCATED_MARK = "[truncated"
TRACK_POINTS = "points"
TRACK_ALL = "all"


class EditError(Exception):
    """Raised when an edit cannot be applied to the running script."""


def oem_encoding():
    try:
        return "cp%d" % ctypes.windll.kernel32.GetOEMCP()
    except Exception:
        return "cp437"


def line_mapping(old_lines, new_lines):
    """Map 1-based line numbers of old_lines to new_lines.

    Returns (start_map, end_map). start_map[n] is the new line that corresponds to old
    line n (for a deleted/replaced line: the first line of the replacement, i.e. where
    execution would land). end_map[n] is the last new line that still belongs to the
    region that contained old line n (used for 'resume after line n' positions).
    """
    sm = difflib.SequenceMatcher(None, old_lines, new_lines, autojunk=False)
    start_map, end_map = {}, {}
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                start_map[i1 + k + 1] = j1 + k + 1
                end_map[i1 + k + 1] = j1 + k + 1
        else:
            for k in range(i1, i2):
                start_map[k + 1] = j1 + 1
                end_map[k + 1] = j2 if j2 > j1 else j1
    return start_map, end_map


def write_file_atomic(path, data):
    """Replace `path` with `data` (bytes); retried because a reader may hold it open."""
    tmp = path + ".tmp"
    with open(tmp, "wb") as fh:
        fh.write(data)
    for _ in range(50):
        try:
            os.replace(tmp, path)
            return
        except PermissionError:
            time.sleep(0.005)
    with open(path, "wb") as fh:
        fh.write(data)
    try:
        os.remove(tmp)
    except OSError:
        pass


# ----------------------------------------------------------------------------
# Batch instrumentation
# ----------------------------------------------------------------------------

# The pause subroutine sits at the TOP of the instrumented file so that its byte offsets
# never move when the body is rewritten while the script is paused inside it.
# control.cmd (rewritten by the GUI) defines __DBG_BPS, __DBG_CPS, __DBG_MODE, __DBG_TRACK.
# One redirect per capture: appending to an existing file is ~10x slower than creating one.
_CAPTURE = b'@(set&echo __DBG_ERRORLEVEL=%__dbg_el%&echo __DBG_CD=&cd)>"%__DBG_DIR%\\vars_%__dbg_seq%.txt"\r\n'
BATCH_HEAD = (
    b"@goto :__dbg_main\r\n"
    b":__dbg__\r\n"
    b'@set "__dbg_el=%errorlevel%"\r\n'
    b'@for /f "usebackq delims=" %%L in ("%__DBG_DIR%\\control.txt") do @set "%%L"\r\n'
    b"@set /a __dbg_seq+=1 >nul\r\n"
    b'@call set "__dbg_t=%%__DBG_BPS:,%1,=%%"\r\n'
    b'@if not "%__dbg_t%"=="%__DBG_BPS%" goto :__dbg_stop\r\n'
    b'@if "%__DBG_MODE%"=="step" goto :__dbg_stop\r\n'
    b'@if "%__DBG_TRACK%"=="all" goto :__dbg_trace\r\n'
    b'@call set "__dbg_t=%%__DBG_CPS:,%1,=%%"\r\n'
    b'@if not "%__dbg_t%"=="%__DBG_CPS%" goto :__dbg_trace\r\n'
    b"@exit /b %__dbg_el%\r\n"
    b":__dbg_trace\r\n"
    + _CAPTURE +
    b"@echo " + TRACE_MARKER.encode() + b" %1 0 0 0 %__dbg_seq%\r\n"
    b"@exit /b %__dbg_el%\r\n"
    b":__dbg_stop\r\n"
    + _CAPTURE +
    b"@echo " + STOP_MARKER.encode() + b" %1 0 0 0 %__dbg_seq%\r\n"
    b'@set /p "__dbg_cmd="\r\n'
    b"@exit /b %__dbg_el%\r\n"
    b":__dbg_main\r\n"
    b'@set >"%__DBG_DIR%\\vars_0.txt"\r\n'
)
BATCH_INITIAL_PAD = 4096   # slack so that edits before the paused line can grow
BATCH_MAX_LINE = 8000


def batch_is_executable(stripped_lower):
    s = stripped_lower
    return not (
        not s
        or s.startswith(b")")
        or s.startswith(b":")
        or s == b"rem"
        or s.startswith(b"rem ")
        or s.startswith(b"rem\t")
    )


def batch_executable_lines(lines_bytes):
    """1-based line numbers the debugger can stop on (list of bytes lines, no newlines)."""
    executable = set()
    continued = False
    for i, text in enumerate(lines_bytes, 1):
        stripped = text.strip().lstrip(b"@").strip().lower()
        if not continued and batch_is_executable(stripped):
            executable.add(i)
        trail = text.rstrip()
        carets = len(trail) - len(trail.rstrip(b"^"))
        continued = carets % 2 == 1
    return executable


def batch_make_pad(n):
    """n bytes of harmless '@rem' lines (n == 0 -> nothing; n must not be 1..5)."""
    out = b""
    while n > 0:
        c = min(n, BATCH_MAX_LINE)
        if n > BATCH_MAX_LINE and 0 < n - BATCH_MAX_LINE < 6:
            c = BATCH_MAX_LINE - 6
        out += b"@rem" + b" " * (c - 6) + b"\r\n"
        n -= c
    return out


def build_batch(lines_bytes, pad_bytes, lead_spaces=None):
    """Build the instrumented file.

    Returns (file_bytes, executable_lines, content_offsets) where content_offsets maps a
    line number to the absolute byte offset of that line's content in the file (the
    position cmd.exe resumes at after the debugger subroutine returns).
    """
    executable = batch_executable_lines(lines_bytes)
    body = b""
    rel = {}
    for i, text in enumerate(lines_bytes, 1):
        if i in executable:
            body += b"@call :__dbg__ %d\r\n" % i
        rel[i] = len(body)
        if lead_spaces and i in lead_spaces:
            body += b" " * lead_spaces[i]
        body += text + b"\r\n"
    rel[len(lines_bytes) + 1] = len(body)      # "end of script" position
    pad = batch_make_pad(pad_bytes)
    base = len(BATCH_HEAD) + len(pad)
    offsets = {i: base + off for i, off in rel.items()}
    return BATCH_HEAD + pad + body, executable, offsets


def parse_set_output(text):
    """Parse the output of cmd's `set` (plus our ERRORLEVEL/CD additions)."""
    result = {}
    cur = None
    expect_cd = False
    for line in text.split("\n"):
        line = line.rstrip("\r")
        if expect_cd:
            result["CD"] = line
            expect_cd = False
            cur = None
            continue
        if line.rstrip() == "__DBG_CD=":
            expect_cd = True
            continue
        m = re.match(r"^([^=]+)=(.*)$", line)
        if m:
            cur = m.group(1)
            result[cur] = m.group(2)
        elif cur is not None and line:
            result[cur] += "\n" + line
    if "__DBG_ERRORLEVEL" in result:
        result["ERRORLEVEL"] = result.pop("__DBG_ERRORLEVEL")
    for k in list(result):
        if k.lower().startswith("__dbg"):
            del result[k]
    return result


def ps_executable_lines(lines):
    """Best-effort set of PowerShell lines a breakpoint can hit."""
    exe = set()
    in_block = False
    for i, line in enumerate(lines, 1):
        s = line.strip()
        if in_block:
            if "#>" in s:
                in_block = False
            continue
        if s.startswith("<#"):
            if "#>" not in s:
                in_block = True
            continue
        if not s or s.startswith("#") or s.startswith("}"):
            continue
        exe.add(i)
    return exe


# ----------------------------------------------------------------------------
# PowerShell harness (written to a temp dir and run with -Command)
# ----------------------------------------------------------------------------

PS_HARNESS = r'''param([string]$__DbgScript, [string]$__DbgState, [string]$__DbgShadowArg)
$__DbgArgString = [string]$env:__DBG_ARGS
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch {}
$global:__DbgHarness = $true
$global:__DbgShadow = $__DbgShadowArg
$global:__DbgAuto = @{}
foreach ($__DbgTmp in (Get-Variable -Scope Global | ForEach-Object Name)) { $global:__DbgAuto[$__DbgTmp] = $true }
foreach ($__DbgTmp in @('_','PSItem','foreach','switch','Matches','input','MyInvocation','PSBoundParameters',
                   'PSCommandPath','PSScriptRoot','LASTEXITCODE','StackTrace','this','PSCmdlet','PSDebugContext')) {
    $global:__DbgAuto[$__DbgTmp] = $true
}
$null = $global:__DbgAuto.Remove('args')
$global:__DbgStateDir = $__DbgState
$global:__DbgSourceFile = Join-Path $__DbgState 'source.ps1'
$global:__DbgControlFile = Join-Path $__DbgState 'control.json'
$global:__DbgStmtEnd = 0
$global:__DbgEnds = @()
$global:__DbgGen = 0
$global:__DbgReloadAfter = -1
$global:__DbgReloadGen = 0
$global:__DbgInjecting = $false
$global:__DbgVars = $null
$global:__DbgFuncs = $null
$global:__DbgTopDepth = -1
$global:__DbgLineMax = 0
$global:__DbgBpObjs = @{}
$global:__DbgSeq = 0
$global:__DbgCtlTime = [DateTime]::MinValue
$global:__DbgCtl = @{ bps = (New-Object 'System.Collections.Generic.HashSet[int]'); cps = (New-Object 'System.Collections.Generic.HashSet[int]');
                     mode = 'run'; depth = 0; temp = -1; track = 'points' }
$global:__DbgSw = [Diagnostics.Stopwatch]::StartNew()

function global:__DbgMsg($text) { [Console]::Out.WriteLine("__DBG_MSG__ $text"); [Console]::Out.Flush() }

function global:__DbgRefreshControl {
    # Breakpoints, capture points and the run mode come from control.json, which the GUI
    # rewrites whenever they change; checked at most every 100 ms.
    # Returns $true when the control file was re-read.
    $global:__DbgSw.Restart()
    try {
        $__m = [IO.File]::GetLastWriteTimeUtc($global:__DbgControlFile)
        if ($__m -eq $global:__DbgCtlTime) { return $false }
        $__j = [IO.File]::ReadAllText($global:__DbgControlFile) | ConvertFrom-Json
        $__bps = New-Object 'System.Collections.Generic.HashSet[int]'
        if ($__j.bps) { foreach ($__b in @($__j.bps)) { $null = $__bps.Add([int]$__b) } }
        $__cps = New-Object 'System.Collections.Generic.HashSet[int]'
        if ($__j.cps) { foreach ($__b in @($__j.cps)) { $null = $__cps.Add([int]$__b) } }
        $global:__DbgCtl = @{ bps = $__bps; cps = $__cps; mode = [string]$__j.mode; depth = [int]$__j.depth;
                             temp = [int]$__j.temp; track = [string]$__j.track }
        $global:__DbgCtlTime = $__m
        return $true
    } catch { return $false }
}

function global:__DbgSyncBreakpoints {
    # PowerShell's per-statement cost grows with the number of breakpoints bound to the
    # running code, so in run mode only the lines that matter carry a breakpoint. Stepping
    # and 'track every line' need a breakpoint on every line. Changes bind to code that is
    # already compiled and running.
    $__ctl = $global:__DbgCtl
    $__want = New-Object 'System.Collections.Generic.HashSet[int]'
    if ($__ctl.mode -ne 'run' -or $__ctl.track -eq 'all') {
        for ($__l = 1; $__l -le $global:__DbgLineMax; $__l++) { $null = $__want.Add($__l) }
    } else {
        foreach ($__l in $__ctl.bps) { if ($__l -ge 1 -and $__l -le $global:__DbgLineMax) { $null = $__want.Add($__l) } }
        foreach ($__l in $__ctl.cps) { if ($__l -ge 1 -and $__l -le $global:__DbgLineMax) { $null = $__want.Add($__l) } }
        if ($__ctl.temp -ge 1 -and $__ctl.temp -le $global:__DbgLineMax) { $null = $__want.Add($__ctl.temp) }
    }
    $__remove = @()
    foreach ($__l in @($global:__DbgBpObjs.Keys)) { if (-not $__want.Contains($__l)) { $__remove += $__l } }
    foreach ($__l in $__remove) {
        try { Remove-PSBreakpoint -Breakpoint $global:__DbgBpObjs[$__l] -ErrorAction SilentlyContinue } catch { }
        $global:__DbgBpObjs.Remove($__l)
    }
    $__add = @()
    foreach ($__l in $__want) { if (-not $global:__DbgBpObjs.ContainsKey($__l)) { $__add += $__l } }
    if ($__add.Count -gt 0) {
        foreach ($__b in @(Set-PSBreakpoint -Script $global:__DbgShadow -Line $__add -Action $global:__DbgAction)) {
            $global:__DbgBpObjs[$__b.Line] = $__b
        }
    }
}

function global:__DbgHex($bytes, $count) {
    $n = [Math]::Min($count, $bytes.Length)
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $n; $i++) { $null = $sb.Append($bytes[$i].ToString('X2')); if ($i -lt $n - 1) { $null = $sb.Append(' ') } }
    if ($bytes.Length -gt $n) { $null = $sb.Append(' ...') }
    return $sb.ToString()
}

function global:__DbgCount($v) {
    try { if ($null -ne $v.Count) { return [int]$v.Count } } catch { }
    try { if ($null -ne $v.Length) { return [int]$v.Length } } catch { }
    return -1
}

function global:__DbgShort($v, $max) {
    # One-level, bounded-cost preview used for elements and properties.
    if ($null -eq $v) { return '$null' }
    if ($v -is [string]) { if ($v.Length -gt $max) { return $v.Substring(0, $max) + "...($($v.Length) chars)" } return $v }
    if ($v -is [ValueType]) { return "$v" }
    if ($v -is [byte[]]) { return "byte[$($v.Length)] " + (__DbgHex $v 16) }
    if ($v -is [System.Collections.IDictionary]) { return "@{...} ($($v.Count) entries)" }
    if ($v -is [System.Collections.IEnumerator]) { return '<enumerator>' }
    if ($v -is [System.Collections.IEnumerable]) { $c = __DbgCount $v; return "$($v.GetType().Name)" + $(if ($c -ge 0) { " ($c items)" } else { '' }) }
    $s = "$v"
    if ($s.Length -gt $max) { $s = $s.Substring(0, $max) + "...($($s.Length) chars)" }
    return $s
}

function global:__DbgFormat($v) {
    # Bounded-cost preview: never serializes a whole large value. Save value to file
    # gives the complete value.
    if ($null -eq $v) { return '$null' }
    # Never stringify enumerators ($foreach, $switch, $input): converting one to a
    # string walks it to the end and silently ends the script's own loop.
    if ($v -is [System.Collections.IEnumerator]) { return '<enumerator>' }
    if ($v -is [string]) {
        if ($v.Length -gt 4000) { return $v.Substring(0, 4000) + " ... [truncated, $($v.Length) chars; use Save value to file]" }
        return $v
    }
    if ($v -is [ValueType]) { return "$v" }
    if ($v -is [scriptblock]) { return '{' + $v.ToString() + '}' }
    if ($v -is [byte[]]) {
        if ($v.Length -gt 64) { return "byte[$($v.Length)] " + (__DbgHex $v 64) + " [truncated; use Save value to file]" }
        return "byte[$($v.Length)] " + (__DbgHex $v 64)
    }
    if ($v -is [System.Collections.IDictionary]) {
        $parts = New-Object System.Collections.Generic.List[string]
        $i = 0
        foreach ($k in $v.Keys) {
            if ($i -ge 50) { $parts.Add("... [truncated, $($v.Count) entries; use Save value to file]"); break }
            $parts.Add("$k = " + (__DbgShort $v[$k] 200)); $i++
        }
        return '@{ ' + ($parts -join '; ') + ' }'
    }
    if ($v -is [System.Collections.IEnumerable]) {
        $parts = New-Object System.Collections.Generic.List[string]
        $i = 0
        $c = __DbgCount $v
        foreach ($item in $v) {
            if ($i -ge 30) { $parts.Add("... [truncated, $c items; use Save value to file]"); break }
            $parts.Add((__DbgShort $item 200)); $i++
        }
        return '@( ' + ($parts -join ', ') + ' )'
    }
    if ($v -is [System.Management.Automation.PSObject] -or $v.PSObject.Properties.Count -gt 0) {
        try {
            $parts = New-Object System.Collections.Generic.List[string]
            $i = 0
            foreach ($p in $v.PSObject.Properties) {
                if ($i -ge 50) { $parts.Add('... [truncated; use Save value to file]'); break }
                $parts.Add("$($p.Name) = " + (__DbgShort $p.Value 200)); $i++
            }
            if ($parts.Count -gt 0) { return '{ ' + ($parts -join '; ') + ' }' }
        } catch { }
    }
    $s = "$v"
    if ($s.Length -gt 4000) { $s = $s.Substring(0, 4000) + " ... [truncated, $($s.Length) chars; use Save value to file]" }
    return $s
}

function global:__DbgWriteValue($v, $path) {
    $ext = [IO.Path]::GetExtension($path).ToLowerInvariant()
    if ($ext -eq '.json') {
        $s = ConvertTo-Json -InputObject $v -Depth 10 -WarningAction SilentlyContinue
        [IO.File]::WriteAllText($path, [string]$s, [Text.Encoding]::UTF8)
        return
    }
    if ($v -is [byte[]]) { [IO.File]::WriteAllBytes($path, $v); return }
    if ($null -eq $v) { $s = '$null' }
    elseif ($v -is [string] -or $v -is [ValueType]) { $s = "$v" }
    elseif ($v -is [scriptblock]) { $s = $v.ToString() }
    elseif ($v -is [System.Collections.IDictionary]) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($k in $v.Keys) { $null = $sb.AppendLine("$k = $($v[$k])") }
        $s = $sb.ToString()
    }
    elseif ($v -is [System.Collections.IEnumerable]) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($item in $v) {
            if ($null -eq $item -or $item -is [string] -or $item -is [ValueType]) { $null = $sb.AppendLine("$item") }
            else { $null = $sb.AppendLine((($item | Format-List * | Out-String -Width 4096) -replace '[ \t]+(\r?\n)', '$1').Trim()); $null = $sb.AppendLine() }
        }
        $s = $sb.ToString()
    }
    else { $s = (($v | Format-List * | Out-String -Width 4096) -replace '[ \t]+(\r?\n)', '$1').Trim() + "`r`n" }
    if ($ext -eq '.bin') { [IO.File]::WriteAllBytes($path, [Text.Encoding]::UTF8.GetBytes([string]$s)) }
    else { [IO.File]::WriteAllText($path, [string]$s, [Text.Encoding]::UTF8) }
}

function global:__DbgSetBreakpoints($upto) {
    if ($upto -gt $global:__DbgLineMax) { $global:__DbgLineMax = $upto }
    $global:__DbgCtlTime = [DateTime]::MinValue
    $null = __DbgRefreshControl
    __DbgSyncBreakpoints
}

# Variables PowerShell creates in every scope; never exported across a live-edit restart.
$global:__DbgPerScope = @{}
foreach ($__DbgTmp in @('?','^','$','_','PSItem','this','input','foreach','switch','Matches','PSCmdlet','PSDebugContext',
                   'MyInvocation','PSBoundParameters','PSCommandPath','PSScriptRoot','true','false','null',
                   'ConsoleFileName','ExecutionContext','HOME','Host','PID','ShellId','PSVersionTable','PSHOME',
                   'PSCulture','PSUICulture','PSEdition','StackTrace','Error','LASTEXITCODE','MaximumAliasCount',
                   'MaximumDriveCount','MaximumErrorCount','MaximumFunctionCount','MaximumVariableCount')) {
    $global:__DbgPerScope[$__DbgTmp] = $true
}

function global:__DbgCheck($k) {
    # Called between top-level statements of the running chunk (inserted on the same
    # line, so line numbers are unchanged). When an edit is pending, export the chunk's
    # variables and functions and unwind to the driver, which restarts from the edited
    # text at this statement boundary.
    $global:__DbgStmtEnd = $global:__DbgEnds[$k]
    if ($global:__DbgReloadAfter -lt 0) { return }
    $__vars = @{}
    foreach ($__v in (Get-Variable -Scope 1)) {
        $__n = $__v.Name
        if ($__n -like '__Dbg*' -or $global:__DbgPerScope.ContainsKey($__n)) { continue }
        if ($__v.Value -is [System.Collections.IEnumerator]) { continue }
        if (($__v.Options -band [System.Management.Automation.ScopedItemOptions]::Constant) -ne 0) { continue }
        $__vars[$__n] = $__v.Value
    }
    $__funcs = @{}
    foreach ($__f in (Get-ChildItem function:)) {
        try { if ($__f.ScriptBlock.File -eq $global:__DbgShadow) { $__funcs[$__f.Name] = $__f.ScriptBlock } } catch { }
    }
    $global:__DbgVars = $__vars
    $global:__DbgFuncs = $__funcs
    throw '__DbgReload'
}

function global:__DbgBuildChunk($text, $stmts, $idx, $inject) {
    # Text of statements idx.. to the end of the script, padded so that line numbers
    # match the editor, with a __DbgCheck call in front of every top-level statement.
    $__first = $stmts[$idx].Extent
    $__sb = New-Object System.Text.StringBuilder
    if ($inject) {
        $null = $__sb.Append($inject)
        if ($__first.StartLineNumber -ge 2) {
            $null = $__sb.Append("`n" * ($__first.StartLineNumber - 1))
            $null = $__sb.Append(' ' * ($__first.StartColumnNumber - 1))
        } else { $null = $__sb.Append('; ') }
    } else {
        $null = $__sb.Append("`n" * ($__first.StartLineNumber - 1))
        $null = $__sb.Append(' ' * ($__first.StartColumnNumber - 1))
    }
    # The check for statement k is appended to the END of statement k-1 (same line), so
    # every statement stays the first thing on its line: PowerShell line breakpoints
    # only bind to the first statement of a line.
    $__pos = $__first.StartOffset
    for ($__k = $idx; $__k -lt $stmts.Count - 1; $__k++) {
        $__x = $stmts[$__k].Extent
        $null = $__sb.Append($text.Substring($__pos, $__x.EndOffset - $__pos))
        $null = $__sb.Append("; __DbgCheck $($__k + 1)")
        $__pos = $__x.EndOffset
    }
    $null = $__sb.Append($text.Substring($__pos))
    return $__sb.ToString()
}

function global:__DbgStatements($ast) {
    $__s = @()
    if ($ast.ParamBlock) { $__s += $ast.ParamBlock }
    if ($ast.EndBlock) { $__s += $ast.EndBlock.Statements }
    return $__s
}

$global:__DbgAction = {
    $__ln = $_.Line
    if ($global:__DbgSw.ElapsedMilliseconds -ge 100) { if (__DbgRefreshControl) { __DbgSyncBreakpoints } }
    $__ctl = $global:__DbgCtl
    $__mode = $__ctl.mode
    $__isBp = $__ctl.bps.Contains($__ln) -or ($__ctl.temp -eq $__ln)
    $__trackAll = ($__ctl.track -eq 'all')
    $__isCp = $__ctl.cps.Contains($__ln)
    # Fast path: nothing to do on this line.
    if (-not $__isBp -and $__mode -eq 'run' -and -not $__trackAll -and -not $__isCp) { return }
    if ($global:__DbgInjecting) { return }
    $__stack = @(Get-PSCallStack)
    # A breakpoint on a line without a statement fires against the first statement that
    # does run; only the hit whose line matches the frame's current line is real.
    if ($__stack.Count -lt 2 -or $__stack[1].ScriptLineNumber -ne $__ln) { return }
    try { if ($__stack[1].Position.Text -like '__DbgCheck*') { return } } catch { }
    $__depth = $__stack.Count
    $__stop = $__isBp -or $__mode -eq 'step' -or
              ($__mode -eq 'step_over' -and $__depth -le $__ctl.depth) -or
              ($__mode -eq 'step_out' -and $__depth -lt $__ctl.depth)
    if (-not $__stop -and -not $__trackAll -and -not $__isCp) { return }
    if ($__ctl.temp -eq $__ln) { $__ctl.temp = -1 }
    $__list = New-Object System.Collections.Generic.List[object]
    $__scopes = New-Object System.Collections.Generic.List[object]
    $__i = 1
    while ($__i -lt 64) {
        try { $__sv = @(Get-Variable -Scope $__i -ErrorAction Stop) } catch { break }
        $__isHarness = $false
        foreach ($__x in $__sv) { if ($__x.Name -eq '__DbgHarness') { $__isHarness = $true; break } }
        if ($__isHarness) { break }
        $__scopes.Add($__sv)
        $__i++
    }
    $__n = $__scopes.Count
    for ($__k = 0; $__k -lt $__n; $__k++) {
        if ($__k -eq $__n - 1) { $__label = 'Script' }
        elseif ($__k -eq 0) { $__label = 'Local' }
        else { $__label = "Caller $__k" }
        if ($__k -lt $__n - 1 -and ($__k + 1) -lt $__stack.Count) {
            $__fn = $__stack[$__k + 1].Command
            if ($__fn) { $__label = "$__label ($__fn)" }
        }
        foreach ($__v in $__scopes[$__k]) {
            $__name = $__v.Name
            if ($__name -like '__Dbg*') { continue }
            $__auto = [bool]$global:__DbgAuto.ContainsKey($__name)
            $__val = $__v.Value
            $__type = if ($null -eq $__val) { 'null' } else { $__val.GetType().Name }
            if ($__val -is [System.Collections.IEnumerator]) { $__str = '<enumerator>' }
            elseif ($__auto) { $__str = try { __DbgShort $__val 200 } catch { '?' } }
            else { $__str = __DbgFormat $__val }
            $__list.Add(@{ n = $__name; v = $__str; t = $__type; s = $__label; a = $__auto; i = ($__k + 1) })
        }
    }
    foreach ($__v in (Get-Variable -Scope Global)) {
        $__name = $__v.Name
        if ($__name -like '__Dbg*' -or $__name -eq 'args' -or $global:__DbgAuto.ContainsKey($__name)) { continue }
        $__val = $__v.Value
        $__type = if ($null -eq $__val) { 'null' } else { $__val.GetType().Name }
        $__list.Add(@{ n = $__name; v = (__DbgFormat $__val); t = $__type; s = 'Global'; a = $false; i = 'Global' })
    }
    $global:__DbgSeq++
    $__json = ConvertTo-Json -InputObject $__list.ToArray() -Compress -Depth 3
    [IO.File]::WriteAllText((Join-Path $global:__DbgStateDir "vars_$($global:__DbgSeq).json"), $__json, [Text.Encoding]::UTF8)
    if (-not $__stop) {
        [Console]::Out.WriteLine("__DBG_TRACE__ $__ln $__depth $($global:__DbgStmtEnd) $($global:__DbgGen) $($global:__DbgSeq)")
        [Console]::Out.Flush()
        return
    }
    [Console]::Out.WriteLine("__DBG_STOP__ $__ln $__depth $($global:__DbgStmtEnd) $($global:__DbgGen) $($global:__DbgSeq)")
    [Console]::Out.Flush()
    while ($true) {
        $__cmd = [Console]::In.ReadLine()
        if ($null -eq $__cmd -or $__cmd -eq 'go') { break }
        if ($__cmd.StartsWith('dump|')) {
            $__parts = $__cmd.Split('|', 4)
            try {
                if ($__parts[1] -eq 'Global') { $__dv = Get-Variable -Scope Global -Name $__parts[2] -ErrorAction Stop }
                else { $__dv = Get-Variable -Scope ([int]$__parts[1]) -Name $__parts[2] -ErrorAction Stop }
                __DbgWriteValue $__dv.Value $__parts[3]
                __DbgMsg "dump-ok|$($__parts[3])"
            } catch { __DbgMsg "dump-error|$($_.Exception.Message)" }
        } elseif ($__cmd.StartsWith('reload|')) {
            $__parts = $__cmd.Split('|')
            if ($global:__DbgStmtEnd -le 0) {
                __DbgMsg "reload-error|Live editing is not available for this script (it uses begin/process/end blocks or using statements)."
            } else {
                $__t = $null; $__e = $null
                try {
                    $__src = [IO.File]::ReadAllText($global:__DbgSourceFile)
                    $null = [System.Management.Automation.Language.Parser]::ParseInput($__src, [ref]$__t, [ref]$__e)
                } catch { $__e = @($_.Exception.Message) }
                if ($__e -and $__e.Count -gt 0) {
                    $__m = $__e[0]
                    if ($__m.Extent) { __DbgMsg "reload-error|Line $($__m.Extent.StartLineNumber): $($__m.Message)" }
                    else { __DbgMsg "reload-error|$__m" }
                } else {
                    $global:__DbgReloadAfter = [int]$__parts[1]
                    $global:__DbgReloadGen = [int]$__parts[2]
                    __DbgMsg "reload-ok|$($__parts[2])"
                }
            }
        }
    }
    $global:__DbgCtlTime = [DateTime]::MinValue
    $null = __DbgRefreshControl
    __DbgSyncBreakpoints
}

$null = __DbgRefreshControl
$__DbgText = [IO.File]::ReadAllText($global:__DbgSourceFile)
$__DbgTokens = $null; $__DbgErrors = $null
$__DbgAst = [System.Management.Automation.Language.Parser]::ParseInput($__DbgText, [ref]$__DbgTokens, [ref]$__DbgErrors)
if ($__DbgErrors -and $__DbgErrors.Count -gt 0) {
    foreach ($__DbgE in $__DbgErrors) { [Console]::Error.WriteLine("Syntax error at line $($__DbgE.Extent.StartLineNumber): $($__DbgE.Message)") }
    exit 1
}
$__DbgWhole = ($null -ne $__DbgAst.BeginBlock -or $null -ne $__DbgAst.ProcessBlock -or $null -ne $__DbgAst.DynamicParamBlock -or
               ($__DbgAst.UsingStatements -and $__DbgAst.UsingStatements.Count -gt 0))
$__DbgExit = 0
if ($__DbgWhole) {
    __DbgMsg 'mode|whole'
    [IO.File]::WriteAllText($global:__DbgShadow, $__DbgText, [Text.Encoding]::UTF8)
    __DbgSetBreakpoints ((($__DbgText -split "`n").Count) + 1)
    try {
        $__DbgInvoke = [scriptblock]::Create('$__DbgHarness = $true; & $global:__DbgShadow ' + $__DbgArgString)
        & $__DbgInvoke
        if ($null -ne $LASTEXITCODE) { $__DbgExit = $LASTEXITCODE }
    } catch {
        [Console]::Error.WriteLine("Unhandled error: " + $_.Exception.Message)
        if ($_.InvocationInfo) { [Console]::Error.WriteLine($_.InvocationInfo.PositionMessage) }
        $__DbgExit = 1
    }
} else {
    __DbgMsg 'mode|driver'
    # The script runs as ONE compiled chunk (native speed) with a __DbgCheck call spliced in
    # front of every top-level statement. When an edit is applied, the next check exports the
    # variables and functions, unwinds, and a new chunk built from the edited text resumes at
    # that statement with everything re-injected.
    $__DbgStmts = __DbgStatements $__DbgAst
    $__DbgIdx = 0
    $__DbgInject = $null
    try {
        while ($__DbgIdx -lt $__DbgStmts.Count) {
            $global:__DbgEnds = @($__DbgStmts | ForEach-Object { $_.Extent.EndLineNumber })
            $global:__DbgStmtEnd = $global:__DbgEnds[$__DbgIdx]
            $__DbgChunk = __DbgBuildChunk $__DbgText $__DbgStmts $__DbgIdx $__DbgInject
            [IO.File]::WriteAllText($global:__DbgShadow, $__DbgChunk, [Text.Encoding]::UTF8)
            __DbgSetBreakpoints ((($__DbgChunk -split "`n").Count) + 1)
            if ($__DbgInject) {
                $__DbgInvoke = [scriptblock]::Create('$__DbgHarness = $true; $global:__DbgTopDepth = @(Get-PSCallStack).Count + 2; & $global:__DbgShadow')
            } else {
                $__DbgInvoke = [scriptblock]::Create('$__DbgHarness = $true; $global:__DbgTopDepth = @(Get-PSCallStack).Count + 2; & $global:__DbgShadow ' + $__DbgArgString)
            }
            $global:__DbgInjecting = [bool]$__DbgInject
            $__DbgReload = $false
            try {
                & $__DbgInvoke
            } catch {
                if ($_.FullyQualifiedErrorId -eq '__DbgReload' -or $_.Exception.Message -eq '__DbgReload') { $__DbgReload = $true }
                else { throw }
            }
            if (-not $__DbgReload) { break }
            # --- restart from the edited text ------------------------------------------
            $__DbgNewText = [IO.File]::ReadAllText($global:__DbgSourceFile)
            $__DbgTokens = $null; $__DbgErrors = $null
            $__DbgNewAst = [System.Management.Automation.Language.Parser]::ParseInput($__DbgNewText, [ref]$__DbgTokens, [ref]$__DbgErrors)
            $__DbgAfter = $global:__DbgReloadAfter
            if ($__DbgErrors -and $__DbgErrors.Count -gt 0) {
                __DbgMsg "info|Edited script has a syntax error; continuing with the previous version."
                $__DbgAfter = $global:__DbgStmtEnd - 1
            } else {
                $__DbgText = $__DbgNewText
                $__DbgAst = $__DbgNewAst
                $__DbgStmts = __DbgStatements $__DbgAst
                $global:__DbgGen = $global:__DbgReloadGen
            }
            $global:__DbgReloadAfter = -1
            $__DbgIdx = $__DbgStmts.Count
            for ($__DbgJ = 0; $__DbgJ -lt $__DbgStmts.Count; $__DbgJ++) {
                if ($__DbgStmts[$__DbgJ] -is [System.Management.Automation.Language.ParamBlockAst]) { continue }
                if ($__DbgStmts[$__DbgJ].Extent.StartLineNumber -gt $__DbgAfter) { $__DbgIdx = $__DbgJ; break }
            }
            # Injection line: explicit assignments so the compiler gives the variables fast local slots.
            $__DbgSb = New-Object System.Text.StringBuilder
            foreach ($__DbgN in $global:__DbgVars.Keys) {
                if ($__DbgN -match '^[A-Za-z_][A-Za-z0-9_]*$') { $null = $__DbgSb.Append("`$$__DbgN = `$global:__DbgVars['$__DbgN']; ") }
                else { $null = $__DbgSb.Append("Set-Variable -Name '" + $__DbgN.Replace("'", "''") + "' -Value `$global:__DbgVars['" + $__DbgN.Replace("'", "''") + "']; ") }
            }
            foreach ($__DbgN in $global:__DbgFuncs.Keys) {
                $null = $__DbgSb.Append("Set-Item -Path ('function:' + '" + $__DbgN.Replace("'", "''") + "') -Value `$global:__DbgFuncs['" + $__DbgN.Replace("'", "''") + "']; ")
            }
            $null = $__DbgSb.Append('$global:__DbgInjecting = $false')
            $__DbgInject = $__DbgSb.ToString()
            $__DbgNext = if ($__DbgIdx -lt $__DbgStmts.Count) { $__DbgStmts[$__DbgIdx].Extent.StartLineNumber } else { 0 }
            __DbgMsg "reloaded|$($global:__DbgGen)|$__DbgNext"
        }
        if ($null -ne $LASTEXITCODE) { $__DbgExit = $LASTEXITCODE }
    } catch {
        [Console]::Error.WriteLine("Unhandled error: " + $_.Exception.Message)
        if ($_.InvocationInfo) { [Console]::Error.WriteLine($_.InvocationInfo.PositionMessage) }
        $__DbgExit = 1
    }
}
[Console]::Out.Flush()
exit $__DbgExit
'''


# ----------------------------------------------------------------------------
# Debug sessions
# ----------------------------------------------------------------------------

class Var:
    __slots__ = ("name", "value", "type", "scope", "auto", "index")

    def __init__(self, name, value, type_="", scope="", auto=False, index=""):
        self.name = name
        self.value = value
        self.type = type_
        self.scope = scope
        self.auto = auto
        self.index = index      # scope index for live lookups (PowerShell)

    @property
    def key(self):
        return (self.scope, self.name.lower())


class DebugSession:
    """Runs a script in a child process; the script itself decides where to stop.

    Breakpoints, capture points and the run mode are written to a control file the
    running script re-reads. Events posted to `events` queue:
      ("output", text)
      ("stop", line, depth, [Var...], paused, stmt_end)   paused=False for capture points
      ("msg", kind, payload)
      ("end", exit_code, killed)
    """

    kind = "generic"
    supports_depth = False
    supports_live_edit = False

    def __init__(self, script_path, lines, args, events):
        self.script_path = os.path.abspath(script_path)
        self.lines = list(lines)            # text as last applied to the running script
        self.args = args or ""
        self.events = events
        self.proc = None
        self.lock = threading.Lock()
        self.breakpoints = set()
        self.capture_points = set()
        self.temp_breakpoint = None
        self.mode = "run"
        self.mode_depth = 0                 # absolute depth as reported by the script
        self.track = TRACK_POINTS
        self.state_dir = tempfile.mkdtemp(prefix="scriptdbg_")
        self.encoding = "utf-8"
        self._stdin_lock = threading.Lock()
        self._reply = None
        self._reply_event = threading.Event()
        self.exit_code = None
        self.killed = False
        self.paused = False
        self.current_line = None
        self.current_depth = 0              # relative to the script's top level
        self.current_depth_abs = 0
        self.current_stmt_end = 0
        self.base_depth = None
        self.gen = 0
        self.gen_maps = {}
        self.reload_pending = False
        self.shadow_path = None
        self.started = False

    # --- control -----------------------------------------------------------
    def set_breakpoints(self, bps):
        with self.lock:
            self.breakpoints = set(bps)
        self.write_control()

    def set_capture_points(self, cps):
        with self.lock:
            self.capture_points = set(cps)
        self.write_control()

    def set_temp_breakpoint(self, line):
        with self.lock:
            self.temp_breakpoint = line
        self.write_control()

    def set_mode(self, mode, depth_abs=0):
        with self.lock:
            self.mode = mode
            self.mode_depth = depth_abs
        self.write_control()

    def set_track(self, track):
        with self.lock:
            self.track = track
        self.write_control()

    def _control_lines(self, kind):
        """Breakpoint/capture lines in the coordinates the running script may still use."""
        with self.lock:
            lines = set(self.breakpoints if kind == "bps" else self.capture_points)
            if self.reload_pending and self.gen > 0:
                m = self.gen_maps.get(self.gen - 1, {})
                inverse = {}
                for old, new in m.items():
                    inverse.setdefault(new, old)
                lines |= {inverse[ln] for ln in lines if ln in inverse}
        return sorted(lines)

    def write_control(self):
        if not self.state_dir or not os.path.isdir(self.state_dir):
            return
        with self.lock:
            ctl = {"mode": self.mode, "depth": self.mode_depth, "temp": self.temp_breakpoint or -1,
                   "track": self.track}
        ctl["bps"] = self._control_lines("bps")
        ctl["cps"] = self._control_lines("cps")
        try:
            write_file_atomic(os.path.join(self.state_dir, self._control_name()), self._control_text(ctl))
        except OSError:
            pass

    def map_line(self, gen, line):
        """Translate a line reported for text generation `gen` into current numbering."""
        with self.lock:
            while gen < self.gen:
                m = self.gen_maps.get(gen)
                if m:
                    line = m.get(line, line)
                gen += 1
        return line

    def start(self):
        cmd, env, cwd = self._build()
        self.write_control()
        self.proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            cwd=cwd,
            env=env,
            creationflags=CREATE_NO_WINDOW,
        )
        self.started = True
        threading.Thread(target=self._reader, daemon=True).start()

    def resume(self, mode, depth_abs=0):
        self.set_mode(mode, depth_abs)
        self.paused = False
        self.send_line("go")

    def send_line(self, text):
        with self._stdin_lock:
            try:
                self.proc.stdin.write((text + "\r\n").encode("utf-8", "replace"))
                self.proc.stdin.flush()
            except (OSError, ValueError, AttributeError):
                pass

    def request(self, command, timeout=20):
        """Send a command while paused and wait for the script's reply (kind, payload)."""
        self._reply_event.clear()
        self._reply = None
        self.send_line(command)
        if not self._reply_event.wait(timeout):
            raise EditError("The paused script did not respond in time.")
        return self._reply

    def stop(self):
        self.killed = True
        if self.proc and self.proc.poll() is None:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(self.proc.pid)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW,
            )

    def dump_variable(self, var, path):
        """Write a variable's full value to `path`; default uses the captured value."""
        with open(path, "w", encoding="utf-8", newline="") as fh:
            fh.write(var.value)
        return None

    def apply_edit(self, new_lines):
        raise EditError("Live editing is not supported for this script type.")

    # --- reader thread -----------------------------------------------------
    def _reader(self):
        try:
            for raw in iter(self.proc.stdout.readline, b""):
                text = raw.decode(self.encoding, "replace")
                positions = [(text.find(m), m) for m in (STOP_MARKER, TRACE_MARKER, MSG_MARKER)]
                positions = [(p, m) for p, m in positions if p >= 0]
                if not positions:
                    self.events.put(("output", text))
                    continue
                idx, marker = min(positions)
                if idx > 0:
                    self.events.put(("output", text[:idx] + "\n"))
                body = text[idx + len(marker):].strip()
                if marker == MSG_MARKER:
                    kind, _, payload = body.partition("|")
                    if kind in ("dump-ok", "dump-error", "reload-ok", "reload-error"):
                        self._reply = (kind, payload)
                        self._reply_event.set()
                    else:
                        if kind == "reloaded":
                            self.reload_pending = False
                            self.write_control()
                        self.events.put(("msg", kind, payload))
                    continue
                parts = body.split()
                try:
                    line = int(parts[0])
                    depth = int(parts[1]) if len(parts) > 1 else 0
                    stmt_end = int(parts[2]) if len(parts) > 2 else 0
                    gen = int(parts[3]) if len(parts) > 3 else 0
                    seq = int(parts[4]) if len(parts) > 4 else 0
                except (IndexError, ValueError):
                    if marker == STOP_MARKER:
                        self.send_line("go")
                    continue
                if self.base_depth is None:
                    self.base_depth = depth
                rel_depth = depth - self.base_depth
                line = self.map_line(gen, line)
                stmt_end = self.map_line(gen, stmt_end) if stmt_end else 0
                try:
                    variables = self._read_vars(seq)
                except Exception as exc:  # noqa: BLE001
                    variables = []
                    self.events.put(("output", f"[debugger] could not read variables: {exc}\n"))
                pause = marker == STOP_MARKER
                if pause:
                    with self.lock:
                        if self.temp_breakpoint == line:
                            self.temp_breakpoint = None
                    self.current_line, self.current_depth, self.current_depth_abs = line, rel_depth, depth
                    self.current_stmt_end = stmt_end
                    self.paused = True
                self.events.put(("stop", line, rel_depth, variables, pause, stmt_end))
        except Exception as exc:  # noqa: BLE001
            self.events.put(("output", f"[debugger] reader error: {exc}\n"))
        finally:
            code = self.proc.wait()
            self.exit_code = code
            self._reply_event.set()
            self._cleanup()
            self.events.put(("end", code, self.killed))

    def _cleanup(self):
        shutil.rmtree(self.state_dir, ignore_errors=True)
        if self.shadow_path:
            try:
                os.remove(self.shadow_path)
            except OSError:
                pass

    def _shadow_target(self, ext):
        folder = os.path.dirname(self.script_path)
        base = os.path.splitext(os.path.basename(self.script_path))[0]
        name = f"~dbg_{base}{ext}"
        path = os.path.join(folder, name)
        try:
            with open(path, "wb"):
                pass
            return path
        except OSError:
            self.events.put(("output", "[debugger] script folder is not writable; running an instrumented "
                                       "copy from a temp folder ($PSScriptRoot / %~dp0 will point there)\n"))
            return os.path.join(self.state_dir, name)

    def _vars_file(self, seq, ext):
        path = os.path.join(self.state_dir, f"vars_{seq}{ext}")
        with open(path, "rb") as fh:
            data = fh.read()
        try:
            os.remove(path)
        except OSError:
            pass
        return data

    # --- to implement ------------------------------------------------------
    def _build(self):
        raise NotImplementedError

    def _read_vars(self, seq):
        raise NotImplementedError

    def _control_name(self):
        raise NotImplementedError

    def _control_text(self, ctl):
        raise NotImplementedError


class BatchSession(DebugSession):
    kind = "batch"
    supports_depth = False
    supports_live_edit = True

    def __init__(self, script_path, lines, args, events, text_encoding=None):
        super().__init__(script_path, lines, args, events)
        self.initial_vars = None
        self.encoding = oem_encoding()
        self.text_encoding = text_encoding or self.encoding
        self.offsets = {}
        self.lead_spaces = {}

    def _control_name(self):
        return "control.txt"

    def _control_text(self, ctl):
        # Read by the instrumented script with `for /f`, never executed, so it is not
        # subject to script scanning on every line.
        mode = "step" if ctl["mode"] != "run" else "run"
        bps = list(ctl["bps"])
        if ctl["temp"] and ctl["temp"] > 0:
            bps.append(ctl["temp"])
        text = (f'__DBG_BPS=,{",".join(str(b) for b in bps)},\r\n'
                f'__DBG_CPS=,{",".join(str(c) for c in ctl["cps"])},\r\n'
                f'__DBG_MODE={mode}\r\n'
                f'__DBG_TRACK={ctl["track"]}\r\n')
        return text.encode("ascii")

    def _encode_lines(self, lines):
        return [ln.encode(self.text_encoding, "replace") for ln in lines]

    def _build(self):
        base, ext = os.path.splitext(os.path.basename(self.script_path))
        if ext.lower() not in BATCH_EXTS:
            ext = ".cmd"
        self.shadow_path = self._shadow_target(ext)
        data, _, self.offsets = build_batch(self._encode_lines(self.lines), BATCH_INITIAL_PAD)
        with open(self.shadow_path, "wb") as fh:
            fh.write(data)
        env = dict(os.environ)
        env["__DBG_DIR"] = self.state_dir
        inner = f'"{self.shadow_path}"'
        if self.args.strip():
            inner += " " + self.args.strip()
        cmd = f'cmd.exe /d /s /c "{inner}"'
        return cmd, env, os.path.dirname(self.script_path)

    def _read_vars(self, seq):
        text = self._vars_file(seq, ".txt").decode(self.encoding, "replace")
        raw = parse_set_output(text)
        if self.initial_vars is None:
            # Environment as it was before the script's first line ran.
            try:
                self.initial_vars = parse_set_output(self._vars_file(0, ".txt").decode(self.encoding, "replace"))
            except OSError:
                self.initial_vars = dict(raw)
        out = []
        for name in sorted(raw, key=str.lower):
            value = raw[name]
            auto = (name not in ("ERRORLEVEL", "CD")
                    and name in self.initial_vars and self.initial_vars[name] == value)
            out.append(Var(name, value, "", "", auto))
        return out

    def apply_edit(self, new_lines):
        """Rewrite the instrumented file while paused, keeping cmd's resume offset valid.

        Returns the line number (in the new text) execution will resume at, or None if
        the script now ends there.
        """
        if not self.paused or self.current_line is None:
            raise EditError("Edits can only be applied while the script is paused.")
        cur = self.current_line
        old_off = self.offsets.get(cur)
        if old_off is None:
            raise EditError("Cannot determine where the script is paused.")
        start_map, _ = line_mapping(self.lines, new_lines)
        new_cur = start_map.get(cur, len(new_lines) + 1)
        new_cur_key = new_cur if new_cur <= len(new_lines) else len(new_lines) + 1
        new_bytes = self._encode_lines(new_lines)
        lead = {}
        data, _, rel = build_batch(new_bytes, 0)
        target = rel[new_cur_key]
        need = old_off - target
        if need < 0:
            raise EditError(
                f"This edit adds {-need} more bytes before the paused line than the running "
                f"script can absorb ({BATCH_INITIAL_PAD} bytes of slack per run). Move the "
                f"inserted text below the paused line, or Stop and Run again.")
        if 0 < need < 6:
            if new_cur_key <= len(new_lines):
                lead[new_cur_key] = need
            else:
                new_bytes = new_bytes + [b" " * need]
            need = 0
        data, _, offsets = build_batch(new_bytes, need, lead)
        if offsets[new_cur_key] != old_off:
            raise EditError("Internal error: could not preserve the resume position.")
        with open(self.shadow_path, "r+b") as fh:
            fh.write(data)
            fh.truncate()
        self.offsets = offsets
        self.lead_spaces = lead
        self.lines = list(new_lines)
        self.current_line = new_cur if new_cur <= len(new_lines) else None
        return self.current_line


class PowerShellSession(DebugSession):
    kind = "powershell"
    supports_depth = True
    supports_live_edit = True

    def __init__(self, script_path, lines, args, events, ps_exe="powershell.exe"):
        super().__init__(script_path, lines, args, events)
        self.ps_exe = ps_exe or "powershell.exe"
        self.encoding = "utf-8"
        self.live_mode = None

    def _control_name(self):
        return "control.json"

    def _control_text(self, ctl):
        return json.dumps(ctl).encode("utf-8")

    def _write_source(self, lines):
        with open(os.path.join(self.state_dir, "source.ps1"), "w", encoding="utf-8-sig", newline="\n") as fh:
            fh.write("\n".join(lines) + "\n")

    def _build(self):
        harness = os.path.join(self.state_dir, "harness.ps1")
        with open(harness, "w", encoding="utf-8-sig", newline="\r\n") as fh:
            fh.write(PS_HARNESS)
        self._write_source(self.lines)
        self.shadow_path = self._shadow_target(".ps1")

        def q(s):
            return "'" + s.replace("'", "''") + "'"

        command = f"& {q(harness)} {q(self.script_path)} {q(self.state_dir)} {q(self.shadow_path)}; exit $LASTEXITCODE"
        cmd = [self.ps_exe, "-NoProfile", "-NoLogo", "-ExecutionPolicy", "Bypass", "-Command", command]
        env = dict(os.environ)
        env["__DBG_ARGS"] = self.args.strip()
        return cmd, env, os.path.dirname(self.script_path)

    def _read_vars(self, seq):
        data = json.loads(self._vars_file(seq, ".json").decode("utf-8-sig"))
        if isinstance(data, dict):
            data = [data]
        out = []
        for item in data or []:
            out.append(Var(str(item.get("n", "")), str(item.get("v", "")), str(item.get("t", "")),
                           str(item.get("s", "")), bool(item.get("a", False)), str(item.get("i", ""))))
        return out

    def dump_variable(self, var, path):
        if self.paused and var.index:
            kind, payload = self.request(f"dump|{var.index}|{var.name}|{path}")
            if kind != "dump-ok":
                raise EditError(payload)
            return None
        super().dump_variable(var, path)
        if TRUNCATED_MARK in var.value:
            return ("Only a preview of this value was captured. Pause the script at a "
                    "breakpoint and save again to get the complete value.")
        return None

    def apply_edit(self, new_lines):
        """Hand the edited text to the harness; it switches over at the next statement."""
        if not self.paused:
            raise EditError("Edits can only be applied while the script is paused.")
        if self.live_mode == "whole":
            raise EditError("Live editing is not available for this script: it uses "
                            "begin/process/end blocks or using statements.")
        start_map, end_map = line_mapping(self.lines, new_lines)
        after = end_map.get(self.current_stmt_end, len(new_lines)) if self.current_stmt_end else 0
        new_gen = self.gen + 1
        self._write_source(new_lines)
        kind, payload = self.request(f"reload|{after}|{new_gen}")
        if kind != "reload-ok":
            self._write_source(self.lines)      # leave the last good text in place
            raise EditError(payload)
        with self.lock:
            self.gen_maps[self.gen] = start_map
            self.gen = new_gen
            self.reload_pending = True
        self.lines = list(new_lines)
        self.current_line = start_map.get(self.current_line, self.current_line)
        self.current_stmt_end = after
        return self.current_line


def make_session(path, lines, args, events, ps_exe=None, text_encoding=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in BATCH_EXTS:
        return BatchSession(path, lines, args, events, text_encoding)
    if ext in PS_EXTS:
        return PowerShellSession(path, lines, args, events, ps_exe)
    raise ValueError(f"Unsupported script type: {ext or '(none)'}")


def find_powershells():
    found = []
    sysroot = os.environ.get("SystemRoot", r"C:\Windows")
    ps51 = os.path.join(sysroot, "System32", "WindowsPowerShell", "v1.0", "powershell.exe")
    if os.path.exists(ps51):
        found.append(("PowerShell 5.1", ps51))
    pwsh = shutil.which("pwsh.exe")
    if pwsh:
        found.append(("PowerShell 7", pwsh))
    if not found:
        found.append(("powershell.exe", "powershell.exe"))
    return found


# ----------------------------------------------------------------------------
# Breakpoint persistence
# ----------------------------------------------------------------------------

def _bp_store_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    folder = os.path.join(base, "ScriptDebugger")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "breakpoints.json")


def load_saved_points(path):
    """Return (breakpoints, capture_points) saved for `path`."""
    try:
        with open(_bp_store_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        entry = data.get(os.path.abspath(path).lower(), [])
        if isinstance(entry, dict):
            return (set(int(x) for x in entry.get("bp", [])), set(int(x) for x in entry.get("cp", [])))
        return set(int(x) for x in entry), set()
    except Exception:  # noqa: BLE001
        return set(), set()


def save_points(path, bps, cps):
    try:
        store = _bp_store_path()
        try:
            with open(store, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception:  # noqa: BLE001
            data = {}
        key = os.path.abspath(path).lower()
        if bps or cps:
            data[key] = {"bp": sorted(bps), "cp": sorted(cps)}
        else:
            data.pop(key, None)
        with open(store, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=1)
    except Exception:  # noqa: BLE001
        pass


# ----------------------------------------------------------------------------
# GUI widgets
# ----------------------------------------------------------------------------

class CodeView(ttk.Frame):
    """Editable source view with a clickable breakpoint / capture-point gutter."""

    def __init__(self, master, on_toggle_breakpoint, on_toggle_capture, on_run_to_line, on_text_changed):
        super().__init__(master)
        self.on_toggle_breakpoint = on_toggle_breakpoint
        self.on_toggle_capture = on_toggle_capture
        self.on_run_to_line = on_run_to_line
        self.on_text_changed = on_text_changed
        self.breakpoints = set()
        self.capture_points = set()
        self.executable = None
        self.line_count = 0
        self.kind = "batch"
        self._change_job = None
        self._loading = False

        self.font = ("Consolas", 10)
        self.gutter = tk.Canvas(self, width=64, bg="#eceff3", highlightthickness=0)
        self.text = tk.Text(self, wrap="none", font=self.font, undo=True, maxundo=-1, autoseparators=True,
                            bg="#ffffff", fg="#1c1c1c", selectbackground="#b8d4f5",
                            inactiveselectbackground="#b8d4f5", padx=6, spacing1=1, spacing3=1,
                            insertwidth=2)
        self.vs = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self.hs = ttk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=self._on_yscroll, xscrollcommand=self.hs.set)

        self.gutter.grid(row=0, column=0, sticky="ns")
        self.text.grid(row=0, column=1, sticky="nsew")
        self.vs.grid(row=0, column=2, sticky="ns")
        self.hs.grid(row=1, column=1, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        self.text.tag_configure("cpline", background="#e4eefb")
        self.text.tag_configure("bpline", background="#ffe3e3")
        self.text.tag_configure("current", background="#fff2a8")
        self.text.tag_configure("comment", foreground="#6a8f3c")
        self.text.tag_configure("label", foreground="#8a4bb8")
        self.text.tag_configure("keyword", foreground="#1f4fb8")
        self.text.tag_configure("variable", foreground="#b05a00")
        self.text.tag_configure("string", foreground="#a31515")
        self.text.tag_raise("bpline")
        self.text.tag_raise("current")
        self.text.tag_raise("sel")

        self.gutter.bind("<Button-1>", self._gutter_click)
        self.gutter.bind("<Control-Button-1>", self._gutter_ctrl_click)
        self.gutter.bind("<Button-3>", self._gutter_ctrl_click)
        self.text.bind("<Button-3>", self._context_menu)
        self.text.bind("<Configure>", lambda e: self.redraw_gutter())
        self.text.bind("<MouseWheel>", lambda e: self.after_idle(self.redraw_gutter))
        self.text.bind("<KeyRelease>", lambda e: self.after_idle(self.redraw_gutter))
        self.text.bind("<ButtonRelease-1>", lambda e: self.after_idle(self.redraw_gutter))
        self.text.bind("<<Modified>>", self._modified)
        self.text.bind("<Control-o>", lambda e: "break")     # keep Text's default from inserting a newline
        self.gutter.bind("<MouseWheel>", self._wheel_on_gutter)

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Toggle breakpoint   (F9)",
                              command=lambda: self.on_toggle_breakpoint(self.cursor_line))
        self.menu.add_command(label="Toggle capture point   (Ctrl+F9)",
                              command=lambda: self.on_toggle_capture(self.cursor_line))
        self.menu.add_command(label="Run to this line", command=lambda: self.on_run_to_line(self.cursor_line))
        self.menu.add_separator()
        self.menu.add_command(label="Comment out line(s)", command=self.comment_selection)
        self.menu.add_command(label="Uncomment line(s)", command=self.uncomment_selection)

    @property
    def cursor_line(self):
        try:
            return int(self.text.index("insert").split(".")[0])
        except (ValueError, tk.TclError):
            return 1

    # scrolling -------------------------------------------------------------
    def _on_yscroll(self, *args):
        self.vs.set(*args)
        self.redraw_gutter()

    def _yview(self, *args):
        self.text.yview(*args)
        self.redraw_gutter()

    def _wheel_on_gutter(self, event):
        self.text.yview_scroll(int(-event.delta / 120) * 3, "units")
        self.redraw_gutter()

    # content ---------------------------------------------------------------
    def get_lines(self):
        content = self.text.get("1.0", "end-1c")
        return content.split("\n")

    def load(self, lines, kind):
        self._loading = True
        self.kind = kind
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", "\n".join(lines))
        self.text.edit_reset()
        self.text.edit_modified(False)
        self._loading = False
        self.refresh_content(lines)
        self.text.mark_set("insert", "1.0")
        self.set_current(None)

    def refresh_content(self, lines):
        self.line_count = max(1, len(lines))
        self._highlight(lines, self.kind)
        digits = len(str(self.line_count))
        self.gutter.configure(width=34 + 8 * digits)
        self.update_idletasks()
        self.redraw_gutter()

    def set_editable(self, editable):
        self.text.configure(state="normal" if editable else "disabled",
                            cursor="xterm" if editable else "arrow")

    def _modified(self, event):
        if self._loading or not self.text.edit_modified():
            return
        self.text.edit_modified(False)
        if self._change_job:
            self.after_cancel(self._change_job)
        self._change_job = self.after(150, self._fire_change)

    def _fire_change(self):
        self._change_job = None
        self.on_text_changed(self.get_lines())

    def comment_selection(self):
        self._prefix_lines(add=True)

    def uncomment_selection(self):
        self._prefix_lines(add=False)

    def _prefix_lines(self, add):
        if str(self.text.cget("state")) != "normal":
            return
        try:
            first = int(self.text.index("sel.first").split(".")[0])
            last = int(self.text.index("sel.last").split(".")[0])
            if self.text.index("sel.last").endswith(".0") and last > first:
                last -= 1
        except tk.TclError:
            first = last = self.cursor_line
        prefix = "rem " if self.kind == "batch" else "# "
        for ln in range(first, last + 1):
            line = self.text.get(f"{ln}.0", f"{ln}.end")
            if add:
                self.text.insert(f"{ln}.0", prefix)
            else:
                stripped = line.lstrip()
                lead = len(line) - len(stripped)
                low = stripped.lower()
                if self.kind == "batch":
                    for p in ("rem ", "rem\t", "::"):
                        if low.startswith(p):
                            self.text.delete(f"{ln}.{lead}", f"{ln}.{lead + len(p)}")
                            break
                    else:
                        if low == "rem":
                            self.text.delete(f"{ln}.{lead}", f"{ln}.{lead + 3}")
                elif stripped.startswith("# "):
                    self.text.delete(f"{ln}.{lead}", f"{ln}.{lead + 2}")
                elif stripped.startswith("#"):
                    self.text.delete(f"{ln}.{lead}", f"{ln}.{lead + 1}")

    def _highlight(self, lines, kind):
        for tag in ("comment", "label", "keyword", "variable", "string"):
            self.text.tag_remove(tag, "1.0", "end")
        if kind == "batch":
            kw = re.compile(r"(?i)\b(echo|set|if|else|for|in|do|goto|call|exit|setlocal|endlocal|rem|pause|"
                            r"not|exist|defined|errorlevel|shift|pushd|popd|cd|del|copy|move|md|mkdir|rd|rmdir|"
                            r"type|start|choice|timeout|findstr|find|cls|title|color|prompt)\b")
            var = re.compile(r"%[^%\s]*%|![^!\s]*!|%%?~?[a-zA-Z0-9]\b")
            for i, line in enumerate(lines, 1):
                s = line.strip().lstrip("@").strip()
                low = s.lower()
                if low.startswith("::") or low == "rem" or low.startswith("rem ") or low.startswith("rem\t"):
                    self.text.tag_add("comment", f"{i}.0", f"{i}.end")
                    continue
                if s.startswith(":"):
                    self.text.tag_add("label", f"{i}.0", f"{i}.end")
                    continue
                for m in kw.finditer(line):
                    self.text.tag_add("keyword", f"{i}.{m.start()}", f"{i}.{m.end()}")
                for m in var.finditer(line):
                    self.text.tag_add("variable", f"{i}.{m.start()}", f"{i}.{m.end()}")
        else:
            kw = re.compile(r"(?i)\b(function|param|if|elseif|else|foreach|for|while|do|switch|return|try|catch|"
                            r"finally|throw|break|continue|in|begin|process|end|class|enum|using|exit)\b")
            var = re.compile(r"\$\{[^}]+\}|\$[A-Za-z_][\w:]*|\$[\?\^\$_]")
            strs = re.compile(r"'[^']*'|\"(?:[^\"`]|`.)*\"")
            in_block = False
            for i, line in enumerate(lines, 1):
                s = line.strip()
                if in_block:
                    self.text.tag_add("comment", f"{i}.0", f"{i}.end")
                    if "#>" in s:
                        in_block = False
                    continue
                if s.startswith("<#"):
                    self.text.tag_add("comment", f"{i}.0", f"{i}.end")
                    in_block = "#>" not in s
                    continue
                for m in kw.finditer(line):
                    self.text.tag_add("keyword", f"{i}.{m.start()}", f"{i}.{m.end()}")
                for m in strs.finditer(line):
                    self.text.tag_add("string", f"{i}.{m.start()}", f"{i}.{m.end()}")
                for m in var.finditer(line):
                    self.text.tag_add("variable", f"{i}.{m.start()}", f"{i}.{m.end()}")
                hash_pos = self._comment_start(line)
                if hash_pos is not None:
                    for tag in ("keyword", "string", "variable"):
                        self.text.tag_remove(tag, f"{i}.{hash_pos}", f"{i}.end")
                    self.text.tag_add("comment", f"{i}.{hash_pos}", f"{i}.end")

    @staticmethod
    def _comment_start(line):
        in_s = in_d = False
        for idx, ch in enumerate(line):
            if ch == "'" and not in_d:
                in_s = not in_s
            elif ch == '"' and not in_s:
                in_d = not in_d
            elif ch == "#" and not in_s and not in_d:
                return idx
        return None

    # gutter ----------------------------------------------------------------
    def redraw_gutter(self):
        self.gutter.delete("all")
        if self.line_count == 0:
            return
        width = int(self.gutter.cget("width"))
        idx = self.text.index("@0,0")
        while True:
            dli = self.text.dlineinfo(idx)
            if dli is None:
                break
            y, h = dli[1], dli[3]
            ln = int(idx.split(".")[0])
            if ln > self.line_count:
                break
            cy = y + h / 2
            if ln in self.breakpoints:
                r = min(6, h // 2 - 1)
                self.gutter.create_oval(4, cy - r, 4 + 2 * r, cy + r, fill="#d9342b", outline="#8c1d16")
            if ln in self.capture_points:
                r = min(6, h // 2 - 1)
                x0 = 18
                self.gutter.create_polygon(x0 + r, cy - r, x0 + 2 * r, cy, x0 + r, cy + r, x0, cy,
                                           fill="#2f6fd6", outline="#1b4a9c")
            color = "#5a6472" if self.executable is None or ln in self.executable else "#b5bcc6"
            self.gutter.create_text(width - 4, cy, text=str(ln), anchor="e", font=self.font, fill=color)
            nxt = self.text.index(f"{idx}+1line")
            if nxt == idx or int(nxt.split(".")[0]) == ln:
                break
            idx = nxt

    def _line_at(self, y):
        return int(self.text.index(f"@0,{y}").split(".")[0])

    def _gutter_click(self, event):
        ln = self._line_at(event.y)
        if 1 <= ln <= self.line_count:
            self.on_toggle_breakpoint(ln)

    def _gutter_ctrl_click(self, event):
        ln = self._line_at(event.y)
        if 1 <= ln <= self.line_count:
            self.on_toggle_capture(ln)

    def _context_menu(self, event):
        try:
            self.text.mark_set("insert", f"@{event.x},{event.y}")
        except tk.TclError:
            pass
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def set_points(self, bps, cps):
        self.breakpoints = set(bps)
        self.capture_points = set(cps)
        self.text.tag_remove("bpline", "1.0", "end")
        self.text.tag_remove("cpline", "1.0", "end")
        for ln in self.capture_points:
            self.text.tag_add("cpline", f"{ln}.0", f"{ln + 1}.0")
        for ln in self.breakpoints:
            self.text.tag_add("bpline", f"{ln}.0", f"{ln + 1}.0")
        self.redraw_gutter()

    def set_current(self, line):
        self.text.tag_remove("current", "1.0", "end")
        if line and 1 <= line <= self.line_count:
            self.text.tag_add("current", f"{line}.0", f"{line + 1}.0")
            self.text.see(f"{line}.0")
            self.after_idle(self.redraw_gutter)


class VarTable(ttk.Frame):
    """Variables pane with filter, 'show automatic/environment' toggle and save-to-file."""

    def __init__(self, master, title, on_save):
        super().__init__(master)
        self.on_save = on_save
        top = ttk.Frame(self)
        top.pack(fill="x", padx=4, pady=(4, 2))
        ttk.Label(top, text=title, font=("Segoe UI", 9, "bold")).pack(side="left")
        self.show_auto = tk.BooleanVar(value=False)
        self.auto_check = ttk.Checkbutton(top, text="Show inherited / automatic", variable=self.show_auto,
                                          command=self.refresh)
        self.auto_check.pack(side="right")
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *a: self.refresh())
        ttk.Entry(top, textvariable=self.filter_var, width=14).pack(side="right", padx=(0, 8))
        ttk.Label(top, text="Filter:").pack(side="right")
        ttk.Button(top, text="Save value...", command=self._save_selected).pack(side="left", padx=(10, 0))
        ttk.Label(top, text="red = changed, green = new", foreground="#7a8290").pack(side="left", padx=(10, 0))

        cols = ("name", "value", "type", "scope")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for col, text, width in (("name", "Name", 150), ("value", "Value", 320),
                                 ("type", "Type", 90), ("scope", "Scope", 110)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, minwidth=40, stretch=(col == "value"))
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=(0, 4))
        vs.pack(side="left", fill="y", pady=(0, 4))
        self.tree.tag_configure("changed", foreground="#c42c1c")
        self.tree.tag_configure("new", foreground="#1a7f37")
        self.tree.tag_configure("auto", foreground="#7a8290")
        self.tree.bind("<Double-1>", self._show_full)
        self.tree.bind("<Button-3>", self._context_menu)
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Save value to file...", command=self._save_selected)
        self.menu.add_command(label="Show full value", command=self._show_full)
        self.vars = {}
        self.changed = set()
        self.new = set()

    def update(self, vars_map, changed=(), new=()):
        self.vars = vars_map
        self.changed = set(changed)
        self.new = set(new)
        self.refresh()

    def clear(self):
        self.update({}, (), ())

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        flt = self.filter_var.get().strip().lower()
        show_auto = self.show_auto.get()
        order = {"Local": 0, "Caller": 1, "Script": 2, "Global": 3}

        def sort_key(item):
            v = item[1]
            return (order.get(v.scope.split(" ")[0], 1), v.scope, v.name.lower())

        for key, v in sorted(self.vars.items(), key=sort_key):
            if v.auto and not show_auto and key not in self.changed and key not in self.new:
                continue
            if flt and flt not in v.name.lower():
                continue
            tags = ()
            if key in self.new:
                tags = ("new",)
            elif key in self.changed:
                tags = ("changed",)
            elif v.auto:
                tags = ("auto",)
            display = v.value.replace("\r", "").replace("\n", "\u21b5 ")
            if len(display) > 400:
                display = display[:400] + " \u2026"
            self.tree.insert("", "end", iid=f"{v.scope}\x00{v.name}",
                             values=(v.name, display, v.type, v.scope), tags=tags)

    def selected_var(self):
        sel = self.tree.selection()
        if not sel:
            return None
        scope, name = sel[0].split("\x00", 1)
        return self.vars.get((scope, name.lower()))

    def _context_menu(self, event):
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.menu.grab_release()

    def _save_selected(self):
        v = self.selected_var()
        if v:
            self.on_save(v)

    def _show_full(self, event=None):
        v = self.selected_var()
        if not v:
            return
        win = tk.Toplevel(self)
        win.title(f"{v.name}")
        win.geometry("640x400")
        txt = tk.Text(win, wrap="word", font=("Consolas", 10))
        txt.pack(fill="both", expand=True)
        head = f"Name:  {v.name}\n"
        if v.type:
            head += f"Type:  {v.type}\n"
        if v.scope:
            head += f"Scope: {v.scope}\n"
        txt.insert("1.0", head + "\nValue:\n" + v.value)
        txt.configure(state="disabled")
        ttk.Button(win, text="Save value to file...", command=lambda: self.on_save(v)).pack(pady=4)


class HistoryTable(ttk.Frame):
    """Chronological log of every variable change seen at pauses / capture points."""

    def __init__(self, master):
        super().__init__(master)
        top = ttk.Frame(self)
        top.pack(fill="x", padx=4, pady=(4, 2))
        ttk.Label(top, text="Variable history (changes between captures, in execution order)",
                  font=("Segoe UI", 9, "bold")).pack(side="left")
        self.count = ttk.Label(top, text="0 changes", foreground="#7a8290")
        self.count.pack(side="right")
        cols = ("n", "line", "name", "old", "new")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", selectmode="browse")
        for col, text, width, stretch in (("n", "#", 50, False), ("line", "Lines", 80, False),
                                          ("name", "Variable", 140, False),
                                          ("old", "Old value", 200, True), ("new", "New value", 240, True)):
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, minwidth=40, stretch=stretch)
        vs = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=(0, 4))
        vs.pack(side="left", fill="y", pady=(0, 4))
        self.tree.tag_configure("new", foreground="#1a7f37")
        self.tree.tag_configure("removed", foreground="#7a8290")
        self.n = 0
        self.on_select_line = None
        self.tree.bind("<<TreeviewSelect>>", self._select)

    def clear(self):
        self.tree.delete(*self.tree.get_children())
        self.n = 0
        self.count.configure(text="0 changes")

    def add(self, where, name, old, new, kind="changed"):
        self.n += 1

        def fmt(s):
            return s.replace("\r", "").replace("\n", "\u21b5 ")[:300]

        self.tree.insert("", "end", values=(self.n, where, name, fmt(old), fmt(new)), tags=(kind,))
        self.count.configure(text=f"{self.n} changes")
        self.tree.yview_moveto(1.0)

    def _select(self, event):
        sel = self.tree.selection()
        if sel and self.on_select_line:
            vals = self.tree.item(sel[0], "values")
            m = re.findall(r"\d+", str(vals[1]))
            if m:
                self.on_select_line(int(m[-1]))


# ----------------------------------------------------------------------------
# Main application
# ----------------------------------------------------------------------------

TRACK_CHOICES = [("Pauses + capture points (fast)", TRACK_POINTS), ("Every line (slow)", TRACK_ALL)]


class DebuggerApp(tk.Tk):
    def __init__(self, initial_path=None):
        super().__init__()
        self.title(APP_TITLE)
        width = min(1360, self.winfo_screenwidth() - 40)
        height = min(840, self.winfo_screenheight() - 80)
        self.geometry(f"{width}x{height}+10+10")
        self.minsize(900, 600)
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        self.script_path = None
        self.script_lines = []          # what the editor currently shows
        self.applied_lines = []         # what the running script was last given
        self.file_lines = []            # what is on disk
        self.file_encoding = "utf-8"
        self.file_bom = False
        self.kind = None
        self.breakpoints = set()
        self.capture_points = set()
        self.executable = None
        self.session = None
        self.events = queue.Queue()
        self.prev_vars = None
        self.last_line = None
        self.state = "idle"
        self.edits_pending = False

        self._build_menu()
        self._build_toolbar()
        self._build_panes()
        self._build_statusbar()
        self._bind_keys()
        self._set_state("idle")
        self.after(30, self._poll_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        if initial_path:
            self.after(100, lambda: self.open_script(initial_path))

    # --- UI construction ---------------------------------------------------
    def _build_menu(self):
        menubar = tk.Menu(self)
        filem = tk.Menu(menubar, tearoff=0)
        filem.add_command(label="Open script...", accelerator="Ctrl+O", command=self.open_dialog)
        filem.add_command(label="Save script", accelerator="Ctrl+S", command=self.save_file)
        filem.add_command(label="Reload from disk", accelerator="Ctrl+R", command=self.reload_script)
        filem.add_separator()
        filem.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=filem)
        dbg = tk.Menu(menubar, tearoff=0)
        dbg.add_command(label="Run / Continue", accelerator="F5", command=self.cmd_run)
        dbg.add_command(label="Pause", accelerator="Ctrl+Break", command=self.cmd_pause)
        dbg.add_command(label="Step Into", accelerator="F11", command=self.cmd_step_into)
        dbg.add_command(label="Step Over", accelerator="F10", command=self.cmd_step_over)
        dbg.add_command(label="Step Out", accelerator="Shift+F11", command=self.cmd_step_out)
        dbg.add_command(label="Stop", accelerator="Shift+F5", command=self.cmd_stop)
        dbg.add_separator()
        dbg.add_command(label="Apply edits to running script", accelerator="Ctrl+Enter", command=self.apply_edits)
        dbg.add_separator()
        dbg.add_command(label="Toggle breakpoint at cursor", accelerator="F9",
                        command=lambda: self.toggle_breakpoint(self.code.cursor_line))
        dbg.add_command(label="Toggle capture point at cursor", accelerator="Ctrl+F9",
                        command=lambda: self.toggle_capture_point(self.code.cursor_line))
        dbg.add_command(label="Clear all breakpoints and capture points", command=self.clear_points)
        menubar.add_cascade(label="Debug", menu=dbg)
        helpm = tk.Menu(menubar, tearoff=0)
        helpm.add_command(label="How it works", command=self.show_help)
        menubar.add_cascade(label="Help", menu=helpm)
        self.config(menu=menubar)

    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(6, 4))
        bar.pack(fill="x")
        ttk.Button(bar, text="Open...", command=self.open_dialog).pack(side="left")
        self.btn_save = ttk.Button(bar, text="Save", command=self.save_file)
        self.btn_save.pack(side="left", padx=(4, 0))
        self.file_label = ttk.Label(bar, text="No script loaded", width=24, anchor="w")
        self.file_label.pack(side="left", padx=(8, 4))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=4)
        self.btn_run = ttk.Button(bar, text="\u25b6 Run  (F5)", command=self.cmd_run)
        self.btn_run.pack(side="left", padx=2)
        self.btn_pause = ttk.Button(bar, text="\u23f8 Pause", command=self.cmd_pause)
        self.btn_pause.pack(side="left", padx=2)
        self.btn_step_into = ttk.Button(bar, text="Step Into (F11)", width=14, command=self.cmd_step_into)
        self.btn_step_into.pack(side="left", padx=2)
        self.btn_step_over = ttk.Button(bar, text="Step Over (F10)", width=14, command=self.cmd_step_over)
        self.btn_step_over.pack(side="left", padx=2)
        self.btn_step_out = ttk.Button(bar, text="Step Out", command=self.cmd_step_out)
        self.btn_step_out.pack(side="left", padx=2)
        self.btn_stop = ttk.Button(bar, text="\u25a0 Stop", command=self.cmd_stop)
        self.btn_stop.pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=4)
        self.btn_apply = ttk.Button(bar, text="Apply edits", command=self.apply_edits)
        self.btn_apply.pack(side="left", padx=2)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=4)
        ttk.Label(bar, text="Track vars:").pack(side="left")
        self.track_var = tk.StringVar(value=TRACK_CHOICES[0][0])
        self.track_combo = ttk.Combobox(bar, textvariable=self.track_var, state="readonly", width=27,
                                        values=[c[0] for c in TRACK_CHOICES])
        self.track_combo.pack(side="left", padx=(4, 8))
        self.track_combo.bind("<<ComboboxSelected>>", lambda e: self._track_changed())
        ttk.Label(bar, text="Args:").pack(side="left")
        self.args_var = tk.StringVar()
        ttk.Entry(bar, textvariable=self.args_var, width=16).pack(side="left", padx=(4, 8))
        self.ps_choices = find_powershells()
        self.ps_var = tk.StringVar(value=self.ps_choices[0][0])
        self.ps_combo = ttk.Combobox(bar, textvariable=self.ps_var, state="readonly", width=14,
                                     values=[c[0] for c in self.ps_choices])
        self.ps_label = ttk.Label(bar, text="PS:")

    def _build_panes(self):
        outer = ttk.PanedWindow(self, orient="horizontal")
        outer.pack(fill="both", expand=True, padx=4, pady=(0, 2))

        left = ttk.PanedWindow(outer, orient="vertical")
        self.code = CodeView(left, self.toggle_breakpoint, self.toggle_capture_point, self.run_to_line,
                             self.on_text_changed)
        left.add(self.code, weight=4)

        out_frame = ttk.Frame(left)
        head = ttk.Frame(out_frame)
        head.pack(fill="x", padx=4, pady=(4, 2))
        ttk.Label(head, text="Script output", font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Button(head, text="Clear", command=self.clear_output).pack(side="right")
        self.output = tk.Text(out_frame, height=10, wrap="word", font=("Consolas", 10), bg="#1e1e1e",
                              fg="#d8d8d8", insertbackground="#d8d8d8", state="disabled")
        ovs = ttk.Scrollbar(out_frame, orient="vertical", command=self.output.yview)
        self.output.configure(yscrollcommand=ovs.set)
        self.output.tag_configure("dbg", foreground="#7cc7ff")
        inp = ttk.Frame(out_frame)
        inp.pack(side="bottom", fill="x", padx=4, pady=(2, 4))
        ttk.Label(inp, text="Send to script stdin:").pack(side="left")
        self.input_var = tk.StringVar()
        self.input_entry = ttk.Entry(inp, textvariable=self.input_var)
        self.input_entry.pack(side="left", fill="x", expand=True, padx=4)
        self.input_entry.bind("<Return>", lambda e: self.send_input())
        ttk.Button(inp, text="Send", command=self.send_input).pack(side="left")
        self.output.pack(side="left", fill="both", expand=True, padx=(4, 0))
        ovs.pack(side="left", fill="y")
        left.add(out_frame, weight=1)
        outer.add(left, weight=3)

        right = ttk.PanedWindow(outer, orient="vertical")
        self.vars_table = VarTable(right, "Variables", self.save_variable)
        right.add(self.vars_table, weight=3)
        self.history = HistoryTable(right)
        self.history.on_select_line = self._jump_to_line
        right.add(self.history, weight=2)
        outer.add(right, weight=2)

    def _build_statusbar(self):
        bar = ttk.Frame(self, relief="sunken", padding=(6, 2))
        bar.pack(fill="x", side="bottom")
        self.status = ttk.Label(bar, text="Open a .bat, .cmd or .ps1 script to begin.")
        self.status.pack(side="left")
        self.bp_status = ttk.Label(bar, text="", foreground="#7a8290")
        self.bp_status.pack(side="right")

    def _bind_keys(self):
        self.bind("<F5>", lambda e: self.cmd_run())
        self.bind("<Shift-F5>", lambda e: self.cmd_stop())
        self.bind("<F10>", lambda e: self.cmd_step_over() or "break")
        self.bind("<F11>", lambda e: self.cmd_step_into())
        self.bind("<Shift-F11>", lambda e: self.cmd_step_out())
        self.bind("<F9>", lambda e: self.toggle_breakpoint(self.code.cursor_line))
        self.bind("<Control-F9>", lambda e: self.toggle_capture_point(self.code.cursor_line))
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<Control-s>", lambda e: self.save_file())
        self.bind("<Control-r>", lambda e: self.reload_script())
        self.bind("<Control-Return>", lambda e: self.apply_edits())
        self.bind("<Control-Pause>", lambda e: self.cmd_pause())

    # --- state ---------------------------------------------------------------
    def _set_state(self, state):
        self.state = state
        loaded = state in ("loaded", "finished")
        running = state == "running"
        paused = state == "paused"
        if running or paused:
            depth_ok = bool(self.session and self.session.supports_depth)
        else:
            depth_ok = self.kind == "powershell"

        def en(btn, cond):
            btn.state(["!disabled"] if cond else ["disabled"])

        en(self.btn_run, loaded or paused)
        self.btn_run.configure(text="\u25b6 Continue  (F5)" if paused else "\u25b6 Run  (F5)")
        en(self.btn_pause, running)
        en(self.btn_step_into, loaded or paused)
        en(self.btn_step_over, (loaded or paused) and depth_ok)
        en(self.btn_step_out, paused and depth_ok)
        en(self.btn_stop, running or paused)
        en(self.btn_save, bool(self.script_path))
        en(self.btn_apply, paused and self.edits_pending)
        self.code.set_editable(loaded or paused)
        if self.kind == "powershell":
            self.ps_label.pack(side="left")
            self.ps_combo.pack(side="left", padx=(4, 0))
        else:
            self.ps_label.pack_forget()
            self.ps_combo.pack_forget()
        self._update_point_status()
        self._update_title()

    def _update_point_status(self):
        if self.script_path:
            self.bp_status.configure(
                text=f"{len(self.breakpoints)} breakpoint(s), {len(self.capture_points)} capture point(s)")
        else:
            self.bp_status.configure(text="")

    def _update_title(self):
        if not self.script_path:
            self.title(APP_TITLE)
            return
        name = os.path.basename(self.script_path)
        flags = ""
        if self.script_lines != self.file_lines:
            flags += " *"
        if self.edits_pending and self.state == "paused":
            flags += "  [edits not yet applied]"
        self.title(f"{APP_TITLE} - {name}{flags}")

    def _status(self, text):
        self.status.configure(text=text)

    def _track_mode(self):
        for label, value in TRACK_CHOICES:
            if label == self.track_var.get():
                return value
        return TRACK_POINTS

    def _track_changed(self):
        if self.session:
            self.session.set_track(self._track_mode())
        if self._track_mode() == TRACK_ALL:
            self._status("Tracking every line: complete history, but slow on long loops and large values.")
        else:
            self._status("Tracking at pauses and capture points only (Ctrl+click the gutter to add a capture point).")

    # --- file handling --------------------------------------------------------
    def _confirm_discard(self):
        if self.script_path and self.script_lines != self.file_lines:
            ans = messagebox.askyesnocancel(APP_TITLE, "The script has unsaved changes. Save them first?")
            if ans is None:
                return False
            if ans:
                self.save_file()
        return True

    def open_dialog(self):
        path = filedialog.askopenfilename(
            title="Open script",
            filetypes=[("Scripts", "*.bat *.cmd *.ps1"), ("Batch files", "*.bat *.cmd"),
                       ("PowerShell scripts", "*.ps1"), ("All files", "*.*")])
        if path:
            self.open_script(path)

    def open_script(self, path, confirm=True):
        if self.session and self.state in ("running", "paused"):
            if not messagebox.askyesno(APP_TITLE, "A script is still running. Stop it and open the new file?"):
                return
            self.cmd_stop()
        if confirm and not self._confirm_discard():
            return
        path = os.path.abspath(path)
        ext = os.path.splitext(path)[1].lower()
        if ext in BATCH_EXTS:
            kind = "batch"
        elif ext in PS_EXTS:
            kind = "powershell"
        else:
            messagebox.showerror(APP_TITLE, f"Unsupported file type '{ext}'. Supported: .bat, .cmd, .ps1")
            return
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Cannot read file:\n{exc}")
            return
        self.file_bom = data.startswith(b"\xef\xbb\xbf")
        if kind == "batch":
            if self.file_bom:
                text = data[3:].decode("utf-8", "replace")
                self.file_encoding = "utf-8"
            else:
                self.file_encoding = oem_encoding()
                text = data.decode(self.file_encoding, "replace")
        else:
            text = None
            for enc in ("utf-8-sig", "utf-16", "cp1252"):
                try:
                    text = data.decode(enc)
                    self.file_encoding = {"utf-8-sig": "utf-8"}.get(enc, enc)
                    break
                except UnicodeDecodeError:
                    continue
            if text is None:
                text = data.decode("utf-8", "replace")
                self.file_encoding = "utf-8"
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        if lines and lines[-1] == "" and len(lines) > 1:
            lines = lines[:-1]
        self.script_path = path
        self.script_lines = list(lines)
        self.file_lines = list(lines)
        self.applied_lines = list(lines)
        self.kind = kind
        self.executable = self._compute_executable(lines)
        self.code.executable = self.executable
        self.code.load(lines, kind)
        bps, cps = load_saved_points(path)
        self.breakpoints = {b for b in bps if 1 <= b <= len(lines)}
        self.capture_points = {c for c in cps if 1 <= c <= len(lines)}
        self.code.set_points(self.breakpoints, self.capture_points)
        self.file_label.configure(
            text=f"{os.path.basename(path)}   [{'Batch' if kind == 'batch' else 'PowerShell'}]")
        self.vars_table.clear()
        self.history.clear()
        self.clear_output()
        self.prev_vars = None
        self.edits_pending = False
        self._set_state("loaded")
        self._status(f"Loaded {len(lines)} lines. Click the gutter for a breakpoint (F9), Ctrl+click for a "
                     f"capture point (Ctrl+F9), then Run (F5) or Step Into (F11). Nothing runs until you do.")
        self.vars_table.auto_check.configure(
            text="Show inherited environment" if kind == "batch" else "Show automatic variables")

    def reload_script(self):
        if self.script_path:
            self.open_script(self.script_path)

    def save_file(self):
        if not self.script_path:
            return
        lines = self.code.get_lines()
        data = "\r\n".join(lines) + "\r\n"
        try:
            raw = data.encode(self.file_encoding, "replace")
            if self.file_bom:
                raw = b"\xef\xbb\xbf" + raw
            with open(self.script_path, "wb") as fh:
                fh.write(raw)
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not save:\n{exc}")
            return
        self.script_lines = list(lines)
        self.file_lines = list(lines)
        self._update_title()
        self._status(f"Saved {os.path.basename(self.script_path)}.")

    def _compute_executable(self, lines):
        if self.kind == "batch":
            return batch_executable_lines([ln.encode(self.file_encoding, "replace") for ln in lines])
        return ps_executable_lines(lines)

    # --- editing ----------------------------------------------------------------
    def on_text_changed(self, new_lines):
        if not self.script_path or new_lines == self.script_lines:
            return
        old_lines = self.script_lines
        start_map, _ = line_mapping(old_lines, new_lines)

        def remap(points):
            out = set()
            for p in points:
                target = start_map.get(p)
                if target and target <= len(new_lines):
                    out.add(target)
            return out

        self.breakpoints = remap(self.breakpoints)
        self.capture_points = remap(self.capture_points)
        self.script_lines = list(new_lines)
        self.executable = self._compute_executable(new_lines)
        self.code.executable = self.executable
        self.code.refresh_content(new_lines)
        self.code.set_points(self.breakpoints, self.capture_points)
        save_points(self.script_path, self.breakpoints, self.capture_points)
        if self.session and self.state == "paused":
            self.edits_pending = new_lines != self.applied_lines
            cur = self.session.current_line
            if cur:
                amap, _ = line_mapping(self.applied_lines, new_lines)
                self.code.set_current(amap.get(cur, cur))
            if self.edits_pending:
                self._status("Script edited. Apply edits (Ctrl+Enter) or just Continue/Step: the change is "
                             "applied to the running script before it resumes.")
        self._set_state(self.state)

    def apply_edits(self):
        """Push editor contents into the paused script. Returns True on success."""
        if not (self.session and self.state == "paused"):
            return False
        new_lines = self.code.get_lines()
        if new_lines == self.applied_lines:
            self.edits_pending = False
            self._set_state(self.state)
            return True
        if not self.session.supports_live_edit:
            messagebox.showerror(APP_TITLE, "Live editing is not supported for this script.")
            return False
        try:
            new_cur = self.session.apply_edit(new_lines)
        except EditError as exc:
            messagebox.showerror(APP_TITLE, f"The edit could not be applied:\n\n{exc}")
            self._status("Edit not applied. Fix the problem and try again, or Stop and Run.")
            return False
        self.applied_lines = list(new_lines)
        self.script_lines = list(new_lines)
        self.edits_pending = False
        self._push_points()
        self.code.set_current(new_cur)
        self._append_output("[debugger] edits applied to the running script\n", "dbg")
        if self.kind == "batch":
            if new_cur:
                self._status(f"Edits applied. Execution resumes at line {new_cur}.")
            else:
                self._status("Edits applied. The script will end when resumed (paused line was removed).")
        else:
            self._status("Edits applied. They take effect when the current top-level statement finishes "
                         "(a running loop keeps its old body until it completes).")
        self._set_state("paused")
        return True

    # --- breakpoints / capture points --------------------------------------------
    def _snap(self, line):
        if self.executable is None or line in self.executable:
            return line
        for ln in range(line + 1, len(self.script_lines) + 1):
            if ln in self.executable:
                return ln
        return None

    def _push_points(self):
        if self.session:
            with self.session.lock:
                self.session.breakpoints = set(self.breakpoints)
                self.session.capture_points = set(self.capture_points)
            self.session.write_control()

    def _toggle_point(self, line, points, label):
        if not self.script_path or not line:
            return
        target = self._snap(line)
        if target is None:
            self._status(f"No executable line at or after line {line}.")
            return
        if target in points:
            points.discard(target)
        else:
            points.add(target)
            if target != line:
                self._status(f"Line {line} cannot be stopped on; {label} placed on line {target} instead.")
        self.code.set_points(self.breakpoints, self.capture_points)
        save_points(self.script_path, self.breakpoints, self.capture_points)
        self._push_points()
        self._update_point_status()

    def toggle_breakpoint(self, line):
        self._toggle_point(line, self.breakpoints, "breakpoint")

    def toggle_capture_point(self, line):
        self._toggle_point(line, self.capture_points, "capture point")
        if line and self.script_path and self._track_mode() == TRACK_ALL:
            self._status("Capture points only matter in 'Pauses + capture points' tracking mode.")

    def clear_points(self):
        self.breakpoints = set()
        self.capture_points = set()
        self.code.set_points(self.breakpoints, self.capture_points)
        if self.script_path:
            save_points(self.script_path, self.breakpoints, self.capture_points)
        self._push_points()
        self._update_point_status()

    def run_to_line(self, line):
        if not self.script_path:
            return
        target = self._snap(line)
        if target is None:
            return
        if self.state in ("loaded", "finished"):
            self._start_session("run", temp_bp=target)
        elif self.state == "paused":
            self.session.set_temp_breakpoint(target)
            self._resume("run")

    # --- debug commands -------------------------------------------------------
    def cmd_run(self):
        if self.state in ("loaded", "finished"):
            self._start_session("run")
        elif self.state == "paused":
            self._resume("run")

    def cmd_step_into(self):
        if self.state in ("loaded", "finished"):
            self._start_session("step")
        elif self.state == "paused":
            self._resume("step")

    def cmd_step_over(self):
        if self.state in ("loaded", "finished"):
            self._start_session("step")
        elif self.state == "paused":
            if self.session.supports_depth:
                self._resume("step_over", self.session.current_depth_abs)
            else:
                self._resume("step")

    def cmd_step_out(self):
        if self.state == "paused" and self.session.supports_depth:
            self._resume("step_out", self.session.current_depth_abs)

    def cmd_pause(self):
        if self.state == "running" and self.session:
            self.session.set_mode("step")
            if self.kind == "powershell" and self._track_mode() != TRACK_ALL:
                self._status("Pausing at the next breakpoint, capture point or top-level statement "
                             "(a running loop with none of those cannot be interrupted; use Stop).")
            else:
                self._status("Pausing at the next line...")

    def cmd_stop(self):
        if self.session and self.state in ("running", "paused"):
            self._status("Stopping...")
            self.session.stop()

    def _start_session(self, mode, temp_bp=None):
        if not self.script_path:
            return
        lines = self.code.get_lines()
        self.script_lines = list(lines)
        ps_exe = None
        if self.kind == "powershell":
            for label, exe in self.ps_choices:
                if label == self.ps_var.get():
                    ps_exe = exe
        self.events = queue.Queue()
        try:
            self.session = make_session(self.script_path, lines, self.args_var.get(), self.events, ps_exe,
                                        text_encoding=self.file_encoding)
        except ValueError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.applied_lines = list(lines)
        self.edits_pending = False
        with self.session.lock:
            self.session.breakpoints = set(self.breakpoints)
            self.session.capture_points = set(self.capture_points)
            self.session.mode = mode
            self.session.track = self._track_mode()
            self.session.temp_breakpoint = temp_bp
        self.vars_table.clear()
        self.history.clear()
        self.clear_output()
        self.prev_vars = None
        self.last_line = None
        self.code.set_current(None)
        try:
            self.session.start()
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not start the script:\n{exc}")
            self.session = None
            return
        note = " (unsaved editor contents)" if lines != self.file_lines else ""
        self._append_output(f"[debugger] started {os.path.basename(self.script_path)}{note}\n", "dbg")
        self._set_state("running")
        self._status("Running...")

    def _resume(self, mode, depth_abs=0):
        if self.edits_pending or self.code.get_lines() != self.applied_lines:
            self.edits_pending = True
            if not self.apply_edits():
                return
        self.code.set_current(None)
        self._set_state("running")
        self._status("Running...")
        self.session.resume(mode, depth_abs)

    def send_input(self):
        text = self.input_var.get()
        if self.session and self.state in ("running", "paused"):
            self.session.send_line(text)
            self._append_output(text + "\n", "dbg")
            self.input_var.set("")
        else:
            self._status("The script is not running; nothing to send input to.")

    # --- variables ----------------------------------------------------------------
    def save_variable(self, var):
        if var is None:
            return
        if self.kind == "powershell":
            types = [("Text (console formatting)", "*.txt"), ("JSON", "*.json"),
                     ("Raw bytes (byte[] values)", "*.bin"), ("All files", "*.*")]
        else:
            types = [("Text", "*.txt"), ("All files", "*.*")]
        safe = re.sub(r"[^\w.-]", "_", var.name) or "value"
        path = filedialog.asksaveasfilename(title=f"Save value of {var.name}", defaultextension=".txt",
                                            initialfile=f"{safe}.txt", filetypes=types)
        if not path:
            return
        try:
            if self.session and self.state == "paused":
                note = self.session.dump_variable(var, path)
            else:
                with open(path, "w", encoding="utf-8", newline="") as fh:
                    fh.write(var.value)
                note = ("Only a preview of this value was captured. Pause the script at a breakpoint "
                        "and save again to get the complete value." if TRUNCATED_MARK in var.value else None)
        except (EditError, OSError) as exc:
            messagebox.showerror(APP_TITLE, f"Could not save the value:\n{exc}")
            return
        size = os.path.getsize(path) if os.path.exists(path) else 0
        self._status(f"Saved {var.name} to {path} ({size:,} bytes).")
        self._append_output(f"[debugger] saved {var.name} to {path} ({size:,} bytes)\n", "dbg")
        if note:
            messagebox.showwarning(APP_TITLE, note)

    # --- events from the session ---------------------------------------------
    def _poll_events(self):
        try:
            for _ in range(500):
                ev = self.events.get_nowait()
                self._handle_event(ev)
        except queue.Empty:
            pass
        self.after(30, self._poll_events)

    def _handle_event(self, ev):
        kind = ev[0]
        if kind == "output":
            self._append_output(ev[1])
        elif kind == "msg":
            _, mkind, payload = ev
            if mkind == "mode" and self.session:
                self.session.live_mode = payload
                if payload == "whole":
                    self._append_output("[debugger] this script uses begin/process/end or using statements; "
                                        "live editing is disabled for this run\n", "dbg")
            elif mkind == "reloaded":
                parts = payload.split("|")
                nxt = parts[1] if len(parts) > 1 else "?"
                self._append_output(f"[debugger] edited script is now in effect (next statement at line {nxt})\n", "dbg")
            elif mkind == "info":
                self._append_output(f"[debugger] {payload}\n", "dbg")
        elif kind == "stop":
            _, line, depth, variables, paused, stmt_end = ev
            self._apply_variables(line, variables, paused)
            if paused:
                self.code.set_current(line)
                self._set_state("paused")
                where = f"line {line}"
                if self.session and self.session.supports_depth and depth > 0:
                    where += f" (call depth {depth})"
                self._status(f"Paused before {where}. F5 continue, F11 step into, F10 step over. "
                             f"You can edit the script now.")
        elif kind == "end":
            _, code, killed = ev
            self.code.set_current(None)
            if self.prev_vars is not None:
                self.vars_table.update(self.prev_vars, (), ())
            self.edits_pending = False
            self._set_state("finished")
            if killed:
                self._append_output("[debugger] script stopped by user\n", "dbg")
                self._status("Stopped. Press Run to start again.")
            else:
                self._append_output(f"[debugger] script finished with exit code {code}\n", "dbg")
                self._status(f"Finished with exit code {code}. Press Run or Step to start again.")
            self.session = None

    def _apply_variables(self, line, variables, paused):
        cur = {v.key: v for v in variables}
        changed, new = set(), set()
        if self.prev_vars is not None:
            if self._track_mode() == TRACK_ALL or self.last_line is None:
                where = str(self.last_line or "")
            elif self.last_line == line:
                where = str(line)
            else:
                where = f"{self.last_line}\u2192{line}"
            for key, v in cur.items():
                old = self.prev_vars.get(key)
                if v.auto:
                    continue
                if old is None:
                    new.add(key)
                    self.history.add(where, v.name, "<new>", v.value, "new")
                elif old.value != v.value:
                    changed.add(key)
                    self.history.add(where, v.name, old.value, v.value)
            for key, old in self.prev_vars.items():
                if key not in cur and not old.auto:
                    self.history.add(where, old.name, old.value, "<removed>", "removed")
        self.prev_vars = cur
        self.last_line = line
        if paused:
            self.vars_table.update(cur, changed, new)

    def _jump_to_line(self, line):
        if line:
            self.code.text.see(f"{line}.0")
            self.code.redraw_gutter()

    # --- output ------------------------------------------------------------------
    def _append_output(self, text, tag=None):
        self.output.configure(state="normal")
        self.output.insert("end", text, (tag,) if tag else ())
        if int(self.output.index("end-1c").split(".")[0]) > 5000:
            self.output.delete("1.0", "1000.0")
        self.output.see("end")
        self.output.configure(state="disabled")

    def clear_output(self):
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.configure(state="disabled")

    # --- misc ----------------------------------------------------------------------
    def show_help(self):
        messagebox.showinfo(
            "How it works",
            "1. Open a .bat/.cmd or .ps1 file. It is loaded but NOT executed.\n"
            "2. Click in the gutter (or press F9) for a breakpoint: execution pauses there.\n"
            "   Ctrl+click the gutter (or Ctrl+F9) for a capture point (blue diamond): variables are "
            "recorded there without pausing.\n"
            "3. Press Run (F5) to execute until the first breakpoint, or Step Into (F11) to pause on the "
            "first line.\n"
            "4. While paused the Variables pane shows every variable as it is right before the highlighted "
            "line runs. Red = changed since the previous capture, green = newly created. Right-click a "
            "variable (or use 'Save value...') to write its complete value to a file.\n"
            "5. 'Track vars' chooses how much is recorded: 'Pauses + capture points' costs nothing on "
            "other lines; 'Every line' logs every assignment but is slow on long loops.\n"
            "6. While paused you can edit the script in the editor (right-click for comment/uncomment). "
            "Press Apply edits (Ctrl+Enter), or simply Continue/Step, and the running script picks up the "
            "change. Ctrl+S saves the editor contents back to the file.\n\n"
            "Batch: edits apply at the next line, even inside goto loops. The slack for text inserted "
            "ABOVE the paused line is 4 KB per run.\n"
            "PowerShell: the script runs one top-level statement at a time; edits take effect at the next "
            "top-level statement, so a loop that is currently running keeps its old body until it ends. "
            "Functions defined before the edit keep their original line numbers.\n\n"
            "Limitations: FOR-loop variables (%%i) are not environment variables and are not listed; "
            "scripts that read stdin (pause, set /p, Read-Host) get their input from the "
            "'Send to script stdin' box.")

    def _on_close(self):
        if self.session and self.state in ("running", "paused"):
            if not messagebox.askyesno(APP_TITLE, "A script is still running. Stop it and exit?"):
                return
            self.session.stop()
        if not self._confirm_discard():
            return
        self.destroy()


def main():
    path = sys.argv[1] if len(sys.argv) > 1 else None
    app = DebuggerApp(path)
    app.mainloop()


if __name__ == "__main__":
    main()
