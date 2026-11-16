# Sample for the assembly-load pause and native debugger hand-off (benign).
# Builds a tiny .NET assembly, loads it from BYTES (no file on disk at load time),
# calls into it, then starts a child process.
$source = @'
public static class Payload {
    public static string Run(string who) { return "Payload says hello to " + who; }
}
'@
$dll = Join-Path $env:TEMP "scriptdbg_payload.dll"
Add-Type -TypeDefinition $source -OutputAssembly $dll -OutputType Library
$bytes = [IO.File]::ReadAllBytes($dll)
Remove-Item $dll

# --> With "pause on assembly load" enabled, the debugger stops right here, after the
#     load and before Payload.Run executes. Attach dnSpy/WinDbg to the PID shown.
$asm = [Reflection.Assembly]::Load($bytes)
$result = $asm.GetType("Payload").GetMethod("Run").Invoke($null, @("the debugger"))
Write-Output $result

# --> A child process: WinDbg with .childdbg 1 / sxe cpr breaks when it is created.
$proc = Start-Process cmd.exe -ArgumentList "/c echo child process ran" -PassThru -Wait -NoNewWindow
Write-Output "child exit code: $($proc.ExitCode)"
