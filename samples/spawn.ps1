# Sample for the process guard (benign stand-in for process hollowing).
# Creates a child process SUSPENDED (as a hollowing loader would), does some "work" on it,
# resumes it, and the child then creates a grandchild. With the process guard on, each
# of these processes is frozen before it executes and can be handed to a native debugger.
$source = @'
using System;
using System.Runtime.InteropServices;
public static class Spawner {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct STARTUPINFO { public int cb; public string lpReserved, lpDesktop, lpTitle;
        public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public short wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError; }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int dwProcessId, dwThreadId; }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool CreateProcessW(string app, string cmd, IntPtr pa, IntPtr ta, bool inherit, uint flags,
        IntPtr env, string dir, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
    [DllImport("kernel32.dll")] static extern uint ResumeThread(IntPtr hThread);
    [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr h, uint ms);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr h);
    const uint CREATE_SUSPENDED = 0x4, CREATE_NO_WINDOW = 0x08000000;
    public static int RunSuspended(string commandLine) {
        var si = new STARTUPINFO(); si.cb = Marshal.SizeOf(typeof(STARTUPINFO));
        PROCESS_INFORMATION pi;
        if (!CreateProcessW(null, commandLine, IntPtr.Zero, IntPtr.Zero, false, CREATE_SUSPENDED | CREATE_NO_WINDOW,
                            IntPtr.Zero, null, ref si, out pi)) throw new Exception("CreateProcess failed");
        Console.WriteLine("created suspended PID " + pi.dwProcessId + " (a loader would inject here)");
        ResumeThread(pi.hThread);
        WaitForSingleObject(pi.hProcess, 60000);
        CloseHandle(pi.hThread); CloseHandle(pi.hProcess);
        return pi.dwProcessId;
    }
}
'@
Add-Type -TypeDefinition $source

$childPid = [Spawner]::RunSuspended('cmd.exe /c "echo child running & cmd.exe /c echo grandchild running"')
Write-Output "child $childPid finished"
