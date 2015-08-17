using System;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ManageMediaKey
{
    /// <summary>The program.</summary>
    public class Program
    {
        private const int WhKeyboardLl = 13;

        private const int WmKeydown = 0x0100;

        private const int SwMaximize = 3;

        private const int SwMinimize = 6;

        private const uint WaitTimeout = 0x00000102;

        private static readonly LowLevelKeyboardProc Proc = HookCallback;

        private static IntPtr _hookId = IntPtr.Zero;

        private static IntPtr _miroHandle = IntPtr.Zero;

        private static IntPtr _processHandle = IntPtr.Zero;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        /// <summary>The main.</summary>
        public static void Main()
        {
            _hookId = SetHook(Proc);
            Application.Run();
            UnhookWindowsHookEx(_hookId);
        }

        private static IntPtr GetMiroHandle()
        {
            if (_processHandle != IntPtr.Zero)
            {
                if (WaitForSingleObject(_processHandle, 0) == WaitTimeout)
                {  
                    return _miroHandle; 
                }

                CloseHandle(_processHandle);
            }

            // reset
            _miroHandle = IntPtr.Zero;
            _processHandle = IntPtr.Zero;

            const string WmiQueryString = "SELECT ProcessId, ExecutablePath, Name FROM Win32_Process WHERE Name = 'Miro.exe'";

            using (var searcher = new ManagementObjectSearcher(WmiQueryString))
            using (var results = searcher.Get())
            {
                var query = from p in Process.GetProcesses()
                            join mo in results.Cast<ManagementObject>() 
                            on p.Id equals (int)(uint)mo["ProcessId"]
                            select new
                                       {
                                           Process = p,
                                       };

                var processItem = query.FirstOrDefault();
                if (processItem != null)
                {
                    var process = processItem.Process;
                    if (process != null)
                    {
                        _miroHandle = process.MainWindowHandle;
                        _processHandle = OpenProcess(ProcessAccessFlags.Synchronize, false, process.Id);
                    }
                }
            }

            return _miroHandle;
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curProcess = Process.GetCurrentProcess())
            using (var curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WhKeyboardLl, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WmKeydown)
            {
                var vkCode = (Keys)Marshal.ReadInt32(lParam);

                // switch to miro and "press" the space bar
                if (vkCode == Keys.MediaPlayPause)
                {
                    // set to miro and maximize
                    var handle = GetMiroHandle();

                    if (handle != IntPtr.Zero)
                    {
                        var currentHandle = GetCurrentAppHandle();

                        SetForegroundWindow(handle);
                        ShowWindow(handle, SwMaximize);

                        // send a space bar
                        SendKeys.Send(" ");

                        // minimize miro
                        ShowWindow(handle, SwMinimize);

                        // switch back to the application that initially had focus
                        SetForegroundWindow(currentHandle);                        
                    }
                }
            }

            // this passes the keystroke to the next program
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private static IntPtr GetCurrentAppHandle()
        {
            var hWnd = GetForegroundWindow();
            uint procId;
            GetWindowThreadProcessId(hWnd, out procId);
            var proc = Process.GetProcessById((int)procId);
            return proc.MainWindowHandle;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(ProcessAccessFlags processAccess, bool bInheritHandle, int processId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint WaitForSingleObject(IntPtr hHandle, uint dwMilliseconds);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr hObject);

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }
    }
}
