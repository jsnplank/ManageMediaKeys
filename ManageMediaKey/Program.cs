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

        private static readonly LowLevelKeyboardProc Proc = HookCallback;

        private static IntPtr _hookId = IntPtr.Zero;

        private static IntPtr _miroHandle;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        /// <summary>The main.</summary>
        public static void Main()
        {
            _miroHandle = GetMiroHandle();

            _hookId = SetHook(Proc);
            Application.Run();
            UnhookWindowsHookEx(_hookId);
        }

        private static IntPtr GetMiroHandle()
        {
            var wmiQueryString = "SELECT ProcessId, ExecutablePath, Name FROM Win32_Process WHERE Name = 'Miro.exe'";
            using (var searcher = new ManagementObjectSearcher(wmiQueryString))
            using (var results = searcher.Get())
            {
                var query = from p in Process.GetProcesses()
                            join mo in results.Cast<ManagementObject>() 
                            on p.Id equals (int)(uint)mo["ProcessId"]
                            select new
                                       {
                                           Process = p,
                                       };

                var miroProcessId = IntPtr.Zero;
                foreach (var item in query)
                {
                    miroProcessId = item.Process.MainWindowHandle;
                }

                return miroProcessId;
            }
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
                if (vkCode == Keys.MediaPlayPause && _miroHandle != IntPtr.Zero)
                {
                    var currentHandle = GetCurrentAppHandle();

                    // set to miro and maximize
                    SetForegroundWindow(_miroHandle);
                    ShowWindow(_miroHandle, SwMaximize);

                    // send a space bar
                    SendKeys.Send(" ");

                    // minimize miro
                    ShowWindow(_miroHandle, SwMinimize);

                    // switch back to the application that initially had focus
                    SetForegroundWindow(currentHandle);
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
    }
}
