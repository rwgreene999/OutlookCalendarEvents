using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;



namespace OutlookCalendarEvents
{
    internal class Worker : BackgroundWorker
    {

        #region Win32 operations 
        private delegate bool EnumWindowsProc(IntPtr hWnd, ref RunOptions data);

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        private static readonly IntPtr HWND_TOP = new IntPtr(0);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private static readonly uint WM_CLOSE = 0x0010;
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 SWP_SHOWWINDOW = 0x0040;
        const UInt32 SWP_ASYNCWINDOWPOS = 0x4000; 
        private const UInt32 SWP_NOACTIVATE = 0x0010;


        private const UInt32 SW_MAXIMIZE = 3;
        private const UInt32 SW_MINIMIZE = 6;
        private const UInt32 SW_RESTORE = 9;
        private const UInt32 SW_SHOW = 5;
        private const UInt32 SW_SHOWNORMAL = 0;
        private const UInt32 SW_SHOWNOACTIVATE = 4;
        private const UInt32 SW_SHOWMINNOACTIVE = 7;

        const int WM_LBUTTONDOWN = 0x201;
        const int WM_LBUTTONUP = 0x202;



        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, ref RunOptions data);


        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hwnd, WindowEnumProc func, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdSHow);

        [DllImport("user32.dll")]
        static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("User32.Dll", EntryPoint = "PostMessageA")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint wMsg,
                                        int wParam, int lParam);

        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public ShowWindowCommands showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        internal enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            Minimized = 2,
            Maximized = 3,
        }
        #endregion

        public System.Threading.ManualResetEvent _busy = new System.Threading.ManualResetEvent(true);
        internal Worker()
        {

        }
        public void DoTheWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            RunOptions ops = (RunOptions)e.Argument;
            Debug.WriteLine($"OutlookCalendarWorker Started, looping every {ops.secondsBetweenCheck} seconds between checks"); 
            while( !e.Cancel)
            {                
                System.Threading.Thread.Sleep(ops.secondsBetweenCheck);
                _busy.Wait();
                EnumWindows(new EnumWindowsProc(EnumProc), ref ops);
            }
        }

        /// <summary>
        /// Called by each EnumWindows for each window found 
        /// </summary>
        /// <param name="hWndMainWindow"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public static bool EnumProc(IntPtr hWndMainWindow, ref RunOptions data)
        {
            StringBuilder sbTitle = new StringBuilder(512);
            if (GetWindowText(hWndMainWindow, sbTitle, sbTitle.Capacity) > 1)
            {
                string Title = sbTitle.ToString();
                if (Title.Contains("Reminder"))
                {
                    if ( Title.Contains(" Reminder")
                        && !Title.StartsWith("0 ")
                        )
                    {
                        ProcessOutlookEventWindow(hWndMainWindow, sbTitle, Title);
                    }
                    else
                    {
                        Debug.WriteLine($"{Title} not a match");
                    }
                }
            }
            return true;
        }

        private static void ProcessOutlookEventWindow(IntPtr hWndMainWindow, StringBuilder sbTitle, string Title)
        {
            TryAndEmptyEventsFromWindow(hWndMainWindow);

            if (GetWindowText(hWndMainWindow, sbTitle, sbTitle.Capacity) > 1
                && !(sbTitle.ToString()).StartsWith("0")
                )
            {
                HandleNotificationToUser(hWndMainWindow, sbTitle);
            }
            else
            {
                AttemptToCloseTheWindow(hWndMainWindow, sbTitle);
            }
        }

        private static void AttemptToCloseTheWindow(IntPtr hWndMainWindow, StringBuilder sbTitle)
        {
            PostMessage(hWndMainWindow, WM_CLOSE, 0, 0);
            //ShowWindow(hWndMainWindow, SW_SHOWMINNOACTIVE);
            // Debug.WriteLine($"{sbTitle} now says zero events");
        }

        private static void HandleNotificationToUser(IntPtr hWndMainWindow, StringBuilder sbTitle)
        {
            NotifyFlashAllWindows.Flash(NotifyFlashAllWindows.ScreenSide.all);
        }

        private static void TryAndEmptyEventsFromWindow(IntPtr hWndMainWindow)
        {
            // dealing with a quirk of Outlook 
            // After you (the user) Dismiss or Snooze calendar events, the line item is removed, but 
            // the Reminder window title will still contain the number of previous events.
            // SO .... 
            // We find the SysListView inside, and click it.  After a couple fo seconds, this cause the 
            // title to change to "0 reminders' if there are no other real outstanding reminders 
            IntPtr hWndChild = FindChildWindowOfClassSysListView(hWndMainWindow);
            if (hWndChild != null)
            {
                // post a message to SysListView, and if there are no events, it will change the title to 0 reminders which we ignore                SendMessage(hWndChild, WM_LBUTTONDOWN, 0, 0);
                SendMessage(hWndChild, WM_LBUTTONUP, 0, 0);
                System.Threading.Thread.Sleep(5000);
            }
        }

        private static IntPtr FindChildWindowOfClassSysListView(IntPtr hWndParent)
        {
            IntPtr hwndChild = IntPtr.Zero;
            StringBuilder sbClassName = new StringBuilder(512);
            EnumChildWindows(hWndParent, (hwnd, param) =>
            {
                sbClassName = new StringBuilder(512);
                if (GetClassName(hwnd, sbClassName, sbClassName.Capacity) > 0)
                {

                    if (sbClassName.ToString().Contains("SysListView"))
                    {
                        Debug.WriteLine($"BINGO: {hwnd} {sbClassName} ");
                        hwndChild = hwnd;
                        return 0;
                    }
                }
                Debug.WriteLine($"{hwnd} {sbClassName} ");
                return 1;
            }, IntPtr.Zero);
            return hwndChild;
        }
    }
}
