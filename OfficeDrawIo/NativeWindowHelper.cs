using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDrawIo
{
    internal static class NativeWindowHelper
    {
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int cmdShow);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string className, string windowName);

        public static IntPtr FindWindow(string title)
        {
            return FindWindow(null, title);
        }

        public static void RestoreFromMinimized(IntPtr hwnd)
        {
            const int SW_RESTORE = 9;
            ShowWindow(hwnd, SW_RESTORE);
        }
    }
}
