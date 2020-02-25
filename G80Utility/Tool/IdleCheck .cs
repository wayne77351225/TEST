using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace G80Utility.Tool
{
    class IdleCheck
    {
        [StructLayout(LayoutKind.Sequential)]

        private struct LASTINPUTINFO
        {

            [MarshalAs(UnmanagedType.U4)]

            public int cbSize;

            [MarshalAs(UnmanagedType.U4)]

            public int dwTime; //取得系統最後一次操作時間

        }

        [DllImport("user32.dll")]

        private static extern bool GetLastInputInfo(ref LASTINPUTINFO x);

        // 把最後一次操作滑鼠或鍵盤的時間寫入dwTime
        public static TimeSpan GetLastInputTime()
        {

            var inf = new LASTINPUTINFO();

            inf.cbSize = Marshal.SizeOf(inf);

            inf.dwTime = 0;
            //Environment.TickCount，取得系統後經過的毫秒數
            return TimeSpan.FromMilliseconds((GetLastInputInfo(ref inf)) ? Environment.TickCount - inf.dwTime : 0);

        }

    }
}
