using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;

namespace G80Utility.Tool
{
    class USBConnect
    {//public static string USBpath;
        public static bool isReceiveData;
        public static byte[] mRecevieData;

        //設定IntPtr的參數
        public static int USBHandle = -1;
        public static int usb_flag = 0;

        //CreateFile参数配置
        public const uint GENERIC_READ = 0x80000000;
        public const uint GENERIC_WRITE = 0x40000000;
        public const uint FILE_SHARE_READ = 0x00000001;
        public const uint FILE_SHARE_WRITE = 0x00000002;
        public const int OPEN_EXISTING = 3;
        public const int OVERLAPPED = 0x40000000;
        public const int NO_OVERLAPPED = 0;

        //usb 寫入pending error code
        public static int ERROR_IO_PENDING = 997;

        public static bool IsConnect;

        #region 連接usb設備
        public static int ConnectUSBDevice(string DeviceName)
        {
            if (USBHandle == -1) //確認連線斷線後才開啟新連線
            { 
                USBHandle = Kernel32.CreateFile
               (
                   DeviceName,
                   //讀或寫
                   GENERIC_READ | GENERIC_WRITE,
                   //共享讀寫
                   FILE_SHARE_READ | FILE_SHARE_WRITE,
                   0,
                   OPEN_EXISTING,
                   OVERLAPPED,
                   0
                );
                //Console.WriteLine(" IN_DeviceName = " + DeviceName);                    //查看参数是否传入           
                if (USBHandle == -1)
                {
                    Console.WriteLine(" 失败 HidHandle = 0x" + "{0:x}", USBHandle);
                    return 0;
                }
                else    //连接成功
                {
                    Console.WriteLine(" 成功 HidHandle = 0x" + "{0:x}", USBHandle);
                    return 1;
                }
            }
            else {
                Console.WriteLine("USB已經連線 HidHandle = 0x" + "{0:x}", USBHandle);
                return 1;
            }        
        }
        #endregion

        #region Send Command
        public static bool USBSendCMD(string cmdType, byte[] data, string msg, int recevieLength)
        {
            //將isReceiveData和mRecevieData恢復預設
            isReceiveData = false;
            mRecevieData = null;

            Kernel32.OVERLAPPED overlap = new Kernel32.OVERLAPPED();
            overlap.Offset = 0;
            overlap.OffsetHigh = 0;

            uint write = 0;
            bool isread = Kernel32.WriteFile((IntPtr)USBHandle, data, (uint)(data.Length), ref write, ref overlap);
            if (!isread && Marshal.GetLastWin32Error() == ERROR_IO_PENDING)
            {
                Kernel32.GetOverlappedResult((IntPtr)USBHandle, ref overlap, ref write, true);
            }
            switch (cmdType)
            {
                case "NeedReceive":
                    Task.Factory.StartNew(() =>
                    {
                        USBReceiveData(recevieLength);
                    });
                    break;
                case "NoReceive":
                    break;
            }
            return isReceiveData;
        }
        #endregion



        #region 取得資料
        public static void USBReceiveData(int size)
        {
            byte[] buffer = new byte[size];
            uint dwRead = 0;
            int nReadLen = 0;

            Kernel32.OVERLAPPED overlap = new Kernel32.OVERLAPPED();
            overlap.Offset = 0;
            overlap.OffsetHigh = 0;
  
            int elapsed = 0;

            while (nReadLen < size)
            {
                elapsed += 1000;
                bool re = Kernel32.ReadFile((IntPtr)USBHandle, buffer, (uint)size, ref dwRead, ref overlap);

                if (!re && Marshal.GetLastWin32Error() == ERROR_IO_PENDING)
                {
                    if (elapsed == 3000) //3Secs timeout
                    {
                        isReceiveData = true;
                        break;
                    }
                    Kernel32.GetOverlappedResult((IntPtr)USBHandle, ref overlap, ref dwRead, true);
                }
                nReadLen += (int)dwRead;
                IsConnect = true;
                if (elapsed == 3000) //3Secs timeout
                {
                    isReceiveData = true;
                    break;
                }
            }

            if (nReadLen < size) //timeout 後回覆狀態
            {
                IsConnect = false;
                isReceiveData = true;
                closeHandle(); //timeout後關閉USBHandle
            }
            mRecevieData = buffer;
            Console.WriteLine("USB接收資料" + BitConverter.ToString(buffer));
            isReceiveData = true;
        }
        #endregion

        #region 關閉Handle
        public static void closeHandle()
        {
            USBHandle = -1;
            Kernel32.CloseHandle(USBHandle);
            GC.Collect();
        }
        #endregion
    }
}
