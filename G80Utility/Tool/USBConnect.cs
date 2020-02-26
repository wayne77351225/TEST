using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace G80Utility.Tool
{
    class USBConnect
    {//public static string USBpath;
        public static bool isCMDPass;
        public static int receiveTimes;
        public static bool isReceiveData;
        public static byte[] mRecevieData;
        public static int mLength;

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

        //public static bool isTimeout;
        public static bool isTimeout;
        public static bool isSettingOK;

        //wifi setting result
        public static string WifiSettingResult;


        #region 連接usb設備
        public static int ConnectUSBDevice(string DeviceName)
        {
            USBHandle = Kernel32.CreateFile
       (
           DeviceName,
           //GENERIC_READ |          // | GENERIC_WRITE,//读写，或者一起
           GENERIC_READ | GENERIC_WRITE,
           //FILE_SHARE_READ |       // | FILE_SHARE_WRITE,//共享读写，或者一起
           FILE_SHARE_READ | FILE_SHARE_WRITE,
           0,
           OPEN_EXISTING,
           OVERLAPPED,
           0
        );

            Console.WriteLine(" IN_DeviceName = " + DeviceName);                    //查看参数是否传入           

            if (USBHandle == -1) //INVALID_HANDLE_VALUE实际值等于-1，连接失败
            {

                Console.WriteLine(" 失败 HidHandle = 0x" + "{0:x}", USBHandle);     //查看状态，打印调试用
                return 0;
            }
            else    //连接成功
            {
                Console.WriteLine(" 成功 HidHandle = 0x" + "{0:x}", USBHandle);      //查看状态，打印调试用

                return 1;
            }
        }
        #endregion

        #region Send Command
        public static bool USBSendCMD(string cmdType, byte[] data, string msg, int recevieLength)
        {
            //將isReceiveData和mRecevieData和isSettingOK恢復預設
            isReceiveData = false;
            mRecevieData = null;
            isSettingOK = false;
            isTimeout = false;
            Kernel32.OVERLAPPED overlap = new Kernel32.OVERLAPPED();
            overlap.Offset = 0;
            overlap.OffsetHigh = 0;

            uint write = 0;
            bool isread = false;
            isread = Kernel32.WriteFile((IntPtr)USBHandle, data, (uint)(data.Length), ref write, ref overlap);
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
                    closeHandle();
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
            //初始化Overlapped
            Kernel32.OVERLAPPED overlap = new Kernel32.OVERLAPPED();
            overlap.Offset = 0;
            overlap.OffsetHigh = 0;
            //创建事件对象
            overlap.hEvent = Kernel32.CreateEvent(IntPtr.Zero, Convert.ToInt32(false), Convert.ToInt32(false), null);
            long waitResutl = 0;
            //讀取設備
            while (nReadLen < size)
            {
                bool re = Kernel32.ReadFile((IntPtr)USBHandle, buffer, (uint)size, ref dwRead, ref overlap);
                Console.WriteLine("原始資料" + BitConverter.ToString(buffer));
                if (!re && Marshal.GetLastWin32Error() == ERROR_IO_PENDING)
                {
                    //waitResutl = Kernel32.WaitForSingleObject(overlap.hEvent, 3000); //關閉timeout，連續傳多個命令開啟timout會收不到

                    //if (waitResutl == 258 || waitResutl == 4294967295)
                    //{
                    //    Console.WriteLine("data0..." + "timeout");
                    //    isTimeout = true;
                    //    isReceiveData = true;
                    //    break;
                    //}
                    Kernel32.GetOverlappedResult((IntPtr)USBHandle, ref overlap, ref dwRead, true);
                    Console.WriteLine("readgn..." + dwRead);
                }
                //if (waitResutl == 258 || waitResutl == 4294967295)
                //{
                //    Console.WriteLine("data..." + "timeout");
                //    isTimeout = true;
                //    isReceiveData = true;
                //    break;
                //}
                nReadLen += (int)dwRead;
                //isTimeout = false;
            }
          
            closeHandle();

            //if (dwRead == size)
            //{
                Console.WriteLine("接收資料" + BitConverter.ToString(buffer));
                mRecevieData = buffer;
            //}
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
