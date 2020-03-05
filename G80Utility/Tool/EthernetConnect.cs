using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace G80Utility.Tool
{
    class EthernetConnect
    {
        //Ethernetnet 設定
        public static Socket SocketClient;
        public static readonly ManualResetEvent TimeoutObject = new ManualResetEvent(false);
        public static Thread ThreadClient = null;
        public static int Receive_Size;
        public static string EthernetIPAddress;
        public static bool isReceiveData;
        public static byte[] mRecevieData;
        public static int connectStatus;
        //判斷如果ip改變要重新連線，因為目前是一個連線可以重複讀取
        public static string EthernetIPAddressOld;

        #region socket非同步回撥方法
        private static void CallBackMethod(IAsyncResult asyncresult)
        {
            //使阻塞的執行緒繼續        
            TimeoutObject.Set();
        }
        #endregion

        #region socket連線印表機
        public static int connectToPrinter()
        {
            TimeoutObject.Reset();
            int port = 9100;
            int connectCode = 0; //0:連線失敗,1:連線成功,2:連線逾時
            try
            {
                IPAddress ip = IPAddress.Parse(EthernetIPAddress);
                IPEndPoint ipe = new IPEndPoint(ip, port);
                SocketClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                IAsyncResult ConnectResult = SocketClient.BeginConnect(ipe, CallBackMethod, SocketClient);
                bool success = ConnectResult.AsyncWaitHandle.WaitOne(3000, true);
                EthernetIPAddressOld = EthernetIPAddress;
                if (success)
                {
                    connectCode = 1;
                }
                else
                {
                    connectCode = 2;
                    disconnect();
                }
            }
            catch (Exception e)
            {
                connectCode = 0;
            }
            return connectCode;
        }
        #endregion

        #region 傳送狀態檢查
        public static int EthernetConnectStatus() //測試連線狀態，避免重複連線
        {
            if (connectStatus != 1 || EthernetIPAddressOld != EthernetIPAddress)
            {
                connectStatus = connectToPrinter();　//非連線狀態要先連線
            }
            return connectStatus;
        }
        #endregion

        #region 傳送命令
        public static bool EthernetSendCmd(string cmdType, byte[] data, string msg, int recevieLength)
        {
            //將isReceiveData和mRecevieData恢復預設
            isReceiveData = false;
            mRecevieData = null;
            Receive_Size = recevieLength;

            if (TimeoutObject.WaitOne(2000, false) && connectStatus == 1)
            {
                ((G80MainWindow)Application.Current.MainWindow).EthernetConnectImage.Source = new BitmapImage(new Uri("Images/green_circle.png", UriKind.Relative)); //連線失敗時
                switch (cmdType)
                {
                    case "NeedReceive":
                        ThreadClient = new Thread(EthernetReceive);
                        ThreadClient.IsBackground = true;
                        ThreadClient.Start();
                        SocketClient.Send(data);
                        break;
                    case "NoReceive":
                        SocketClient.Send(data);
                        // disconnect(); //寫入命令才關閉socket，避免寫入後打印機回傳其他數值造成同一個socket在接收的時候出現問題
                        break;
                }
            }
            else
            {
                connectStatus = 2;
                isReceiveData = true;
                disconnect();
            }
            return isReceiveData;
        }
        #endregion

        #region 接收資料
        public static void EthernetReceive()
        {
            try
            {
                byte[] buffer = new byte[Receive_Size];
                //將客戶端套接字接收到的資料存入記憶體緩衝區，並獲取長度  
                int length = SocketClient.Receive(buffer);
                mRecevieData = buffer;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\r\n");
                isReceiveData = true;
                connectStatus = 2;
                disconnect();

            }

        }
        #endregion

        #region 關閉socket連線
        public static void disconnect()
        {
            if (ThreadClient != null && ThreadClient.IsAlive)
            {
                try
                {
                    ThreadClient.Abort();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
            }
            try
            {
                SocketClient.Close();
                connectStatus = 0;
            }
            catch (SocketException e)
            {
                connectStatus = 0;
                Console.WriteLine(e.ToString());
                Console.WriteLine(e.ErrorCode + "");

            }

            isReceiveData = true;//如果不設為true前端會一直收資料
        }
        #endregion
    }
}
