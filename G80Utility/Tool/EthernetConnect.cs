using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows;
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


        #region socket非同步回撥方法
        private static void CallBackMethod(IAsyncResult asyncresult)
        {
            //使阻塞的執行緒繼續        
            TimeoutObject.Set();
        }
        #endregion

        #region socket連線印表機
        public static bool connectToPrinter()
        {
            TimeoutObject.Reset();
            int port = 9100;
            bool isConnect;

            if (EthernetIPAddress != null)
            {
                try
                {
                    IPAddress ip = IPAddress.Parse(EthernetIPAddress);
                    IPEndPoint ipe = new IPEndPoint(ip, port);
                    SocketClient = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                    SocketClient.BeginConnect(ipe, CallBackMethod, SocketClient);
                    isConnect = true;
                }
                catch (Exception e)
                {
                    isConnect = false;
                    Console.WriteLine("連線失敗！\r\n" + e.ToString());
                    ((G80MainWindow)Application.Current.MainWindow).EthernetConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //連線失敗時
                    MessageBox.Show(Application.Current.FindResource("NotSettingEthernetport") as string);              
                    disconnect();
                }
            }
            else
            {
                isConnect = false;
                ((G80MainWindow)Application.Current.MainWindow).EthernetConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //連線失敗時
                MessageBox.Show(Application.Current.FindResource("NotSettingEthernetport") as string);
            }
            return isConnect;
        }
        #endregion

        #region 關閉socket連線
        public static void disconnect()
        {
            if (ThreadClient.IsAlive)
            {
                ThreadClient.Abort();
            }
            SocketClient.Close();
            isReceiveData = true;//如果不設為true前端會一直收資料
        }
        #endregion

        #region 傳送命令
        public static bool EthernetSendCmd(string cmdType, byte[] data, string msg, int recevieLength)
        {
            //將isReceiveData和mRecevieData恢復預設
            isReceiveData = false;
            mRecevieData = null;
            Receive_Size = recevieLength;

            if (TimeoutObject.WaitOne(2000, false))
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
                        break;
                }
            }
            else
            {
                ((G80MainWindow)Application.Current.MainWindow).EthernetConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //連線失敗時
                MessageBox.Show(Application.Current.FindResource("ConnectTimeout") as string);
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
            int ReadLen = 0;

        }
        catch (Exception ex)
        {
            //這邊會收到中斷執行緒的exception所以不特別做ui處理
            Console.WriteLine(ex.Message + "\r\n");
            isReceiveData = true;
        }
    }
    #endregion
}
}
