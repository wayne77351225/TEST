using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace G80Utility.Tool
{
    class RS232Connect
    {
        public static SerialPort mSerialPort = new SerialPort();
        public static bool isReceiveData;
        public static byte[] mRecevieData;
        public static bool IsConnect;

        #region 開啟SerialPort
        public static bool OpenSerialPort(string comPort, string msg)
        {
            bool isError = false;

            //Sets up serial port
            mSerialPort.PortName = comPort;
            Console.WriteLine("class:"+ (int)App.Current.Properties["BaudRateSetting"]);
            mSerialPort.BaudRate = (int)App.Current.Properties["BaudRateSetting"];
            mSerialPort.Parity = Parity.None;
            mSerialPort.DataBits = 8;
            mSerialPort.StopBits = StopBits.One;
            mSerialPort.Handshake = Handshake.RequestToSend;
            mSerialPort.ReadTimeout = 500;
            mSerialPort.WriteTimeout = 500;
            //不加下面兩行就會無法開啟serialport
            mSerialPort.DtrEnable = true;
            mSerialPort.RtsEnable = true;
            try
            {
                mSerialPort.Open();
                mSerialPort.DiscardOutBuffer();
                mSerialPort.DiscardInBuffer();
            }
            catch (Exception) 
            {
                isError = true;
                mSerialPort.Close();
            }
            return isError;
        }
        #endregion

        #region 透過RS232傳送CMD
        public static bool SerialPortSendCMD(string cmdType, byte[] data, string msg, int recevieLength)
        {
            //將isReceiveData和mRecevieData恢復預設
            isReceiveData = false;
            mRecevieData = null;

            if (data == null)
                return isReceiveData;
            List<byte> buffer = new List<byte>();
            foreach (byte bytedata in data)
            {
                buffer.Add(bytedata);
            }
            byte[] dataSend = buffer.ToArray();
            try
            {
                IsConnect = true;
                mSerialPort.Write(dataSend, 0, dataSend.Length);

                switch (cmdType)
                {

                    case "NeedReceive":

                        Task.Factory.StartNew(() =>
                        {
                            ReceiveInfo(recevieLength);
                        });
                        break;
                    case "NoReceive":
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                IsConnect = false;
                mSerialPort.Close();
                isReceiveData = true;//如果不設為true前端會一直收資料
            }
            return isReceiveData;
        }
        #endregion

        #region 接收回復內容
        public static void ReceiveInfo(int count)
        {
            int offset = 0;
            byte[] buffer = new byte[count];

            while (count > 0)
            {
                try
                {
                    int bytesRead = mSerialPort.Read(buffer, offset, count);
                    offset += bytesRead;
                    count -= bytesRead;
                    IsConnect = true;
                }
                catch (Exception)
                {
                    IsConnect = false;
                    isReceiveData = true;
                    break;
                }
            }
            //Console.WriteLine("COM接收資料" + BitConverter.ToString(buffer));
            mRecevieData = buffer;
            isReceiveData = true;
        }
        #endregion

        #region 關閉SerialPort
        public static void CloseSerialPort()
        {
            if (mSerialPort != null && mSerialPort.IsOpen)
            {
                mSerialPort.Close();
            }
        }
        #endregion

        #region 判斷SerailPort連現狀態
        public static bool RS232ConnectStatus()
        {
            if (mSerialPort != null && mSerialPort.IsOpen)
            {
                return true;
            }
            else {

                return false;
            }
        }
        #endregion
    }
}
