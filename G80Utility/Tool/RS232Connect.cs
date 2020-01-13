using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PirnterUtility.Tool
{
    class RS232Connect
    {
        public static SerialPort mSerialPort = new SerialPort();
        public static bool isReceiveData;
        public static byte[] mRecevieData;
        public static int mLength;
        public static bool IsConnect;

        //開啟SerialPort
        public static bool OpenSerialPort(string comPort, string msg)
        {
            bool isError = false;

            //Sets up serial port
            mSerialPort.PortName = comPort;
            mSerialPort.BaudRate = (int)App.Current.Properties["BaudRateSetting"];
            mSerialPort.Parity = Parity.None;
            mSerialPort.DataBits = 8;
            mSerialPort.StopBits = StopBits.One;
            mSerialPort.Handshake = Handshake.RequestToSend;
            mSerialPort.ReadTimeout = 100000;
            mSerialPort.WriteTimeout = 5000;
           
            //不加下面兩行就會無法開啟serialport
            mSerialPort.DtrEnable = true;
            mSerialPort.RtsEnable = true;

            try
            {
                mSerialPort.Open();
                //MessageBox.Show("open port");

            }
            catch (Exception ex) //開啟comport時要用try catch接起來以免遇到comport無法開啟出錯
            {
                isError = true;

                if (ex.ToString().Contains("UnauthorizedAccessException"))
                {
                    MessageBox.Show(msg);
                }
                else
                {
                    MessageBox.Show(ex.ToString());
                }
                mSerialPort.Close();
            }
            return isError;
        }

        //關閉SerialPort
        public static void CloseSerialPort()
        {
            if (mSerialPort != null && mSerialPort.IsOpen)
            {
                mSerialPort.Close();
            }
        }

        //透過RS232傳送CMD
        public static bool SerialPortSendCMD(string cmdType, byte[] data, string msg,int recevieLength)
        {
            //將isReceiveData和mRecevieData恢復預設
            isReceiveData = false;
            mRecevieData = null;
           

            if (mSerialPort.IsOpen)
            {
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

                    switch (cmdType) {

                        case "NeedReceive":

                            Task.Factory.StartNew(() =>
                            {
                                ReceiveInfo(recevieLength); //通訊測試收到資料44長度                       
                            });
                            break;
                        case "NoReceive":
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    if (ex.ToString().Contains("TimeOut") || ex.ToString().Contains("IO"))
                    {
                        MessageBox.Show("Failed to SEND:" + "\n" + "Comport setting error!");
                    }
                    else
                    {
                        MessageBox.Show("Failed to SEND:" + "\n" + ex + "\n");
                    }
                    IsConnect = false;
                    mSerialPort.Close();
                    isReceiveData = true;//如果不設為true前端會一直收資料
                }
            }
            else //mSerialPort close
            {
                MessageBox.Show(msg);
                isReceiveData = true;//如果不設為true前端會一直收資料
                IsConnect = false;
            }
            return isReceiveData;
        }
  

        //接收回復內容
        public static void ReceiveInfo(int count)
        {
            int offset = 0;
            byte[] buffer = new byte[count];
            while (count > 0)
            {
                int bytesRead = mSerialPort.Read(buffer, offset, count);
                offset += bytesRead;
                count -= bytesRead;
            }

            //string receiveData = null;
            
            ////(0~7)前8個是無意義資料
            //for (int i = 8; i < 18 ; i++)
            //{
            //    receiveData += Convert.ToChar(buffer[i]);    //機器型號
            //}

            //receiveData = null;
            //for (int i = 18; i < 28; i++)
            //{
            //    receiveData += Convert.ToChar(buffer[i]);   //軟件版本    
            //}

            //receiveData = null;
            //for (int i = 28; i < 44; i++)
            //{
            //    receiveData += Convert.ToChar(buffer[i]);      //機器序列號
            //}


            //Console.WriteLine("read:" + Encoding.Default.GetString(buffer));
            mRecevieData = buffer;
            IsConnect = true;
        }

        //接收設定回復結果
        public static void ReceiveSettingResult(string msg) {
            string recieved_data = mSerialPort.ReadLine().Trim(); 
            if (recieved_data.Contains("OK"))
            {
                MessageBox.Show(msg);
            }
        }
    }
}
