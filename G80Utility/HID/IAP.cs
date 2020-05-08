﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.IO.Ports;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace G80Utility.HID
{
    class IAP_download
    {


        #region parameter Define

        HIDInterface hid = new HIDInterface();

        //private Form main_win;
        public byte run_step = 0;
        bool convert_bin_done = false;
        public string hex_file_name = null;
        UInt32 download_addr;
        byte[] code_data_bin;
        private static ManualResetEvent TimeoutObject = new ManualResetEvent(false);

        byte[] receive_data;

        struct connectStatusStruct
        {
            public bool preStatus;
            public bool curStatus;
        }

        connectStatusStruct connectStatus = new connectStatusStruct();

        //推送连接状态信息
        public delegate void isConnectedDelegate(bool isConnected);
        public isConnectedDelegate isConnectedFunc;


        //推送接收数据信息
        public delegate void PushReceiveDataDele(byte[] datas);
        public PushReceiveDataDele pushReceiveData;

        #endregion

        //第一步需要初始化，传入vid、pid，并开启自动连接
        public void Initial()
        {

            hid.StatusConnected = StatusConnected;
            hid.DataReceived = DataReceived;

            HIDInterface.HidDevice hidDevice = new HIDInterface.HidDevice();
            hidDevice.vID = 0x28E9; //固定寫死
            hidDevice.pID = 0x028B; //固定寫死
            hidDevice.serial = "";
            hid.AutoConnect(hidDevice);
            pushReceiveData += read_data;
        }

        //不使用则关闭
        public void Close()
        {
            hid.StopAutoConnect();
        }

        //发送数据
        public bool SendBytes(byte[] data)
        {
            return hid.Send(data);
        }

        public void read_data(byte[] data)
        {
            receive_data = data;
            TimeoutObject.Set();
        }

        //接受到数据 event
        public void DataReceived(object sender, byte[] e)
        {
            if (pushReceiveData != null)
            {
                pushReceiveData(e);
            }
        }

        //状态改变 event
        public void StatusConnected(object sender, bool isConnect)
        {
            connectStatus.curStatus = isConnect;
            if (connectStatus.curStatus == connectStatus.preStatus)  //connect
                return;
            connectStatus.preStatus = connectStatus.curStatus;

            if (connectStatus.curStatus)
            {
                isConnectedFunc(true);
            }
            else //disconnect
            {
                isConnectedFunc(false);
                Close();
            }
        }



        public void hex_file_to_bin_array(object sender)
        {
            this.run_step = 1;
            this.convert_bin_done = false;
            G80MainWindow win = (G80MainWindow)sender;
            int line_num = 0, current_line = 0;
            string szLine;
            string szHex = "";
            StreamReader HexReader;
            StreamReader cal_line_num;
            if (this.hex_file_name == null)
            {
                win.Dispatcher.Invoke(win.setCallBack, (byte)1, win.FindResource("SelectFileFirst") as string);
                return;
            }

            try
            {
                HexReader = new StreamReader(this.hex_file_name);
                cal_line_num = new StreamReader(this.hex_file_name);
            }
            catch
            {
                EventArgs ex = new EventArgs();
                win.Dispatcher.Invoke(win.setCallBack, (byte)1, win.FindResource("FileError") as string);
                return;
            }
            win.Dispatcher.Invoke(win.setCallBack, (byte)1, win.FindResource("Parsing") as string);
            do
            {
                line_num++;
            }
            while (null != cal_line_num.ReadLine());

            /* 第一行数据是 */
            StringBuilder file_start_addr = new StringBuilder();

            szLine = HexReader.ReadLine();
            if (szLine.Substring(0, 3) == ":02")//数据结束
            {
                file_start_addr.Append(szLine.Substring(9, szLine.Length - 11));
            }

            szLine = HexReader.ReadLine();
            if (szLine.Substring(0, 3) == ":10")//数据结束
            {
                file_start_addr.Append(szLine.Substring(3, 4));
            }
            win.Dispatcher.Invoke(win.setCallBack, (byte)3, "0X" + file_start_addr.ToString());
            try
            {
                this.download_addr = (UInt32)Convert.ToUInt32(file_start_addr.ToString().Substring(0, 8), 16);
                int count = 16;
                current_line = 2;
                int baud = 0;
                while ((szLine != null) && (szLine.Substring(0, 9) != ":00000001"))
                {
                    if (szLine.Substring(0, 1) == ":") //判断第1字符是否是:
                    {
                        if (szLine.Substring(1, 1) == "1")
                        {
                            szHex += szLine.Substring(9, szLine.Length - 11); //读取有效字符：后0和1
                        }
                    }
                    szLine = HexReader.ReadLine(); //读取一行数据
                    count += szLine.Length - 11;
                    current_line++;

                    if (baud != current_line * 100 / line_num)
                    {
                        baud = current_line * 100 / line_num;
                        win.Dispatcher.Invoke(win.setCallBack, (byte)4, baud.ToString());
                    }
                }
                win.Dispatcher.Invoke(win.setCallBack, (byte)4, "100");

                Int32 Length = Encoding.Default.GetByteCount(szHex);
                code_data_bin = new byte[Length / 2];
                for (Int32 i = 0; i < code_data_bin.Length; i += 1) //两字符合并成一个16进制字节
                {
                    code_data_bin[i] = Convert.ToByte(szHex.Substring(i * 2, 2), 16);
                    if (baud != i * 100 / code_data_bin.Length)
                    {
                        baud = i * 100 / code_data_bin.Length;
                        win.Dispatcher.Invoke(win.setCallBack, (byte)4, baud.ToString());
                    }
                }
                win.Dispatcher.Invoke(win.setCallBack, (byte)4, "100");
                UInt32 code_size = (UInt32)(code_data_bin.Length / 1024);
                win.Dispatcher.Invoke(win.setCallBack, (byte)2, code_size.ToString() + "KB");
                win.Dispatcher.Invoke(win.setCallBack, (byte)1, win.FindResource("ParseCompleted") as string);
                this.run_step = 0;
                this.convert_bin_done = true;
            }
            catch (Exception )
            {
                this.run_step = 0;//恢復default
                win.Dispatcher.Invoke(win.setCallBack, (byte)1, "");//恢復default
                win.Dispatcher.Invoke(win.setCallBack, (byte)7, win.FindResource("CanNotParse") as string);
            }
        }


        private bool read_opt_byte(out byte[] option)
        {
            option = null;

            TimeoutObject.Reset();
            byte[] send = new byte[] { 0x01 };
            SendBytes(send);

            if (TimeoutObject.WaitOne(1000, false))
            {
                option = receive_data;
                return true;
            }
            return false;
        }

        private bool get_erasure_addr(out byte[] option)
        {
            option = null;
            TimeoutObject.Reset();
            byte[] send = new byte[] { 0x05 };
            SendBytes(send);
            if (TimeoutObject.WaitOne(1000, false))
            {
                option = receive_data;
                return true;
            }
            return false;
        }

        private bool erasure_section(UInt32 Start_addr, UInt32 file_length)
        {

            byte[] send_data = new byte[] { 0x02, 0x08, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
            send_data[1] = (byte)(Start_addr >> 0);
            send_data[2] = (byte)(Start_addr >> 8);
            send_data[3] = (byte)(Start_addr >> 16);
            send_data[4] = (byte)(Start_addr >> 24);

            send_data[6] = (byte)(file_length >> 0);
            send_data[7] = (byte)(file_length >> 8);
            send_data[8] = (byte)(file_length >> 16);
            send_data[9] = (byte)(file_length >> 24);

            TimeoutObject.Reset();

            SendBytes(send_data);

            if (TimeoutObject.WaitOne(10000, false))
            {
                if (receive_data[0] == 0x01)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return false;
        }
        private bool download_code(byte[] code_array, G80MainWindow win1)
        {
            G80MainWindow win = (G80MainWindow)win1;
            byte[] send_data = new byte[63];

            send_data[0] = 0x03;
            int baud = 0;
            for (int offset = 0; offset < code_array.Length;)
            {
                TimeoutObject.Reset();
                if (code_array.Length - offset > 62)
                {
                    Array.Copy(code_array, offset, send_data, 1, 62);
                }
                else
                {
                    Array.Copy(code_array, offset, send_data, 1, code_array.Length - offset);
                }

                offset += 62;
                SendBytes(send_data);
                if (baud != offset * 100 / code_array.Length)
                {
                    baud = offset * 100 / code_array.Length;
                    win.Dispatcher.Invoke(win.setCallBack, (byte)6, baud.ToString());
                }
            }

            if (TimeoutObject.WaitOne(1000, false))
            {
                if (receive_data[0] == 0x02)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return false;
        }

        private bool GD32HIDIAP_LeaveIAP()
        {
            byte[] send_data = new byte[] { 0x04 };
            SendBytes(send_data);
            return true;
        }

        public void download_code(object sender)
        {


            G80MainWindow win = (G80MainWindow)sender;
            if (!this.convert_bin_done)
            {
                MessageBox.Show(win.FindResource("ParseFirst") as string);
                return;
            }
            this.run_step = 2;
            win.Dispatcher.Invoke(win.setCallBack, (byte)5, win.FindResource("ReadChipOption") as string);
            byte[] receive_data = new byte[64];
            if (!read_opt_byte(out receive_data))
            {
                MessageBox.Show(win.FindResource("FailedReadChipOption") as string);
                return;
            }
            if (receive_data[1] != 0xaa)
            {
                MessageBox.Show(win.FindResource("ChipIsLocked") as string);
                return;
            }
            win.Dispatcher.Invoke(win.setCallBack, (byte)5, win.FindResource("ReadChipAddress") as string);
            if (!get_erasure_addr(out receive_data))
            {
                MessageBox.Show(win.FindResource("EraseAddressFailed") as string);
                return;
            }
            UInt32 write_addr = (UInt32)receive_data[0] << 0 | (UInt32)receive_data[1] << 8 | (UInt32)receive_data[2] << 16 | (UInt32)receive_data[3] << 24;
            if (this.download_addr != write_addr)
            {
                MessageBox.Show(win.FindResource("AddressError") as string);
                return;
            }
            win.Dispatcher.Invoke(win.setCallBack, (byte)5, win.FindResource("CheckChipSector") as string);
            if (!erasure_section(write_addr, (UInt32)code_data_bin.Length))
            {
                MessageBox.Show(win.FindResource("EraseMemoryError") as string);
                return;
            }
            win.Dispatcher.Invoke(win.setCallBack, (byte)5, win.FindResource("SendingData") as string);
            if (!download_code(code_data_bin, win))
            {
                MessageBox.Show(win.FindResource("BurningFailed") as string);
                return;
            }
            GD32HIDIAP_LeaveIAP();
            win.Dispatcher.Invoke(win.setCallBack, (byte)5, win.FindResource("UpdateToROM") as string);

            this.run_step = 0;
        }
    }
}
