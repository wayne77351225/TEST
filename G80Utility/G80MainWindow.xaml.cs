﻿using G80Utility.Tool;
using Microsoft.Win32;
using PirnterUtility.Models;
using PirnterUtility.Tool;
using PirnterUtility.ViewModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace PirnterUtility
{
    /// <summary>
    /// Window1.xaml 的互動邏輯
    /// </summary>
    public partial class G80MainWindow : Window
    {
        #region 定義變數
        //取得usb註冊碼
        RegistryKey rkLocalMachine = Registry.LocalMachine;

        //device path
        string USBpath;
        string RS232PortName;

        //作業系統版本
        string OSVersion;

        //傳輸通道類別
        string DeviceType;

        //device 清單
        List<Device> deviceList = new List<Device>();

        //property 
        DeviceViewModel viewmodel { get; set; }

        //SerialPort
        string recieved_data;



        #endregion

        public G80MainWindow()
        {
            InitializeComponent();

            //語系選單default設定
            setDefaultLanguage();

            // 取得作業系統版本
            OperatingSystem os = Environment.OSVersion;
            OSVersion = os.Version.Major.ToString() + "." + os.Version.Minor.ToString();

            UIInitial();

            //頁面內容產生後註冊usb device plugin notify
            this.ContentRendered += WindowThd_ContentRendered;

        }

        private void WindowThd_ContentRendered(object sender, EventArgs e)
        {
            registerUSBdetect();
        }


        #region UI初始化
        private void UIInitial()
        {

            //combobox databinding
            viewmodel = new DeviceViewModel();
            DataContext = viewmodel;

            //BaudRate預設
            BaudRateCom.SelectedIndex = 2;
            App.Current.Properties["BaudRateSetting"] = 38400;
            
            //default 為選取RS232
            this.RS232Radio.IsChecked = true;

            //combobox設定item
            //Density
            List<int> DensityList = new List<int>();
            for (int i = 1; i <= 10; i++)
            {
                DensityList.Add(i);
            }
            DensityCom.ItemsSource = DensityList;
            //CodePage
            CodePageCom.ItemsSource = CodePage.getCodePageList();

        }
        #endregion

        //========================Btn點擊事件===========================
        #region 通讯接口测试按鈕事件
        private void ConnectTest_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.RS232_COMMUNICATION_TEST);
            if ((bool)rs232Checkbox.IsChecked)
            {
                SerialPortConnect("CommunicationTest", sendArray);
            }

        }
        #endregion

        #region 重啟印表機按鈕事件
        private void Restart_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.RESTART);
            switch (DeviceType)
            {

                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }

        }
        #endregion

        #region 設定語言按鈕事件
        private void LanguageSetBtn_Click(object sender, RoutedEventArgs e)
        {

            byte[] sendArray = null;

            switch (LanguageSetCom.SelectedIndex)
            {
                case 0:
                    sendArray = StringToByteArray(Command.LANGUAGE_SETTING_GB18030);
                    break;
                case 1:
                    sendArray = StringToByteArray(Command.LANGUAGE_SETTING_BIG5);
                    break;
                case 2:
                    sendArray = StringToByteArray(Command.LANGUAGE_SETTING_KOREAN);
                    break;
                case 3:
                    sendArray = StringToByteArray(Command.LANGUAGE_SETTING_JAPANESE);
                    break;
                case 4:
                    sendArray = StringToByteArray(Command.LANGUAGE_SETTING_SHIFT_JIS);
                    break;
                case 5:
                    sendArray = StringToByteArray(Command.LANGUAGE_SETTING_JIS);
                    break;
            }
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }
        }
        #endregion

        #region FontB設定按鈕事件
        private void FontBSettingBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray;
            if (FontBSettingCom.SelectedIndex == 0)
            {
                sendArray = StringToByteArray(Command.FONTB_OFF_SETTING);
            }
            else
            {
                sendArray = StringToByteArray(Command.FONTB_ON_SETTING);
            }
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }
        }
        #endregion

        #region 走紙方向按鈕事件
        private void Direction_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray;
            if (DirectionCombox.SelectedIndex == 0)
            {
                sendArray = StringToByteArray(Command.DIRECTION_H80250N_SETTING);
            }
            else
            {
                sendArray = StringToByteArray(Command.DIRECTION_80250N_SETTING);
            }
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }
        }
        #endregion

        #region 設定馬達加速度按鈕事件
        private void AccMotorBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = null;

            switch (AccMotorCom.SelectedIndex)
            {
                case 0: //1.0=>01
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_SETTING +"01");
                    break;
                case 1: //0.8=>02
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_SETTING + "02");
                    break;
                case 2: //0.6=>03
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_SETTING + "03");
                    break;
                case 3: //0.4=>04
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_SETTING + "04");
                    break;
                case 4: //0.2=>05
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_SETTING + "05");
                    break;
            }
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }

        }
        #endregion

        #region 設定濃度調節鈕事件
        private void DensityBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = null;

            switch (DensityCom.SelectedIndex)
            {
                case 0: //1=>16
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "16");
                    break;
                case 1: //2=>61
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "61");
                    break;
                case 2: //3=>01
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "01");
                    break;
                case 3: //4=>02
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "02");
                    break;
                case 4: //5=>03
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "03");
                    break;
                case 5: //6=>04
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "04");
                    break;
                case 6: //7=>05
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "05");
                    break;
                case 7: //8=>06
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "06");
                    break;
                case 8: //9=>07
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "07");
                    break;
                case 9: //10=>08
                    sendArray = StringToByteArray(Command.DENSITY_SETTING_HEADER + "08");
                    break;
            }
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }

        }
        #endregion

        //========================取得資料後設定UI=================

        #region 設定打印機型號/軟件版本/機器序號
        private void SetPrinterInfo(byte[] buffer)
        {
            string moudle = null;
            string sfvesion = null;
            string sn = null;
            //(0~7)前8個是無意義資料
            for (int i = 8; i < 18; i++)
            {
                moudle += Convert.ToChar(buffer[i]);    //機器型號
            }
            Console.WriteLine("module:" + moudle);

            for (int i = 18; i < 28; i++)
            {
                sfvesion += Convert.ToChar(buffer[i]);   //軟件版本    
            }
            PrinterModule.Content = moudle + "  " + sfvesion;
            Console.WriteLine("VER:" + sfvesion);
            for (int i = 28; i < 44; i++)
            {
                sn += Convert.ToChar(buffer[i]);      //機器序列號
            }
            PrinterSN.Text = sn;
            Console.WriteLine("SN:" + sn);

        }
        #endregion




        //========================RS232的設定/傳送與接收===========================

        #region RS232 Port設定
        private void RS232_SelectedChnaged(object sender, SelectionChangedEventArgs e)
        {
            Device seletedItem = DeviceSelectRS232.SelectedItem as Device;
            if (seletedItem != null)
            {
                string portName = seletedItem.RS232PortName;
                if (portName != null)
                {
                    RS232PortName = portName;

                }
            }
        }
        #endregion

        #region RS232 BaudRate設定
        private void BaudRateCom_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (BaudRateCom.SelectedIndex)
            {
                case 0:
                    App.Current.Properties["BaudRateSetting"] = 9600;
                    break;
                case 1:
                    App.Current.Properties["BaudRateSetting"] = 19200;
                    break;
                case 2:
                    App.Current.Properties["BaudRateSetting"] = 38400;
                    break;
                case 3:
                    App.Current.Properties["BaudRateSetting"] = 115200;
                    break;
            }
        }
        #endregion

        #region RS232傳送資料
        private void SerialPortConnect(string dataType, byte[] data)
        {
            RS232Connect.CloseSerialPort();
            if (RS232PortName != null)
            {
                bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);

                if (!isError)
                {
                    switch (dataType)
                    {
                        case "PrinterInfo":
                            bool isReceiveData = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, 8);
                            while (!isReceiveData)
                            {

                                if (RS232Connect.mRecevieData != null)
                                {
                                    switch (dataType)
                                    {
                                        case "PrinterInfo":

                                            break;
                                    }
                                }
                            }
                            break;
                        case "CommunicationTest": //通訊測試
                            RS232Connect.SerialPortSendCMD("NeedReceive", data, null, 0);
                            RS232ConnectImage.Source = new BitmapImage(new Uri("Images/green_circle.png", UriKind.Relative));
                            RS232Connect.CloseSerialPort(); //沒立刻關閉有時會漏收命令
                            break;
                        case "BeepOrSetting":
                            RS232Connect.SerialPortSendCMD("NoReceive", data, null, 0);
                            RS232Connect.CloseSerialPort(); //沒立刻關閉有時會漏收命令
                            break;
                    }
                    if (!RS232Connect.IsConnect)
                    {
                        RS232ConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //連線失敗時
                    }

                }
                else
                {
                    RS232ConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative));
                }
            }
            else
            {
                MessageBox.Show(FindResource("NotSettingComport") as string);
                try // just in case serial port is not open could also be acheved using if(serial.IsOpen)
                {
                    RS232Connect.CloseSerialPort();
                }
                catch
                {
                }
            }
        }
        #endregion

        public static byte[] StringToByteArray(string hex)
        {
            string afterConvert = hex.Replace(" ", "");
            return Enumerable.Range(0, afterConvert.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(afterConvert.Substring(x, 2), 16))
                             .ToArray();
        }

        //=========================RS232 port search and get=====================

        #region 取得rs232 port
        private void getSerialPort()
        {
            List<String> ports = SerialPort.GetPortNames().ToList();

            for (int i = 0; i < ports.Count; i++)
            {
                for (int j = i + 1; j < ports.Count; j++)
                {
                    if (ports[i] == ports[j])
                    {
                        ports.Remove(ports[i]);
                    }
                }
                Device device = new Device() { DeviceType = "rs232", RS232PortName = ports[i] };
                deviceList.Add(device);

                viewmodel.addDevice(device);
            }
        }
        #endregion

        //===========================USB and RS232裝置插拔偵測=============================

        #region register usbdetect notify
        private void registerUSBdetect()
        {
            HwndSource hwndSource = HwndSource.FromHwnd(Process.GetCurrentProcess().MainWindowHandle);

            if (hwndSource != null)
            {
                IntPtr windowHandle = hwndSource.Handle;
                hwndSource.AddHook(UsbNotificationHandler);
                USBDetector.RegisterUsbDeviceNotification(windowHandle);
            }
        }
        #endregion

        #region detect usb plugged in
        private IntPtr UsbNotificationHandler(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam, ref bool handled)
        {
            if (msg == USBDetector.UsbDevicechange)
            {
                if (USBRadio.IsChecked == true)
                {
                    //viewmodel.removePort(deviceList);
                    //deviceList.Clear(); //清空避免重複
                    //getUSBInfoandUpdateView();
                    //viewmodel.getDeviceObserve("usb");
                    //DeviceSelect.SelectedIndex = viewmodel.Device.Count - 1;//設定選取
                    //                                                        //因為在usb狀態下開關，不會跑選擇usb device那段，就會少判斷一次
                    //if (viewmodel.Device.Count - 1 != -1)
                    //{
                    //    isPrinterConnected = true;
                    //}
                    //else
                    //{
                    //    isPrinterConnected = false;
                    //}
                }
                else if (EhernetRadio.IsChecked == true)
                {
                    //switch ((int)wparam) //因為wifi arp會記錄連線成功的ip，所以關機也會抓地到ip
                    //{
                    //    case USBDetector.UsbDeviceRemoved: //wifi arp
                    //        viewmodel.removePort(deviceList);
                    //        deviceList.Clear(); //清空避免重複
                    //        break;
                    //    case USBDetector.NewUsbDeviceConnected:
                    //        viewmodel.removePort(deviceList);
                    //        deviceList.Clear(); //清空避免重複
                    //        getWIFIIP();
                    //        viewmodel.getDeviceObserve("wifi");
                    //        DeviceSelect.SelectedIndex = viewmodel.Device.Count - 1;
                    //        break;
                    //}
                }
                //else if (RS232Radio.IsChecked == true)
                //{                   
                viewmodel.removePort(deviceList);
                deviceList.Clear(); //清空避免重複
                getSerialPort();
                viewmodel.getDeviceObserve("rs232");
                DeviceSelectRS232.SelectedIndex = viewmodel.Device.Count - 1;

                //}

            }
            handled = false;
            return IntPtr.Zero;
        }
        #endregion

        //=======================================裝置的選取==================================
        #region 選取傳輸通道
        private void ConnectType_SelectionChanged(object sender, RoutedEventArgs e)
        {

            if (USBRadio.IsChecked == true)
            {
                DeviceType = "USB";
                //deviceList.Clear(); //清空避免重複
                //getUSBInfoandUpdateView();
                //viewmodel.getDeviceObserve("usb");
                //DeviceSelect.SelectedIndex = viewmodel.Device.Count - 1;//設定選取第一筆
                //if (viewmodel.Device.Count - 1 != -1)
                //{
                //    isPrinterConnected = true;
                //}
                //else
                //{
                //    isPrinterConnected = false;
                //}
            }
            else if (EhernetRadio.IsChecked == true)
            {
                DeviceType = "Ethernet";
                //if (isPrinterConnected)
                //{
                //    deviceList.Clear(); //清空避免重複
                //    getWIFIIP();
                //    viewmodel.getDeviceObserve("wifi");
                //    DeviceSelect.SelectedIndex = viewmodel.Device.Count - 1;
                //}
                //else
                //{//因為wifi arp會記錄連線成功的ip，所以關機也會抓地到ip，透過usb是否連線來判斷是否關機
                //    viewmodel.removePort(deviceList);
                //    deviceList.Clear(); //清空避免重複
                //}
            }
            else if (RS232Radio.IsChecked == true)
            {
                DeviceType = "RS232";
                deviceList.Clear(); //清空避免重複
                getSerialPort();
                viewmodel.getDeviceObserve("rs232");
                DeviceSelectRS232.SelectedIndex = viewmodel.Device.Count - 1;
            }

        }
        #endregion

        //===============================語系切換與設定===========================

        #region 語系切換 
        //暫時關閉繁中
        private void Language_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (language.SelectedIndex)
            {
                case 0:
                    LoadLanguage("zh-TW");
                    break;
                case 1:
                    LoadLanguage("zh-CN");
                    break;
                case 2:
                    LoadLanguage("en-US");

                    break;
            }
        }
        //load language resource
        private void LoadLanguage(String name)
        {
            //用來取得作業系統資訊
            CultureInfo currentCultureInfo = CultureInfo.CurrentCulture;

            //用來取得資源字典內容
            ResourceDictionary langRd = null;

            try
            {
                //currentCultureInfo.Name可以取得目前作業系統語系(ex:zh-TW, zh-CN...)
                //抓取後設定為使用的資源字典內容
                langRd = Application.LoadComponent(
                new Uri(@"Language\" + name + ".xaml ", UriKind.Relative)) as ResourceDictionary;

            }
            catch
            {

            }

            if (langRd != null)
            {
                int ResourceCount = this.Resources.MergedDictionaries.Count;
                if (ResourceCount > 0)
                {
                    for (int i = 1; i < ResourceCount; i++)
                    {

                        //remove index!=0的所有recources，因為index=0的resource是themes
                        this.Resources.MergedDictionaries.RemoveAt(i);
                    }
                }
                this.Resources.MergedDictionaries.Add(langRd); //把抓取到的語系資源檔加入


            }
        }
        #endregion

        #region 依照os系統設定語系
        //暫時關閉繁中
        private void setDefaultLanguage()
        {
            string lan = "zh-CN";
            //string lan = CultureInfo.CurrentCulture.Name;
            Console.WriteLine(lan);
            switch (lan)
            {
                case "zh-TW":
                    language.SelectedIndex = 0;
                    break;
                case "zh-CN":
                    language.SelectedIndex = 1;
                    break;
                case "en-US":
                    language.SelectedIndex = 2;
                    break;
            }

        }





        #endregion


    }
}
