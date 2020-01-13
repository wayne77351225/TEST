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
            BaudRateCom.SelectedIndex = 3;
            App.Current.Properties["BaudRateSetting"] = 115200;
            //default 為選取RS232
            this.RS232Radio.IsChecked = true;

        }
        #endregion



        //========================Btn點擊事件===========================
        #region 通讯接口测试按鈕事件
        private void ConnectTest_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray("1F1B1F535A4A425A46100100");
            if ((bool)rs232Checkbox.IsChecked)
            {
                SerialPortConnect("PrinterInfo", sendArray);
            }
            // byte[] TETS = StringToByteArray("1F1B1F48461101005F5A4C5F5050543031005F564552312E3000086F0000000000000000000000C63D001D55");
            //SetPrinterInfo(TETS); //
        }
        #endregion

        #region 重啟印表機按鈕事件
        private void Restart_Click(object sender, RoutedEventArgs e)
        {
            switch (DeviceType)
            {

                case "RS232":
                    byte[] sendArray = StringToByteArray("1F1B1F535A4A425A46110000");
                    SerialPortConnect("Restart", sendArray);
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
                    switch (dataType )
                    {
                        case "PrinterInfo":
                        bool isReceiveData = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, 44);
                        while (!isReceiveData)
                        {

                            if (RS232Connect.mRecevieData != null)
                            {
                                //Console.WriteLine("result:" + Encoding.Default.GetString(RS232Connect.mRecevieData));
                                switch (dataType)
                                {
                                    case "PrinterInfo":
                                        //byte[] TETS = StringToByteArray("1F1B1F48461101005F5A4C5F5050543031005F564552312E3000086F0000000000000000000000C63D001D55");
                                        SetPrinterInfo(RS232Connect.mRecevieData); //
                                        RS232ConnectImage.Source = new BitmapImage(new Uri("Images/green_circle.png", UriKind.Relative));
                                        break;
                                        //case "GN":
                                        //    SetGNColumnData(RS232Connect.mRecevieData, 3);
                                        //    break;
                                        //case "SENSOR":
                                        //    SetSensorColumnData(RS232Connect.mRecevieData, 3);
                                        //    break;
                                }
                                break;
                            }
                        }
                            break;
                        case "Restart":
                        RS232Connect.SerialPortSendCMD("NoReceive", data, null,0);
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
            return Enumerable.Range(0, hex.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(hex.Substring(x, 2), 16))
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
