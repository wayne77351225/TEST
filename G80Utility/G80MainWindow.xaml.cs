using G80Utility.Tool;
using Microsoft.Win32;
using PirnterUtility.Models;
using PirnterUtility.Tool;
using PirnterUtility.ViewModels;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;


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

            //dip baud rade default
            DIPBaudRateCom.SelectedIndex = 1;
        }
        #endregion

        #region 參數設置核取框是否勾選檢查
        private void IsParaSettingChecked()
        {
            if (SetIPCheckbox.IsChecked == true)
            {
                Config.isSetIPChecked = true;
            }
            else
            {
                Config.isSetIPChecked = false;
            }

            if (SetGatewayCheckbox.IsChecked == true)
            {
                Config.isSetGatewayChecked = true;
            }
            else
            {
                Config.isSetGatewayChecked = false;
            }

            if (SetMACCheckbox.IsChecked == true)
            {
                Config.isSetMacChecked = true;
            }
            else
            {
                Config.isSetMacChecked = false;
            }

            if (AutoDisconnectheckbox.IsChecked == true)
            {
                Config.isAutoDisconnectChecked = true;
            }
            else
            {
                Config.isAutoDisconnectChecked = false;
            }

            if (ConnectClientbox.IsChecked == true)
            {
                Config.isConnectClientChecked = true;
            }
            else
            {
                Config.isConnectClientChecked = false;
            }

            if (EthernetSpeedCheckbox.IsChecked == true)
            {
                Config.isEthernetSpeedChecked = true;
            }
            else
            {
                Config.isEthernetSpeedChecked = false;
            }

            if (DHCPModeCheckbox.IsChecked == true)
            {
                Config.isDHCPModeChecked = true;
            }
            else
            {
                Config.isDHCPModeChecked = false;
            }

            if (USBModeCheckbox.IsChecked == true)
            {
                Config.isUSBModeChecked = true;
            }
            else
            {
                Config.isUSBModeChecked = false;
            }

            if (USBFixedCheckbox.IsChecked == true)
            {
                Config.isUSBFixedChecked = true;
            }
            else
            {
                Config.isUSBFixedChecked = false;
            }

            if (CodePageSetCheckbox.IsChecked == true)
            {
                Config.isCodePageSetChecked = true;
            }
            else
            {
                Config.isCodePageSetChecked = false;
            }

            if (LanguageSetCheckbox.IsChecked == true)
            {
                Config.isLanguageSetChecked = true;
            }
            else
            {
                Config.isLanguageSetChecked = false;
            }

            if (FontBSettingCheckbox.IsChecked == true)
            {
                Config.isFontBSettingChecked = true;
            }
            else
            {
                Config.isFontBSettingChecked = false;
            }

            if (CustomziedFontCheckbox.IsChecked == true)
            {
                Config.isCustomziedFontChecked = true;
            }
            else
            {
                Config.isCustomziedFontChecked = false;
            }

            if (DirectionCheckbox.IsChecked == true)
            {
                Config.isDirectionChecked = true;
            }
            else
            {
                Config.isDirectionChecked = false;
            }

            if (MotorAccControlCheckbox.IsChecked == true)
            {
                Config.isMotorAccControlChecked = true;
            }
            else
            {
                Config.isMotorAccControlChecked = false;
            }

            if (AccMotorCheckbox.IsChecked == true)
            {
                Config.isAccMotorChecked = true;
            }
            else
            {
                Config.isAccMotorChecked = false;
            }

            if (PrintSpeedCheckbox.IsChecked == true)
            {
                Config.isPrintSpeedChecked = true;
            }
            else
            {
                Config.isPrintSpeedChecked = false;
            }

            if (DensityModeCheckbox.IsChecked == true)
            {
                Config.isDensityModeChecked = true;
            }
            else
            {
                Config.isDensityModeChecked = false;
            }

            if (DensityCheckbox.IsChecked == true)
            {
                Config.isDensityChecked = true;
            }
            else
            {
                Config.isDensityChecked = false;
            }

            if (PaperOutReprintCheckbox.IsChecked == true)
            {
                Config.isPaperOutReprintChecked = true;
            }
            else
            {
                Config.isPaperOutReprintChecked = false;
            }

            if (PaperWidthCheckbox.IsChecked == true)
            {
                Config.isPaperWidthChecked = true;
            }
            else
            {
                Config.isPaperWidthChecked = false;
            }

            if (HeadCloseCutCheckbox.IsChecked == true)
            {
                Config.isHeadCloseCutChecked = true;
            }
            else
            {
                Config.isHeadCloseCutChecked = false;
            }

            if (YOffsetCheckbox.IsChecked == true)
            {
                Config.isYOffsetChecked = true;
            }
            else
            {
                Config.isYOffsetChecked = false;
            }

            if (MACShowCheckbox.IsChecked == true)
            {
                Config.isMACShowChecked = true;
            }
            else
            {
                Config.isMACShowChecked = false;
            }

            if (QRCodeCheckbox.IsChecked == true)
            {
                Config.isQRCodeChecked = true;
            }
            else
            {
                Config.isQRCodeChecked = false;
            }

            if (LogoPrintControlCheckbox.IsChecked == true)
            {
                Config.isLogoPrintControlhecked = true;
            }
            else
            {
                Config.isLogoPrintControlhecked = false;
            }

            if (DIPSwitchCheckbox.IsChecked == true)
            {
                Config.isDIPSwitchChecked = true;
            }
            else
            {
                Config.isDIPSwitchChecked = false;
            }
        }
        #endregion

        #region 工廠生產核取框是否勾選檢查
        private void IsFactortyChecked()
        {
            if (CMDQRCodeCheckbox.IsChecked == true)
            {
                Config.isCMDQRCodeChecked = true;
            }
            else
            {
                Config.isCMDQRCodeChecked = false;
            }

            if (CMDGeneralCheckbox.IsChecked == true)
            {
                Config.isCMDGeneralChecked = true;
            }
            else
            {
                Config.isCMDGeneralChecked = false;
            }

            if (CMDPageCheckbox.IsChecked == true)
            {
                Config.isCMDPageChecked = true;
            }
            else
            {
                Config.isCMDPageChecked = false;
            }
        }
        #endregion

        #region 傳送所有讀取指令
        private void readALL()
        {
            byte[] sendArray = null;
            if (Config.isSetIPChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 10 01");
                SendCmd(sendArray, "ReadPara", 12);

            }

            if (Config.isSetGatewayChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 11 01");
                SendCmd(sendArray, "ReadPara", 12);
            }

            if (Config.isSetMacChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 12 01");
                SendCmd(sendArray, "ReadPara", 14);
            }

            if (Config.isAutoDisconnectChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 13 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isConnectClientChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 14 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isEthernetSpeedChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 15 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isDHCPModeChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 16 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isUSBModeChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 17 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isUSBFixedChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "30 18 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isCodePageSetChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 36 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isLanguageSetChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 28 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isFontBSettingChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 23 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isCustomziedFontChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 29 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isDirectionChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 13 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isMotorAccControlChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 14 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isAccMotorChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 15 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isPrintSpeedChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 31 01");
                SendCmd(sendArray, "ReadPara", 10);
            }

            if (Config.isDensityModeChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 26 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isDensityChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 27 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isPaperOutReprintChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 21 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isPaperWidthChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 30 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isHeadCloseCutChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 17 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isYOffsetChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 18 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isMACShowChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 24 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isQRCodeChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 25 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isLogoPrintControlhecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 20 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            if (Config.isDIPSwitchChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 34 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            //最後發送DIP值讀取命令
            sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 35 01");
            SendCmd(sendArray, "ReadPara", 9);

        }
        #endregion

        #region 傳送所有寫入指令
        private void sendALL()
        {
            if (Config.isSetIPChecked)
            {
                SetIP();
            }

            if (Config.isSetGatewayChecked)
            {
                SetGateway();
            }

            if (Config.isSetMacChecked)
            {
                SetMAC();
            }

            if (Config.isAutoDisconnectChecked)
            {
                AutoDisconnect();
            }

            if (Config.isConnectClientChecked)
            {
                ConnectClient();
            }

            if (Config.isEthernetSpeedChecked)
            {
                EthernetSpeed();
            }

            if (Config.isDHCPModeChecked)
            {
                DHCPMode();
            }

            if (Config.isUSBModeChecked)
            {
                USBMode();
            }

            if (Config.isUSBFixedChecked)
            {
                USBFixed();
            }

            if (Config.isCodePageSetChecked)
            {
                CodePageSet();
            }

            if (Config.isLanguageSetChecked)
            {
                LanguageSet();
            }

            if (Config.isFontBSettingChecked)
            {
                FontBSetting();
            }

            if (Config.isCustomziedFontChecked)
            {
                CustomziedFont();
            }

            if (Config.isDirectionChecked)
            {
                SetDirection();
            }

            if (Config.isMotorAccControlChecked)
            {
                MotorAccControl();
            }

            if (Config.isAccMotorChecked)
            {
                AccMotor();
            }

            if (Config.isPrintSpeedChecked)
            {
                PrintSpeed();
            }

            if (Config.isDensityModeChecked)
            {
                DensityMode();
            }

            if (Config.isDensityChecked)
            {
                Density();
            }

            if (Config.isPaperOutReprintChecked)
            {
                PaperOutReprint();
            }

            if (Config.isPaperWidthChecked)
            {
                PaperWidth();
            }

            if (Config.isHeadCloseCutChecked)
            {
                HeadCloseCut();
            }

            if (Config.isYOffsetChecked)
            {
                YOffset();
            }

            if (Config.isMACShowChecked)
            {
                MACShow();
            }

            if (Config.isQRCodeChecked)
            {
                QRCode();
            }

            if (Config.isLogoPrintControlhecked)
            {
                LogoPrintControl();
            }

            if (Config.isDIPSwitchChecked)
            {
                DIPSwitch();
            }

            DIPSetting();
        }
        #endregion

        //========================取得資料後設定UI=================

        #region 取得打印機型號/軟件版本/機器序號
        private void SetPrinterInfo(byte[] buffer)
        {
            string sn = null;
            string moudle = null;
            string sfvesion = null;
            string date = null;

            //(0~7)前8個是無意義資料
            for (int i = 8; i < 24; i++)
            {
                sn += Convert.ToChar(buffer[i]);      //機器序列號
            }
            Console.WriteLine("sn:" + sn);
            for (int i = 24; i < 34; i++)
            {
                moudle += Convert.ToChar(buffer[i]);    //機器型號
            }
            if (moudle.Substring(0, 1) == "_")
            { //移除機器型號開頭的底線
                moudle = moudle.Remove(0, 1);
            }
            for (int i = 34; i < 44; i++)
            {
                sfvesion += Convert.ToChar(buffer[i]);   //軟件版本    
            }

            for (int i = 44; i < 54; i++)
            {
                date += Convert.ToChar(buffer[i]);    //多傳了一次軟件版本
            }

            PrinterSNFacTxt.Text = sn;
            PrinterSNTxt.Text = sn;
            PrinterModuleFac.Content = moudle + sfvesion + "：" + date;
            PrinterModule.Content = moudle + sfvesion + "：" + date;

        }
        #endregion

        #region 參數設置所有欄位設定內容
        public void setParaColumn(byte[] data)
        {
            string receiveData = BitConverter.ToString(data);
            Console.WriteLine(receiveData);

            if (receiveData.Contains(Command.RE_IP_CLASSFY))
            {
                checkIsGetData(SetIPText, null, data, "设置IP地址", false, 0);
            }

            if (receiveData.Contains(Command.RE_GATEWAY_CLASSFY))
            {
                checkIsGetData(SetGatewayText, null, data, "设置网关地址", false, 0);
            }

            if (receiveData.Contains(Command.RE_MAC_CLASSFY))
            {
                if (byteArraytoHexString(data, 8) != "")
                {
                    SetMACText.Text = byteArraytoHexString(data, 8); //因為這邊要轉hexstring所以不共用checkIsGetData()
                }
                else
                {
                    SysStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEF7171")); ;
                    SysStatusText.Text = "设置MAC地址" + FindResource("NotReadParameterYet") as string;
                }
            }

            if (receiveData.Contains(Command.RE_AUTODISCONNECT_CLASSFY))
            {
                checkIsGetData(null, AutoDisconnectCom, data, "自动断线时间", false, 8);
            }

            if (receiveData.Contains(Command.RE_CLIENTCOUNT_CLASSFY))
            {
                checkIsGetData(null, ConnectClientCom, data, "网络连接数量", true, 2);
            }

            if (receiveData.Contains(Command.RE_NETWORK_SPEED_CLASSFY))
            {
                checkIsGetData(null, EthernetSpeedCom, data, "网口通讯速度", false, 1);
            }

            if (receiveData.Contains(Command.RE_DHCP_MODE_CLASSFY))
            {
                checkIsGetData(null, DHCPModeCom, data, "DHCP模式", false, 3);
            }

            if (receiveData.Contains(Command.RE_USB_MODE_CLASSFY))
            {
                checkIsGetData(null, USBModeCom, data, "USB模式", false, 1);
            }

            if (receiveData.Contains(Command.RE_USB_FIX_CLASSFY))
            {
                checkIsGetData(null, USBFixedCom, data, "USB端口", false, 1);
            }

            if (receiveData.Contains(Command.RE_CODEPAGE_CLASSFY))
            {
                string code = receiveData.Substring(receiveData.Length - 2, 2); //取得收到hex string
                List<string> codeList = CodePage.getCodePageList();
                int index = 99; //設這個數代表沒有符合的選項就是讀取不到資料
                for (int i = 0; i < codeList.Count; i++)
                { //取得的會是個位數，前面要補0否則比對會有錯
                    string getItemCode = codeList[i].Split(':')[0];
                    if (getItemCode.Length < 2)
                    {
                        getItemCode = "0" + getItemCode;
                    }
                    if (getItemCode.Contains(code))
                    {
                        index = i;
                        break;
                    }
                }
                if (index != 99)
                {
                    CodePageCom.SelectedIndex = index;
                }
                else
                {
                    setSysStatusColorAndText("设置代码页" + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                }
            }

            if (receiveData.Contains(Command.RE_LANGUAGES_CLASSFY))
            {
                checkIsGetData(null, EthernetSpeedCom, data, "语言设置", true, 6);
            }

            if (receiveData.Contains(Command.RE_FONTB_CLASSFY))
            {
                checkIsGetData(null, FontBSettingCom, data, "FontB字体", false, 1);
            }

            if (receiveData.Contains(Command.RE_CUSTOMFONT_CLASSFY))
            {
                checkIsGetData(null, CustomziedFontCom, data, "定制字体", false, 1);
            }

            if (receiveData.Contains(Command.RE_DIRECTION_CLASSFY))
            {
                checkIsGetData(null, DirectionCombox, data, "走纸方向", false, 1);
            }

            if (receiveData.Contains(Command.RE_MOTOR_ACC_CONTROL_CLASSFY))
            {
                checkIsGetData(null, MotorAccControlCom, data, "马达加速", false, 1);
            }

            if (receiveData.Contains(Command.RE_MOTOR_ACC_CLASSFY))
            {
                switch (byteToIntForOneByte(data))
                {
                    case 10:
                        AccMotorCom.SelectedIndex = 0;
                        break;
                    case 8:
                        AccMotorCom.SelectedIndex = 1;
                        break;
                    case 6:
                        AccMotorCom.SelectedIndex = 2;
                        break;
                    case 4:
                        AccMotorCom.SelectedIndex = 3;
                        break;
                    case 2:
                        AccMotorCom.SelectedIndex = 4;
                        break;

                    default:
                        setSysStatusColorAndText("马达加速度" + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                        break;
                }

            }

            if (receiveData.Contains(Command.RE_PRINT_SPEED_CLASSFY))
            {
                string speed = hexStringToInt(receiveData).ToString();
                switch (speed)
                {
                    case "200":
                        PrintSpeedCom.SelectedIndex = 0;
                        break;
                    case "250":
                        PrintSpeedCom.SelectedIndex = 1;
                        break;
                    case "300":
                        PrintSpeedCom.SelectedIndex = 2;
                        break;
                    default:
                        setSysStatusColorAndText("打印速度" + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                        break;
                }

            }

            if (receiveData.Contains(Command.RE_DENSITY_MODE_CLASSFY))
            {
                checkIsGetData(null, DensityModeCom, data, "浓度模式", false, 1);
            }

            if (receiveData.Contains(Command.RE_DENSITY_CLASSFY))
            {
                checkIsGetData(null, DensityCom, data, "浓度调节", true, 10);
            }

            if (receiveData.Contains(Command.RE_PAPEROUT_CLASSFY))
            {
                checkIsGetData(null, PaperOutReprintCom, data, "纸尽重打", false, 1);
            }

            if (receiveData.Contains(Command.RE_PAPERWIDTH_CLASSFY))
            {
                checkIsGetData(null, PaperWidthCom, data, "打印纸宽", false, 1);
            }

            if (receiveData.Contains(Command.RE_HEADCLOSE_CUT_CLASSFY))
            {
                checkIsGetData(null, HeadCloseCutCom, data, "合盖后自动切纸", false, 1);
            }

            if (receiveData.Contains(Command.RE_YOFFSET_CLASSFY))
            {
                checkIsGetData(null, YOffsetCom, data, "垂直移动单位", false, 1);
            }

            if (receiveData.Contains(Command.RE_MACSHOW_CLASSFY))
            {
                checkIsGetData(null, MACShowCom, data, "MAC地址显示", false, 1);
            }

            if (receiveData.Contains(Command.RE_QRCODE_CLASSFY))
            {
                checkIsGetData(null, QRCodeCom, data, "二维码功能", false, 1);
            }

            if (receiveData.Contains(Command.RE_LOGOPRINT_CLASSFY))
            {
                checkIsGetData(null, LogoPrintControlCom, data, "自检页LOGO设置", false, 1);
            }

            if (receiveData.Contains(Command.RE_DIPSW_CLASSFY))
            {
                checkIsGetData(null, DIPSwitchCom, data, "DIP开关", false, 1);
            }

            if (receiveData.Contains(Command.RE_DIPVALUE_CLASSFY))
            {
                byte[] bytes = new byte[1];
                bytes[0] = data[data.Length - 1];
                BitArray diparray = new BitArray(bytes);
                //前五個選項因為0(false)代表有選，設定給checkbox時會剛好相反，所以這邊取反
                for (int i = 0; i < diparray.Length - 2; i++)
                {
                    if (diparray[i] == true)
                    {
                        diparray[i] = false;
                    }
                    else
                    {
                        diparray[i] = true;
                    }
                }
                CutterCheckBox.IsChecked = diparray[0];
                BeepCheckBox.IsChecked = diparray[1];
                DensityCheckBox.IsChecked = diparray[2];
                ChineseForbiddenCheckBox.IsChecked = diparray[3];
                CharNumberCheckBox.IsChecked = diparray[4];
                CashboxCheckBox.IsChecked = diparray[5];

                string binary = Convert.ToInt32(diparray.Get(6)).ToString() + Convert.ToInt32(diparray.Get(7)).ToString();
                switch (binary)
                {
                    case "11": //19200
                        DIPBaudRateCom.SelectedIndex = 0;
                        break;
                    case "10": //9600
                        DIPBaudRateCom.SelectedIndex = 1;
                        break;
                    case "01": //115200
                        DIPBaudRateCom.SelectedIndex = 2;
                        break;
                    case "00": //38400
                        DIPBaudRateCom.SelectedIndex = 3;
                        break;
                }
                Console.WriteLine(binary);
            }
        }
        #endregion

        #region 判斷參數設定欄位是否取得資料
        private void checkIsGetData(TextBox SelectedText, ComboBox SelectedCom, byte[] data, string msg, bool isSubtractOne, int itemFinalNo)
        {
            if (SelectedText != null) //判斷ip和gateway是否讀取到資料
            {
                if (byteArraytoIPV4(data, 8) != "")
                {
                    SelectedText.Text = byteArraytoIPV4(data, 8);
                }
                else
                {
                    setSysStatusColorAndText(msg + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                }
            }
            else if (SelectedCom != null) //判斷combobox是否讀取到資料
            {
                if (isSubtractOne)
                { //index 從1開始者
                    if (byteToIntForOneByte(data) >= 1 && byteToIntForOneByte(data) <= itemFinalNo)
                    {
                        SelectedCom.SelectedIndex = byteToIntForOneByte(data) - 1;
                    }
                    else
                    {
                        setSysStatusColorAndText(msg + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                    }

                }
                else
                {//index 從0開始者
                    if (byteToIntForOneByte(data) >= 0 && byteToIntForOneByte(data) <= itemFinalNo)
                    {
                        SelectedCom.SelectedIndex = byteToIntForOneByte(data);
                    }
                    else
                    {
                        setSysStatusColorAndText(msg + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                    }
                }
            }
        }
        #endregion

        #region 狀態列設定文字與顏色
        private void setSysStatusColorAndText(string msg, string color)
        {
            SysStatusText.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(color));
            SysStatusText.Text = msg;
        }

        #endregion

        //========================Btn點擊事件===========================

        #region 通讯接口测试按鈕事件
        private void ConnectTest_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.RS232_COMMUNICATION_TEST);
            if ((bool)rs232Checkbox.IsChecked)
            {
                SerialPortConnect("CommunicationTest", sendArray, 0);
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
                    SerialPortConnect("BeepOrSetting", sendArray, 0);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }

        }
        #endregion

        #region 讀取機器序列號(通訊)按鈕事件
        private void ReadSNBtn_Click(object sender, RoutedEventArgs e)
        {
            RoadPrinterSN();
        }
        #endregion

        #region 設置機器序列號(通訊)按鈕事件
        private void SetSNBtn_Click(object sender, RoutedEventArgs e)
        {
            SetPrinterSN("communication");
        }
        #endregion

        #region 讀取所有參數設定
        private void ReadAllBtn_Click(object sender, RoutedEventArgs e)
        {
            IsParaSettingChecked();
            readALL();
        }
        #endregion

        #region 寫入所有參數設定
        private void WriteAllBtn_Click(object sender, RoutedEventArgs e)
        {
            IsParaSettingChecked();
            sendALL();
        }
        #endregion

        #region 設定IP Address按鈕事件
        private void SetIPBtn_Click(object sender, RoutedEventArgs e)
        {
            SetIP();
        }
        #endregion

        #region  設定Gateway按鈕事件
        private void SetGatewayBtn_Click(object sender, RoutedEventArgs e)
        {
            SetGateway();
        }
        #endregion

        #region 設定MAC按鈕事件
        private void SetMACBtn_Click(object sender, RoutedEventArgs e)
        {
            SetMAC();
        }
        #endregion

        #region 設定自動斷線時間按鈕事件
        private void AutoDisconnectBtn_Click(object sender, RoutedEventArgs e)
        {
            AutoDisconnect();
        }
        #endregion

        #region 設定網路連接數量按鈕事件
        private void ConnectClientBtn_Click(object sender, RoutedEventArgs e)
        {
            ConnectClient();
        }
        #endregion

        #region 設定網口通訊速度按鈕事件
        private void EthernetSpeedBtn_Click(object sender, RoutedEventArgs e)
        {
            EthernetSpeed();
        }
        #endregion

        #region 設定DHCP模式按鈕事件
        private void DHCPModeBtn_Click(object sender, RoutedEventArgs e)
        {
            DHCPMode();
        }
        #endregion

        #region 設定USB模式按鈕事件
        private void USBModeBtn_Click(object sender, RoutedEventArgs e)
        {
            USBMode();
        }
        #endregion

        #region 設定USB端口值按鈕事件
        private void USBFixedBtn_Click(object sender, RoutedEventArgs e)
        {
            USBFixed();
        }
        #endregion

        #region 設定代碼頁按鈕事件
        private void CodePageSetBtn_Click(object sender, RoutedEventArgs e)
        {
            CodePageSet();
        }
        #endregion

        #region 打印代碼頁按鈕事件
        private void CodePagePrintBtn_Click(object sender, RoutedEventArgs e)
        {
            if (CodePageCom.SelectedIndex != -1)
            {
                List<byte> codePage = new List<byte>();
                byte[] header = Command.CODEPAGE_PRINT_HEADER;
                byte[] separate = Command.CODEPAGE_PRINT_SEPARATE;
                byte[] char1 = Command.CODEPAGE_PRINT_CHAR1;
                byte[] char2 = Command.CODEPAGE_PRINT_CHAR2;
                byte[] selectedCode;
                byte[] selectedName;

                //加入header
                for (int i = 0; i < header.Length; i++)
                {
                    codePage.Add(header[i]);
                }

                //取得代碼
                string HexCode = CodePageCom.SelectedItem.ToString();
                int Code = Int32.Parse(HexCode.Split(':')[0]);

                //取得代碼和名稱 byte array
                string CodeName = CodePageCom.SelectedItem.ToString();
                selectedCode = BitConverter.GetBytes(Code);
                selectedName = Encoding.Default.GetBytes(CodeName);

                //加入代碼
                codePage.Add(selectedCode[0]);

                //加入區隔
                for (int i = 0; i < separate.Length; i++)
                {
                    codePage.Add(separate[i]);
                }

                //加入名稱
                for (int i = 0; i < selectedName.Length; i++)
                {
                    codePage.Add(selectedName[i]);
                }

                //加入char1
                for (int i = 0; i < char1.Length; i++)
                {
                    codePage.Add(char1[i]);
                }

                //加入名稱
                for (int i = 0; i < selectedName.Length; i++)
                {
                    codePage.Add(selectedName[i]);
                }

                //加入char2
                for (int i = 0; i < char2.Length; i++)
                {
                    codePage.Add(char2[i]);
                }
                byte[] sendArray = codePage.ToArray();
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }

        }
        #endregion

        #region 設定語言按鈕事件
        private void LanguageSetBtn_Click(object sender, RoutedEventArgs e)
        {
            LanguageSet();
        }
        #endregion

        #region FontB設定按鈕事件
        private void FontBSettingBtn_Click(object sender, RoutedEventArgs e)
        {
            FontBSetting();
        }
        #endregion

        #region 設定定制字體按鈕事件
        private void CustomziedFontBtn_Click(object sender, RoutedEventArgs e)
        {
            CustomziedFont();
        }
        #endregion

        #region 走紙方向按鈕事件
        private void Direction_Click(object sender, RoutedEventArgs e)
        {
            SetDirection();
        }
        #endregion

        #region 設定馬達加速開關按鈕事件
        private void MotorAccControlBtn_Click(object sender, RoutedEventArgs e)
        {
            MotorAccControl();
        }
        #endregion

        #region 設定馬達加速度按鈕事件
        private void AccMotorBtn_Click(object sender, RoutedEventArgs e)
        {
            AccMotor();
        }
        #endregion

        #region 設定打印速度按鈕事件
        private void PrintSpeedBtn_Click(object sender, RoutedEventArgs e)
        {
            PrintSpeed();
        }
        #endregion

        #region 設定濃度模式按鈕事件
        private void DensityModeBtn_Click(object sender, RoutedEventArgs e)
        {
            DensityMode();
        }
        #endregion

        #region 設定濃度調節按鈕事件
        private void DensityBtn_Click(object sender, RoutedEventArgs e)
        {
            Density();
        }
        #endregion

        #region 設定紙盡重打按鈕事件
        private void PaperOutReprintBtn_Click(object sender, RoutedEventArgs e)
        {
            PaperOutReprint();
        }
        #endregion

        #region 設定打印紙寬按鈕事件
        private void PaperWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            PaperWidth();
        }
        #endregion

        #region 設定合蓋自動切紙按鈕事件
        private void HeadCloseCutBtn_Click(object sender, RoutedEventArgs e)
        {
            HeadCloseCut();
        }
        #endregion

        #region 設定垂直移動單位按鈕事件
        private void YOffsetBtn_Click(object sender, RoutedEventArgs e)
        {
            YOffset();
        }
        #endregion

        #region 設定MAC顯示按鈕事件
        private void MACShowBtn_Click(object sender, RoutedEventArgs e)
        {
            MACShow();
        }
        #endregion

        #region 設定二維碼按鈕事件
        private void QRCodeBtn_Click(object sender, RoutedEventArgs e)
        {
            QRCode();
        }
        #endregion

        #region 設定自檢頁logo按鈕事件
        private void LogoPrintControlBtn_Click(object sender, RoutedEventArgs e)
        {
            LogoPrintControl();
        }
        #endregion

        #region 設定DIP開關按鈕事件
        private void DIPSwitchBtn_Click(object sender, RoutedEventArgs e)
        {
            DIPSwitch();
        }
        #endregion

        #region DIP值設定按鈕事件
        private void DIPSettingBtn_Click(object sender, RoutedEventArgs e)
        {
            DIPSetting();
        }
        #endregion

        #region 打印自檢頁(短)按鈕事件
        private void PrintTest_S_Click(object sender, RoutedEventArgs e)
        {
            PrintTest("short");
        }
        #endregion

        #region 打印自檢頁(長)按鈕事件
        private void PrintTest_L_Click(object sender, RoutedEventArgs e)
        {
            PrintTest("long");
        }
        #endregion

        #region 打印均勻測試按鈕事件
        private void PrintTest_EVEN_Click(object sender, RoutedEventArgs e)
        {
            PrintEvenTest();
        }
        #endregion

        #region 蜂鳴器測試按鈕事件
        private void BeepTestBtn_Click(object sender, RoutedEventArgs e)
        {
            BeepTest();
        }
        #endregion

        #region 下踢錢箱按鈕事件
        private void OpenCashBoxBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenCashBox();
        }
        #endregion

        #region 連續切紙按鈕事件
        private void CutTimesBtn_Click(object sender, RoutedEventArgs e)
        {
            CutTimes();
        }
        #endregion

        #region 指令測試按鈕事件
        private void CMDTestBtn_Click(object sender, RoutedEventArgs e)
        {
            CMDTest("factory");
        }
        #endregion

        #region 讀取機器序列號(工廠)按鈕事件
        private void RoadPrinterSNFacBtn_Click(object sender, RoutedEventArgs e)
        {
            RoadPrinterSN();
        }
        #endregion

        #region 設置機器序列號(工廠)按鈕事件
        private void SetPrinterSNFacBtn_Click(object sender, RoutedEventArgs e)
        {
            SetPrinterSN("factory");
        }
        #endregion

        //========================參數設置每個寫入命令功能=================

        #region 設定IP Address
        private void SetIP()
        {
            byte[] sendArray = null;
            if (SetIPText.Text != "")
            {
                var address = SetIPText.Text;
                String result = String.Concat(address.Split('.').Select(x => int.Parse(x).ToString("X2")));
                sendArray = StringToByteArray(Command.IP_SETTING_HEADER + result);
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else
            {
                MessageBox.Show(FindResource("ColumnEmpty") as string);
            }
        }
        #endregion

        #region  設定Gateway
        private void SetGateway()
        {
            byte[] sendArray = null;
            if (SetGatewayText.Text != "")
            {
                var gateway = SetGatewayText.Text;
                String result = String.Concat(gateway.Split('.').Select(x => int.Parse(x).ToString("X2")));
                sendArray = StringToByteArray(Command.GATEWAY_SETTING_HEADER + result);
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else
            {
                MessageBox.Show(FindResource("ColumnEmpty") as string);
            }

        }
        #endregion

        #region 設定MAC
        private void SetMAC()
        {
            byte[] sendArray = null;
            Random random = new Random();
            int mac4 = random.Next(0, 255);
            int mac5 = random.Next(0, 255);
            int mac6 = random.Next(0, 255);

            string hexMac4 = mac4.ToString("X2"); //X:16進位,2:2位數
            string hexMac5 = mac5.ToString("X2");
            string hexMac6 = mac6.ToString("X2");

            //寫入MAC Address
            sendArray = StringToByteArray(Command.MAC_ADDRESS_SETTING_HEADER + "00 47 50" + hexMac4 + hexMac5 + hexMac6);
            SetMACText.Text = "00:47:50:" + hexMac4 + ":" + hexMac5 + ":" + hexMac6;
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray, 0);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }
        }
        #endregion

        #region 設定自動斷線時間
        private void AutoDisconnect()
        {
            byte[] sendArray = null;
            if (AutoDisconnectCom.SelectedIndex != -1)
            {
                switch (AutoDisconnectCom.SelectedIndex)
                {
                    case 0: //不設定=>55 00
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "55 00");
                        break;
                    case 1: //1min=>33 01
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 01");
                        break;
                    case 2: //2min=>33 02
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 02");
                        break;
                    case 3: //3min=>33 03
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 03");
                        break;
                    case 4: //4min=>33 04
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 04");
                        break;
                    case 5: //5min=>33 05
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 05");
                        break;
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定網路連接數量
        private void ConnectClient()
        {
            if (ConnectClientCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;
                if (ConnectClientCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.CONNECT_CLIENT_1_SETTING);
                }
                else if (ConnectClientCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.CONNECT_CLIENT_2_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定網口通訊速度
        private void EthernetSpeed()
        {
            if (EthernetSpeedCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;
                if (EthernetSpeedCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.ETHERNET_SPEED_SETTING_10MHZ);
                }
                else if (EthernetSpeedCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.ETHERNET_SPEED_SETTING_100MHZ);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定DHCP模式
        private void DHCPMode()
        {
            if (DHCPModeCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                switch (DHCPModeCom.SelectedIndex)
                {
                    case 0: //STATIC=>11
                        sendArray = StringToByteArray(Command.DHCP_MODE_SETTING_HEADER + "11");
                        break;
                    case 1: //MODE1=>22
                        sendArray = StringToByteArray(Command.DHCP_MODE_SETTING_HEADER + "22");
                        break;
                    case 2: //MODE2=>33
                        sendArray = StringToByteArray(Command.DHCP_MODE_SETTING_HEADER + "33");
                        break;
                    case 3: //MODE3=>44
                        sendArray = StringToByteArray(Command.DHCP_MODE_SETTING_HEADER + "44");
                        break;
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定USB模式
        private void USBMode()
        {
            if (USBModeCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;
                if (USBModeCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.USB_VCOM_SETTING);
                }
                else if (USBModeCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.USB_UTP_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定USB端口值
        private void USBFixed()
        {
            if (USBFixedCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (USBFixedCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.USB_UNFIXED_SETTING);
                }
                else if (USBFixedCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.USB_FIXED_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定代碼頁
        private void CodePageSet()
        {
            if (CodePageCom.SelectedIndex != -1)
            {
                //取得設定代碼
                string HexCode = CodePageCom.SelectedItem.ToString();
                HexCode = HexCode.Split(':')[0];
                if (HexCode.Length < 2)
                {
                    HexCode = "0" + HexCode;
                }
                byte[] sendArray = StringToByteArray(Command.CODEPAGE_SETTING_HEADER + HexCode);
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion       

        #region 設定語言
        private void LanguageSet()
        {

            if (LanguageSetCom.SelectedIndex != -1)
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
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region FontB設定
        private void FontBSetting()
        {
            if (FontBSettingCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (FontBSettingCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.FONTB_OFF_SETTING);
                }
                else if (FontBSettingCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.FONTB_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定定制字體
        private void CustomziedFont()
        {
            if (CustomziedFontCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (CustomziedFontCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.CUSTOMIZED_FONT_OFF_SETTING);
                }
                else if (CustomziedFontCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.CUSTOMIZED_FONT_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 走紙方向
        private void SetDirection()
        {
            if (DirectionCombox.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (DirectionCombox.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.DIRECTION_H80250N_SETTING);
                }
                else if (DirectionCombox.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.DIRECTION_80250N_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定馬達加速開關
        private void MotorAccControl()
        {
            if (MotorAccControlCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (MotorAccControlCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_OFF_SETTING);
                }
                else if (MotorAccControlCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定馬達加速度
        private void AccMotor()
        {
            if (AccMotorCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                switch (AccMotorCom.SelectedIndex)
                {
                    case 0: //1.0=>01
                        sendArray = StringToByteArray(Command.ACCELERATION_OF_MOTOR_SETTING + "01");
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
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定打印速度
        private void PrintSpeed()
        {
            if (PrintSpeedCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;
                switch (PrintSpeedCom.SelectedIndex)
                {
                    case 0: //200 mm/s
                        sendArray = StringToByteArray(Command.PRINT_SPEED_200_SETTING);
                        break;
                    case 1://250 mm/s
                        sendArray = StringToByteArray(Command.PRINT_SPEED_250_SETTING);
                        break;
                    case 2://300 mm/s
                        sendArray = StringToByteArray(Command.PRINT_SPEED_300_SETTING);
                        break;
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定濃度模式
        private void DensityMode()
        {
            if (DensityModeCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (DensityModeCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.DENSITY_MODE_LOW_SETTING);
                }
                else if (DensityModeCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.DENSITY_MODE_HIGH_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定濃度調節
        private void Density()
        {
            if (DensityCom.SelectedIndex != -1)
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
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定紙盡重打
        private void PaperOutReprint()
        {
            if (PaperOutReprintCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;
                if (PaperOutReprintCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.PAPEROUT_REPRINT_OFF_SETTING);
                }
                else if (PaperOutReprintCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.PAPEROUT_REPRINT_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定打印紙寬
        private void PaperWidth()
        {
            if (PaperWidthCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (PaperWidthCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.PAPER_WIDTH_58MM_SETTING);
                }
                else if (PaperWidthCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.PAPER_WIDTH_80MM_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定合蓋自動切紙
        private void HeadCloseCut()
        {
            if (HeadCloseCutCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (HeadCloseCutCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.HEADCLOSE_AUTOCUT_OFF_SETTING);
                }
                else if (HeadCloseCutCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.HEADCLOSE_AUTOCUT_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定垂直移動單位
        private void YOffset()
        {
            if (YOffsetCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (YOffsetCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.Y_OFFSET_1_SETTING);
                }
                else if (YOffsetCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.Y_OFFSET_05_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定MAC顯示
        private void MACShow()
        {
            if (MACShowCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (MACShowCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.MAC_SHOW_DEC_SETTING);
                }
                else if (MACShowCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.MAC_SHOW_HEX_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定二維碼
        private void QRCode()
        {
            if (QRCodeCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (QRCodeCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.QRCODE_OFF_SETTING);
                }
                else if (QRCodeCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.QRCODE_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定自檢頁logo
        private void LogoPrintControl()
        {
            if (LogoPrintControlCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (LogoPrintControlCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.LOGO_PRINT_OFF_SETTING);
                }
                else if (LogoPrintControlCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.LOGO_PRINT_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region 設定DIP開關
        private void DIPSwitch()
        {
            if (DIPSwitchCom.SelectedIndex != -1)
            {
                byte[] sendArray = null;

                if (DIPSwitchCom.SelectedIndex == 0)
                {
                    sendArray = StringToByteArray(Command.DIP_OFF_SETTING);
                }
                else if (DIPSwitchCom.SelectedIndex == 1)
                {
                    sendArray = StringToByteArray(Command.DIP_ON_SETTING);
                }
                switch (DeviceType)
                {
                    case "RS232":
                        SerialPortConnect("BeepOrSetting", sendArray, 0);
                        break;
                    case "USB":

                        break;
                    case "Ethernet":

                        break;
                }
            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
        }
        #endregion

        #region DIP值設定
        private void DIPSetting()
        {

            BitArray dipArray = new BitArray(8);
            byte[] sendArray = null;
            if (CutterCheckBox.IsChecked == true)
            {
                dipArray.Set(0, false);
            }
            else
            {
                dipArray.Set(0, true);
            }

            if (BeepCheckBox.IsChecked == true)
            {
                dipArray.Set(1, false);
            }
            else
            {
                dipArray.Set(1, true);
            }
            if (DensityCheckBox.IsChecked == true)
            {
                dipArray.Set(2, false);
            }
            else
            {
                dipArray.Set(2, true);
            }
            if (ChineseForbiddenCheckBox.IsChecked == true)
            {
                dipArray.Set(3, false);
            }
            else
            {
                dipArray.Set(3, true);
            }
            if (CharNumberCheckBox.IsChecked == true)
            {
                dipArray.Set(4, false);
            }
            else
            {
                dipArray.Set(4, true);
            }
            if (CashboxCheckBox.IsChecked == true)
            {
                dipArray.Set(5, false);
            }
            else
            {
                dipArray.Set(5, true);
            }

            switch (DIPBaudRateCom.SelectedIndex)
            {
                case 0: //19200 00取反11
                    dipArray.Set(6, true);
                    dipArray.Set(7, true);
                    break;
                case 1: //9600 01取反10
                    dipArray.Set(6, true);
                    dipArray.Set(7, false);
                    break;
                case 2: //115200 10取反 01
                    dipArray.Set(6, false);
                    dipArray.Set(7, true);
                    break;
                case 3: //38400 11取反00
                    dipArray.Set(6, false);
                    dipArray.Set(7, false);
                    break;
            }
            for (int i = 0; i < dipArray.Length; i++)
            {
                Console.WriteLine(dipArray.Get(i));
            }

            //bit array轉btye array
            byte[] bytes = new byte[1];
            dipArray.CopyTo(bytes, 0);
            sendArray = StringToByteArray(Command.DIP_VALUE_SETTING_HEADER);
            Array.Resize(ref sendArray, sendArray.Length + 1);
            sendArray[sendArray.Length - 1] = bytes[0];
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("BeepOrSetting", sendArray, 0);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }

        }
        #endregion

        //========================工廠生產功能=================
        #region 打印自檢頁
        private void PrintTest(string printType)
        {
            byte[] sendArray = null;

            if (printType == "short")
            {
                sendArray = StringToByteArray(Command.PRINT_TEST_SHORT);
            }
            else if (printType == "long")
            {
                sendArray = StringToByteArray(Command.PRINT_TEST_LONG);
            }
            SendCmd(sendArray, "BeepOrSetting", 0);

        }
        #endregion

        #region 打印均勻測試
        private void PrintEvenTest()
        {
            byte[] eventest = Command.PRINT_EVEN_TEST;
            SendCmd(eventest, "BeepOrSetting", 0);
        }
        #endregion

        #region 蜂鳴器測試
        private void BeepTest()
        {
            byte[] sendArray = StringToByteArray(Command.BEEP_TEST);
            SendCmd(sendArray, "BeepOrSetting", 0);
        }
        #endregion

        #region 下踢錢箱
        private void OpenCashBox()
        {
            byte[] sendArray = StringToByteArray(Command.OPEN_CASHBOX_TEST);
            SendCmd(sendArray, "BeepOrSetting", 0);
        }
        #endregion

        #region 連續切紙
        private void CutTimes()
        {
            int result;
            string times = CutTimesTxt.Text;
            if (times == "" || !Int32.TryParse(times, out result))
            {
                MessageBox.Show(FindResource("ErrorFormat") as string);
            }
            else
            {
                int contiunue = Int32.Parse(times);
                byte[] sendArray = StringToByteArray(Command.CUT_TIMES);
                for (int i = 0; i < contiunue; i++)
                {
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }
            }

        }
        #endregion

        #region 指令測試
        private void CMDTest(string functionClass)
        {
            switch (functionClass)
            {
                case "factory":
                    IsFactortyChecked();
                    if (Config.isCMDQRCodeChecked)
                    {
                        byte[] qrcode = Command.CMD_TEST_QRCODE;
                        SendCmd(qrcode, "BeepOrSetting", 0);
                    }

                    if (Config.isCMDGeneralChecked)
                    {
                        byte[] general = Command.CMD_TEST_GENERAL;
                        SendCmd(general, "BeepOrSetting", 0);
                    }

                    if (Config.isCMDPageChecked)
                    {
                        byte[] page = Command.CMD_TEST_PAGEMODE;
                        SendCmd(page, "BeepOrSetting", 0);
                    }
                    break;
                case "maintain":
                    break;
            }
        }
        #endregion

        #region 讀取機器序列號
        private void RoadPrinterSN()
        {
            byte[] sendArray = StringToByteArray(Command.DEVICE_INFO_READING);
            SendCmd(sendArray, "ReadSN", 54);

        }
        #endregion

        #region 設置機器序列號
        private void SetPrinterSN(string functionClass)
        {
            string sn = null;

            createSNRegistry();

            //初次未記錄,抓取user輸入值
            if (getSNRegistry() == null)
            {
                switch (functionClass)
                {
                    case "factory":
                        sn = PrinterSNFacTxt.Text;
                        break;
                    case "communication":
                        sn = PrinterSNTxt.Text;
                        break;
                }
            }
            else { //已經有紀錄，抓取最後一次序號值
                sn=getSNRegistry();               
                string number=sn.Substring(10, 6); 
                sn=sn.Remove(10, 6);
                int lastNumber = Int32.Parse(number);
                lastNumber+= 1; //sn每次加1
                number = lastNumber.ToString();
                if (number.Length < 6) {
                    //不能直接使用nubmer.Length，因為在迴圈的過程中length會越來越大
                    int nLen = number.Length;
                    for (int i = 1; i <= 6- nLen; i++) { //不足6位前面要補0
                        number= "0" + number;
                    }
                }
                sn = sn + number;
            }

            if (sn != "" && sn.Length == 16) //寫入序號到打印機
            {
                byte[] snArray = Encoding.Default.GetBytes(sn);
                byte[] sendArray = StringToByteArray(Command.SN_SETTING_HEADER);
                int sendLen = sendArray.Length;
                int snLen = snArray.Length;
                Array.Resize(ref sendArray, sendLen + snLen);
                for (int i = sendLen; i < sendLen + snLen; i++)
                {
                    sendArray[i] = snArray[i - sendLen];
                }
                SendCmd(sendArray, "BeepOrSetting", 0);
                setSNRegistry(sn); //寫入序號到註冊機碼
                //寫入序號到畫面
                PrinterSNFacTxt.Text = sn;
                PrinterSNTxt.Text = sn;
            }
            else if (sn.Length < 16 || sn.Length > 16)
            {
                MessageBox.Show(FindResource("LessLength") as string);
            }
        }
        #endregion

        //==========================註冊機碼的寫入與讀取===========================

        #region 註冊機碼位置的產生
        private void createSNRegistry()
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ZLPPT");
            //判斷是否有ZLPPT目錄
            if (registryKey == null) {
                registryKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ZLPPT");
                registryKey.SetValue("Path", "C:\\");
            }                    
        }
        #endregion

        #region 註冊機碼sn的寫入
        private void setSNRegistry(string sn)
        {
            RegistryKey registryKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ZLPPT"); //修改也要使用create不能用open
            registryKey.SetValue("SN", sn);
        }
        #endregion

        #region 註冊機碼sn的讀取
        private string getSNRegistry()
        {
            var lastSN = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ZLPPT").GetValue("SN");

            return (string)lastSN;
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
        private void SerialPortConnect(string dataType, byte[] data, int receiveLength)
        {
            RS232Connect.CloseSerialPort();
            if (RS232PortName != null)
            {
                bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);

                if (!isError)
                {
                    switch (dataType)
                    {
                        case "ReadPara":
                            bool isReceiveData = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            while (!isReceiveData)
                            {
                                if (RS232Connect.mRecevieData != null)
                                {
                                    setParaColumn(RS232Connect.mRecevieData);
                                    break;
                                }
                            }
                            break;
                        case "ReadSN":
                            bool isReceiveSN = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            while (!isReceiveSN)
                            {
                                if (RS232Connect.mRecevieData != null)
                                {
                                    SetPrinterInfo(RS232Connect.mRecevieData);
                                    break;
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

        //============================各種資料類型的轉換=============================

        #region hex string to byte array
        public static byte[] StringToByteArray(string hex)
        {
            string afterConvert = hex.Replace(" ", "");
            return Enumerable.Range(0, afterConvert.Length)
                             .Where(x => x % 2 == 0)
                             .Select(x => Convert.ToByte(afterConvert.Substring(x, 2), 16))
                             .ToArray();
        }
        #endregion

        #region byte array to IPV4
        public string byteArraytoIPV4(byte[] data, int startindex)
        {
            byte[] IPV4 = new byte[4];

            for (int i = 0; i < 4; i++)
            {
                IPV4[i] = data[startindex + i];
            }
            IPAddress ip = new IPAddress(IPV4);
            return ip.ToString();

        }
        #endregion

        #region byte array to hex string
        public string byteArraytoHexString(byte[] data, int startindex)
        {
            byte[] hexArray = new byte[6];

            for (int i = 0; i < 6; i++)
            {
                hexArray[i] = data[startindex + i];
            }
            string maxHex = BitConverter.ToString(hexArray).Replace("-", ":");
            return maxHex;

        }
        #endregion

        #region one byte to int
        public int byteToIntForOneByte(byte[] data)
        {
            byte convert = data[data.Length - 1];
            int intValue = Convert.ToInt32(convert);

            return intValue;

        }
        #endregion

        #region hex string to int
        public int hexStringToInt(string data)
        {
            string hex = data.Replace("-", "");
            hex = hex.Substring(hex.Length - 4, 4);
            int intValue = Convert.ToInt32(hex, 16);
            return intValue;

        }
        #endregion

        //========================不同通道傳送命令===========================

        #region 不同通道傳送命令
        public void SendCmd(byte[] sendArray, string sendType, int length)
        {
            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect(sendType, sendArray, length);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }
        }
        #endregion

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


        //===============================測試完可刪===========================

        private void Test_Click(object sender, RoutedEventArgs e)
        {

            byte[] sendArray = null;

            sendArray = StringToByteArray("1F 1B 1F 53 5A 4A 42 5A 46 12 01 01");

            switch (DeviceType)
            {
                case "RS232":
                    SerialPortConnect("ReadPara", sendArray, 54);
                    break;
                case "USB":

                    break;
                case "Ethernet":

                    break;
            }

        }


    }
}
