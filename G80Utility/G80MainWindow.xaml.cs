﻿using G80Utility.Models;
using G80Utility.Tool;
using G80Utility.ViewModels;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;

namespace G80Utility
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
        string EthernetIPAddress;

        //作業系統版本
        string OSVersion;

        //傳輸通道類別
        string DeviceType;

        //device 清單
        List<Device> deviceList = new List<Device>();

        //property 
        DeviceViewModel viewmodel { get; set; }

        //打印機實時狀態查詢位置
        string QueryNowStatusPosition;

        //nv logo 放大倍率參數m
        string nvLogo_m_hex;
        //nv logo 打印張數參數n，default打印1張
        string nvLogo_n_hex;
        //nv logo圖片集檔案路徑
        string[] fileNameArray;
        //nv logo 打印下載hex碼
        StringBuilder nvLogo_full_hex=new StringBuilder();

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

        #region 維護維修核打印機信息取框是否勾選檢查
        private void IsPrinterInfoChecked()
        {
            if (FeedLinesCheckbox.IsChecked == true)
            {
                Config.isFeedLinesChecked = true;
            }
            else
            {
                Config.isFeedLinesChecked = false;
            }

            if (PrintedLinesCheckbox.IsChecked == true)
            {
                Config.isPrintedLinesChecked = true;
            }
            else
            {
                Config.isPrintedLinesChecked = false;
            }

            if (CutPaperTimesCheckbox.IsChecked == true)
            {
                Config.isCutPaperTimesChecked = true;
            }
            else
            {
                Config.isCutPaperTimesChecked = false;
            }
            if (HeadOpenTimesCheckbox.IsChecked == true)
            {
                Config.isHeadOpenTimesChecked = true;
            }
            else
            {
                Config.isHeadOpenTimesChecked = false;
            }
            if (PaperOutTimesCheckbox.IsChecked == true)
            {
                Config.isPaperOutTimesChecked = true;
            }
            else
            {
                Config.isPaperOutTimesChecked = false;
            }
            if (ErrorTimesCheckbox.IsChecked == true)
            {
                Config.iErrorTimesChecked = true;
            }
            else
            {
                Config.iErrorTimesChecked = false;
            }
        }
        #endregion

        #region 維護維修核打印機信息取框是否全部勾選
        private bool IsPrinterInfoAllChecked()
        {
            bool isAllChecked = false;
            if (Config.isFeedLinesChecked
                && Config.isPrintedLinesChecked
                && Config.isCutPaperTimesChecked
                && Config.isHeadOpenTimesChecked
                && Config.isPaperOutTimesChecked
                && Config.iErrorTimesChecked)
            {
                isAllChecked = true;
            }

            return isAllChecked;
        }
        #endregion

        #region 維護維修核指令測試取框是否勾選檢查
        private void IsMaintainTestChecked()
        {
            if (CMDQRCode_Maintanin_Checkbox.IsChecked == true)
            {
                Config.isCMDQRCodeMaintainChecked = true;
            }
            else
            {
                Config.isCMDQRCodeMaintainChecked = false;
            }

            if (CMDGeneral_Maintanin_Checkbox.IsChecked == true)
            {
                Config.isCMDGeneralMaintainChecked = true;
            }
            else
            {
                Config.isCMDGeneralMaintainChecked = false;
            }

            if (CMDPage_Maintanin_Checkbox.IsChecked == true)
            {
                Config.isCMDPageMaintainChecked = true;
            }
            else
            {
                Config.isCMDPageMaintainChecked = false;
            }
        }
        #endregion

        #region nvLogo radio button選取確認
        private void nvLogoRadioBtnChecked()
        {

            if (SingleWidthandHeightRadio.IsChecked == true)
            {
                nvLogo_m_hex = "00";
            }
            else if (DoubleWidthRadio.IsChecked == true)
            {
                nvLogo_m_hex = "01";
            }
            else if (DoubleHeightRadio.IsChecked == true)
            {
                nvLogo_m_hex = "02";
            }
            else if (DoubleWidthandHeightRadio.IsChecked == true)
            {
                nvLogo_m_hex = "03";
            }
        }
        #endregion

        //========================取得資料後設定UI=================

        #region 顯示打印機型號/軟件版本/機器序號
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

        #region 顯示參數設置所有欄位設定內容
        public void setParaColumn(byte[] data)
        {
            string receiveData = BitConverter.ToString(data);
            //Console.WriteLine(receiveData);

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
                checkIsGetData(null, LanguageSetCom, data, "语言设置", true, 6);
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

        #region 判斷並顯示打印機實時狀態
        private void showPrinteNowStatus(byte[] data, UIElement uiContent)
        {
            byte[] bytes = new byte[1];
            bytes[0] = data[data.Length - 1];
            setPrinterStatustoUI(bytes, uiContent);
        }
        #endregion

        #region 取得並設定打印機狀態(溫度/電壓)
        private void setPrinterStatus(byte[] data)
        {
            //(0~7)前8個是無意義資料
            //字節1是狀態
            byte[] bytes = new byte[1];
            bytes[0] = data[8];
            setPrinterStatustoUI(bytes, PrinterStatusText);

            byte[] voltageArray = new byte[2];
            byte[] temperatureArray = new byte[2];
            for (int i = 9; i < 11; i++)
            {
                voltageArray[i - 9] = data[i];
            }
            for (int i = 11; i < 13; i++)
            {
                temperatureArray[i - 11] = data[i];
            }
            //電壓(字節2~3是電壓)        
            string voltageHex = BitConverter.ToString(voltageArray).Replace("-", "");
            double voltageDoule = Convert.ToInt32(voltageHex, 16) / 1000.000; //1伏特=1000毫伏
            voltageTxt.Text = voltageDoule.ToString();

            //溫度(字節4~5是溫度)    
            int temperatureInt = byteArraytoHexStringtoInt(temperatureArray);
            temperatureTxt.Text = temperatureInt.ToString();
        }
        #endregion

        #region 打印機狀態設定至畫面功能
        private void setPrinterStatustoUI(byte[] data, UIElement uiContent)
        {
            BitArray bitarray = new BitArray(data); //取得最後一個byte的bitarray 
            StringBuilder status = new StringBuilder();
            if (!bitarray[0]) //bit0 开盖
            {
                status.Append("开盖；");
            }
            if (!bitarray[1]) // bit1 缺纸
            {
                status.Append("缺纸；");

            }
            if (!bitarray[2]) // bit2 切刀错误
            {
                status.Append("切刀错误；");

            }
            if (!bitarray[3]) //钱箱状态
            {
                status.Append("钱箱打开；");
            }
            if (!bitarray[4]) //打印头超温
            {
                status.Append("打印头超温；");

            }
            if (!bitarray[5]) //已发生错误
            {
                status.Append("已发生错误；");

            }
            int count = 0;
            for (int i = 0; i < 6; i++)
            {
                if (bitarray[i])
                {
                    count++;
                }
                if (count == 6)
                {
                    status.Append("就绪");
                    break;
                }
            }

            switch (uiContent.GetType().Name)
            {
                case "Label":
                    Label lable = (Label)uiContent;
                    lable.Content = status;
                    break;
                case "TextBox":
                    TextBox textbox = (TextBox)uiContent;
                    textbox.Text = status.ToString();
                    break;
            }
        }

        #endregion

        #region 打印機信息設定至畫面功能
        private void setPrinterInfotoUI(byte[] data)
        {
            //(0~7)前8個是無意義資料
            byte[] receiveArray4Bytes = new byte[4];
            byte[] receiveArray2Bytes = new byte[2];
            int receiveInt = 0;
            for (int i = 8; i < 12; i++) //走紙行數
            {
                receiveArray4Bytes[i - 8] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray4Bytes);
            FeedLinesTxt.Text = receiveInt.ToString();
            for (int i = 12; i < 15; i++) //打印行數
            {
                receiveArray4Bytes[i - 12] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray4Bytes);
            PrintedLinesTxt.Text = receiveInt.ToString();
            for (int i = 15; i < 17; i++) //切紙次數
            {
                receiveArray2Bytes[i - 15] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            CutPaperTimesTxt.Text = receiveInt.ToString();
            for (int i = 17; i < 19; i++) //開蓋次數
            {
                receiveArray2Bytes[i - 17] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            HeadOpenTimesTxt.Text = receiveInt.ToString();
            for (int i = 19; i < 21; i++) //缺紙次數
            {
                receiveArray2Bytes[i - 19] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            PaperOutTimesTxt.Text = receiveInt.ToString();
            for (int i = 21; i < 23; i++) //故障次數
            {
                receiveArray2Bytes[i - 21] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            ErrorTimesTxt.Text = receiveInt.ToString();
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

        //通訊介面按鈕
        #region 通讯接口测试按鈕事件
        private void ConnectTest_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray;
            if ((bool)rs232Checkbox.IsChecked)
            {
                sendArray = StringToByteArray(Command.RS232_COMMUNICATION_TEST);
                SerialPortConnect("CommunicationTest", sendArray, 0);
            }
            if ((bool)USBCheckbox.IsChecked)
            {
                sendArray = StringToByteArray(Command.USB_COMMUNICATION_TEST);
                USBConnectAndSendCmd("CommunicationTest", sendArray, 0);
            }
            if ((bool)EthernetCheckbox.IsChecked)
            {
                bool isOK = chekckEthernetIPText();
                if (isOK)
                {
                    sendArray = StringToByteArray(Command.ETHERNET_COMMUNICATION_TEST);
                    EthernetConnectAndSendCmd("CommunicationTest", sendArray, 0);
                }
            }
        }
        #endregion

        #region 重啟印表機按鈕事件
        private void Restart_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.RESTART);
            SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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

        //維護維修按鈕
        #region 打印機維護維修tab按鈕事件
        private void MaintainTab_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            QueryNowStatusPosition = "maintain";
            PrinterInfoRead();
            PrinterNowStatus();
            queryPrinterStatus();
        }
        #endregion

        #region 打印機信息查詢按鈕事件
        private void PrinterInfoReadBtn_Click(object sender, RoutedEventArgs e)
        {
            PrinterInfoRead();
        }
        #endregion

        #region 打印機全選按鈕事件
        private void SelectAllBtn_Click(object sender, RoutedEventArgs e)
        {
            IsPrinterInfoChecked();
            if (!Config.isFeedLinesChecked)
            {
                FeedLinesCheckbox.IsChecked = true;
            }

            if (!Config.isPrintedLinesChecked)
            {
                PrintedLinesCheckbox.IsChecked = true;
            }

            if (!Config.isCutPaperTimesChecked)
            {
                CutPaperTimesCheckbox.IsChecked = true;
            }

            if (!Config.isHeadOpenTimesChecked)
            {
                HeadOpenTimesCheckbox.IsChecked = true;
            }

            if (!Config.isPaperOutTimesChecked)
            {
                PaperOutTimesCheckbox.IsChecked = true;
            }
            if (!Config.iErrorTimesChecked)
            {
                ErrorTimesCheckbox.IsChecked = true;
            }

        }
        #endregion

        #region 打印機清除所有信息按鈕事件
        private void CLeanPrinterInfoBtn_Click(object sender, RoutedEventArgs e)
        {
            if (IsPrinterInfoAllChecked())
            {
                //清除所有的打印机统计信息
                cleanPrinterInfo();
            }
            else
            {
                IsPrinterInfoChecked();//先確認選取狀態
                if (Config.isFeedLinesChecked)
                {
                    byte[] sendArray = StringToByteArray(Command.CLEAN_PRINTINFO_FEED_LINES);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }

                if (Config.isPrintedLinesChecked)
                {
                    byte[] sendArray = StringToByteArray(Command.CLEAN_PRINTINFO_PRINTED_LINES);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }

                if (Config.isCutPaperTimesChecked)
                {
                    byte[] sendArray = StringToByteArray(Command.CLEAN_PRINTINFO_CUTPAPER_TIMES);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }

                if (Config.isHeadOpenTimesChecked)
                {
                    byte[] sendArray = StringToByteArray(Command.CLEAN_PRINTINFO_HEADOPEN_TIMES);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }

                if (Config.isPaperOutTimesChecked)
                {
                    byte[] sendArray = StringToByteArray(Command.CLEAN_PRINTINFO_PAPEROUT_TIMES);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }
                if (Config.iErrorTimesChecked)
                {
                    byte[] sendArray = StringToByteArray(Command.CLEAN_PRINTINFO_ERROR_TIMES);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }

            }

            //清除完就要讀取打印機信息
            PrinterInfoRead();
        }
        #endregion

        #region 打印機狀態信息查詢按鈕事件
        private void PrinterStatusQueryBtn_Click(object sender, RoutedEventArgs e)
        {
            QueryNowStatusPosition = "maintain";
            PrinterNowStatus();
            queryPrinterStatus();
        }
        #endregion

        #region 打印自檢頁(短)-維護-按鈕事件
        private void PrintTest_S_Maintanin_Click(object sender, RoutedEventArgs e)
        {
            PrintTest("short");
        }
        #endregion

        #region 打印自檢頁(長)-維護-按鈕事件
        private void PrintTest_L_Maintanin_Click(object sender, RoutedEventArgs e)
        {
            PrintTest("long");
        }
        #endregion

        #region 打印均勻測試-維護-按鈕事件
        private void PrintTest_EVEN_Maintanin_Click(object sender, RoutedEventArgs e)
        {
            PrintEvenTest();
        }
        #endregion

        #region 蜂鳴器測試-維護-按鈕事件
        private void BeepTest_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            BeepTest();
        }
        #endregion

        #region 下踢錢箱-維護-按鈕事件
        private void OpenCashBox_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            OpenCashBox();
        }
        #endregion

        #region 連續切紙-維護-按鈕事件
        private void CutTimes_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            CutTimes("maintain");
        }
        #endregion

        #region 指令測試-維護-按鈕事件
        private void CMDTest_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            CMDTest("factory");
        }
        #endregion

        #region SDRAM測試按鈕事件
        private void SDRAMTestBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.SDRAM_TEST);
            SendCmd(sendArray, "BeepOrSetting", 0);
        }
        #endregion

        #region Flash測試按鈕事件
        private void FlashTestBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.FLASH_TEST);
            SendCmd(sendArray, "BeepOrSetting", 0);
        }
        #endregion

        //工廠生產按鈕
        #region 打印自檢頁(短)-工廠-按鈕事件
        private void PrintTest_S_Click(object sender, RoutedEventArgs e)
        {
            PrintTest("short");
        }
        #endregion

        #region 打印自檢頁(長)-工廠-按鈕事件
        private void PrintTest_L_Click(object sender, RoutedEventArgs e)
        {
            PrintTest("long");
        }
        #endregion

        #region 打印均勻測試-工廠-按鈕事件
        private void PrintTest_EVEN_Click(object sender, RoutedEventArgs e)
        {
            PrintEvenTest();
        }
        #endregion

        #region 蜂鳴器測試-工廠-按鈕事件
        private void BeepTestBtn_Click(object sender, RoutedEventArgs e)
        {
            BeepTest();
        }
        #endregion

        #region 下踢錢箱-工廠-按鈕事件
        private void OpenCashBoxBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenCashBox();
        }
        #endregion

        #region 連續切紙-工廠-按鈕事件
        private void CutTimesBtn_Click(object sender, RoutedEventArgs e)
        {
            CutTimes("factory");
        }
        #endregion

        #region 指令測試-工廠-按鈕事件
        private void CMDTestBtn_Click(object sender, RoutedEventArgs e)
        {
            CMDTest("maintain");
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

        #region 出廠設置按鈕事件
        private void FactoryDefaultBtn_Click(object sender, RoutedEventArgs e)
        {
            //清除所有的打印机统计信息
            cleanPrinterInfo();

            //根据参数设置界面的复选框进行所有参数的下载
            IsParaSettingChecked();
            sendALL();

            //发送打印自检页（长）命令
            PrintTest("long");
        }
        #endregion

        //NVLogo按鈕
        #region 打印logo按鈕事件
        private void PrintLogoBtn_Click(object sender, RoutedEventArgs e)
        {
            nvLogoRadioBtnChecked();
            int number;
            if (NVLogoPieceTXT.Text != null && Int32.TryParse(NVLogoPieceTXT.Text, out number))
            {
                string numberHex = number.ToString("X2");
                byte[] sendArray = StringToByteArray(Command.PRINT_LOGOS_HEADER + numberHex + nvLogo_m_hex);
                SendCmd(sendArray, "BeepOrSetting", 0);
            }
            else {
                MessageBox.Show(FindResource("PrintPieceEmpty") as string);
            }
            
        }
        #endregion

        #region 清除logo下載按鈕事件
        private void ClearLogoBtn_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray = StringToByteArray(Command.CLEAN_LOGOS_INPRINTER);
            SendCmd(sendArray, "BeepOrSetting", 0);
        }
        #endregion

        #region 打開檔案增加圖片集按鈕事件
        private void OpenImgFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Bitmap Image|*.bmp|JPeg Image|*.jpg|Gif Image|*.gif|Png Image|*.png|TIFF Image|*.tif";
            openFileDialog.FilterIndex = 1;
            //record last directory
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Multiselect = true;

            if ((bool)openFileDialog.ShowDialog())
            {
                if (fileNameArray == null) //第一次開啟
                {
                    fileNameArray = openFileDialog.FileNames;                 
                    nvLogo_full_hex.Append("1C71");
                }
                else 
                { //未清除再開啟
                    int oldArrayLen = fileNameArray.Length;
                    int newArrayLen = openFileDialog.FileNames.Length;
                    Array.Resize(ref fileNameArray, oldArrayLen + newArrayLen);
                    for (int i = oldArrayLen; i < oldArrayLen + newArrayLen; i++)
                    {
                        fileNameArray[i] = openFileDialog.FileNames[i - oldArrayLen];
                    }
                    if (fileNameArray.Length > 255)
                    {
                        MessageBox.Show(FindResource("ExceedMaxNumber") as string);
                        return;
                    }
                }
                for (int i = 0; i < fileNameArray.Length; i++)
                {               
                    Uri url = new Uri(fileNameArray[i]);
                    BitmapImage bmpImg = new BitmapImage(url);
                    Bitmap bmp = BitmapTool.BitmapImageToBitmap(bmpImg);
                    if (!BitmapTool.checkBitmapRange(fileNameArray[i], bmp.Width, bmp.Height)) //判斷寬高是否超過標準
                    {
                        fileNameArray = Array.FindAll(fileNameArray, val => val != fileNameArray[i]).ToArray();
                    }
                }

                //取得圖片集張數
                OpenImgNoTxt.Text = fileNameArray.Length.ToString();

                double calHeight = 0;
                for (int i = 0; i < fileNameArray.Length; i++)
                {
                    Uri url = new Uri(fileNameArray[i]);
                    BitmapImage bmpImg = new BitmapImage(url);
                    Bitmap bmp = BitmapTool.ToGray(BitmapTool.BitmapImageToBitmap(bmpImg),1);
                    
                        System.Windows.Controls.Image img = new System.Windows.Controls.Image();
                        img.Stretch = Stretch.Fill;
                        img.StretchDirection = StretchDirection.Both;
                        TextBlock text = new TextBlock();

                        if (!BitmapTool.isBlackWhite(bmp)) //彩色圖片
                        {
                            bmp = BitmapTool.Thresholding(bmp);
                            img.Source = BitmapTool.BitmapToBitmapImage(bmp);
                            nvLogo_full_hex.Append(BitmapTool.getBmpHighandLowHex(bmp.Width, bmp.Height));
                            nvLogo_full_hex.Append(BitmapTool.bitmapToHexString(bmp));
                        }
                        else
                        { //黑白圖片
                            nvLogo_full_hex.Append(BitmapTool.getBmpHighandLowHex(bmp.Width, bmp.Height));
                            nvLogo_full_hex.Append(BitmapTool.bitmapToHexString(bmp));
                             img.Source = bmpImg;
                        }

                        //圖片呈現在畫面上
                        text.FontSize = 16;
                        text.Text = FindResource("Ordinal") as string + (i + 1) + FindResource("piece") as string; //第x張
                        text.Height = 25;
                        int PaddingTop = 30;
                        int TextMargin = 15;
                        //判斷圖片形狀
                        if (bmp.Width == bmp.Height) //square
                        {
                            img.Width = 100;
                            img.Height = 100;
                            Canvas.SetLeft(img, 20); //img 左邊起點
                            Canvas.SetLeft(text, 20); //txt 左邊起點
                        }
                        else if (bmp.Width < bmp.Height) //vertical rec
                        {
                            img.Width = 100;
                            img.Height = 200;
                            Canvas.SetLeft(img, 20); //img 左邊起點
                            Canvas.SetLeft(text, 20); //txt 左邊起點
                        }
                        else
                        { //horizontal rec
                            img.Width = 200;
                            img.Height = 100;
                            Canvas.SetLeft(img, 20); //img 左邊起點
                            Canvas.SetLeft(text, 20); //txt 左邊起點
                        }

                        if (i == 0)
                        {
                            Canvas.SetTop(img, PaddingTop + (TextMargin * 2 + text.Height) * i);
                            Canvas.SetTop(text, PaddingTop + TextMargin + img.Height);
                            calHeight += img.Height; //放在後面家才不會重複算
                        }
                        else
                        {
                            Canvas.SetTop(img, PaddingTop + (TextMargin * 2 + text.Height) * i + calHeight);
                            Canvas.SetTop(text, PaddingTop + (TextMargin * 2 + text.Height) * i + calHeight + TextMargin + img.Height);
                            calHeight += img.Height;
                        }

                        NVlogoImg.Children.Add(img);
                        NVlogoImg.Children.Add(text);      
                }
            }
        }
        #endregion

        #region 清除圖片集按鈕事件
        private void CleanGalleryBtn_Click(object sender, RoutedEventArgs e)
        {
            NVlogoImg.Children.Clear();
            nvLogo_full_hex.Clear();
            fileNameArray = null;
            OpenImgNoTxt.Text="";
        }
        #endregion

        #region 下載圖片集按鈕事件
        private void DonwaldLogoBtn_Click(object sender, RoutedEventArgs e)
        {
            if (nvLogo_full_hex.Length == 0)
            {
                MessageBox.Show(FindResource("GalleryEmpty") as string);
            }
            else {
                nvLogo_n_hex = fileNameArray.Length.ToString("X2");
                //因為hex碼是兩個string,1c71長度佔去4，index插入要=4
                nvLogo_full_hex.Insert(4,nvLogo_n_hex);
                byte[]  sendArray = StringToByteArray(nvLogo_full_hex.ToString());
                Console.WriteLine(nvLogo_full_hex.ToString());
                SendCmd(sendArray, "BeepOrSetting", 0);
            }
        }
        #endregion
        //========================參數設置每個寫入命令功能=================

        #region 設定IP Address
        private void SetIP()
        {
            byte[] sendArray = null;
            string ip;
            if (SetIPText.Text != "")
            {
                ip = SetIPText.Text;
                if (checkIPFormat(ip))
                {
                    String result = String.Concat(ip.Split('.').Select(x => int.Parse(x).ToString("X2")));
                    sendArray = StringToByteArray(Command.IP_SETTING_HEADER + result);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }
                else
                {
                    ip = null;
                    MessageBox.Show(FindResource("ErrorFormat") as string);
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
            string gateway = null;
            if (SetGatewayText.Text != "")
            {
                gateway = SetGatewayText.Text;
                if (checkIPFormat(gateway))
                {
                    String result = String.Concat(gateway.Split('.').Select(x => int.Parse(x).ToString("X2")));
                    sendArray = StringToByteArray(Command.GATEWAY_SETTING_HEADER + result);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }
                else
                {
                    gateway = null;
                    MessageBox.Show(FindResource("ErrorFormat") as string);
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
            SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
                SendCmd(sendArray, "BeepOrSetting", 0);
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
            SendCmd(sendArray, "BeepOrSetting", 0);
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

        //=============================維護維修功能=============================
        #region 讀取打印機所有統計信息
        private void PrinterInfoRead()
        {
            byte[] sendArray = StringToByteArray(Command.READ_PRINTINFO);
            SendCmd(sendArray, "ReadPrinterInfo", 28);
        }
        #endregion

        #region 清除打印機所有統計信息
        private void cleanPrinterInfo()
        {
            byte[] sendArray = StringToByteArray(Command.CLEAN_ALL_PRINTINFO);
            SendCmd(sendArray, "BeepOrSetting", 0);
        }
        #endregion

        #region  打印機狀態(電壓/溫度...)查詢
        private void queryPrinterStatus()
        {
            byte[] sendArray = StringToByteArray(Command.READ_PRINTERSTATUS);
            SendCmd(sendArray, "ReadStatus", 13);
        }
        #endregion

        //==============================工廠生產功能=============================
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
        private void CutTimes(string type)
        {
            int result;
            string times = null;
            if (type == "factory")
            {
                times = CutTimesTxt.Text;
            }
            else if (type == "maintain")
            {

                times = CutTimes_Maintanin_Txt.Text;
            }

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
                    IsMaintainTestChecked();
                    if (Config.isCMDQRCodeMaintainChecked)
                    {
                        byte[] qrcode = Command.CMD_TEST_QRCODE;
                        SendCmd(qrcode, "BeepOrSetting", 0);
                    }

                    if (Config.isCMDGeneralMaintainChecked)
                    {
                        byte[] general = Command.CMD_TEST_GENERAL;
                        SendCmd(general, "BeepOrSetting", 0);
                    }

                    if (Config.isCMDPageMaintainChecked)
                    {
                        byte[] page = Command.CMD_TEST_PAGEMODE;
                        SendCmd(page, "BeepOrSetting", 0);
                    }
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
            //建立註冊機碼目錄
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
            else
            { //已經有紀錄，抓取最後一次序號值
                sn = getSNRegistry();
                string number = sn.Substring(10, 6);
                sn = sn.Remove(10, 6);
                int lastNumber = Int32.Parse(number);
                lastNumber += 1; //sn每次加1
                number = lastNumber.ToString();
                if (number.Length < 6)
                {
                    //不能直接使用nubmer.Length，因為在迴圈的過程中length會越來越大
                    int nLen = number.Length;
                    for (int i = 1; i <= 6 - nLen; i++)
                    { //不足6位前面要補0
                        number = "0" + number;
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

        //==============================打印機即時狀態============================

        #region 打印機即時狀態
        private void PrinterNowStatus()
        {
            byte[] sendArray = StringToByteArray(Command.STATUS_MONITOR);
            SendCmd(sendArray, "ReadNowStatus", 9);
        }
        #endregion


        //==========================註冊機碼的寫入與讀取===========================

        #region 註冊機碼位置的產生
        private void createSNRegistry()
        {
            RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ZLPPT");
            //判斷是否有ZLPPT目錄
            if (registryKey == null)
            {
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
                        case "ReadPrinterInfo": //打印機統計信息
                            bool isReceivePI = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            while (!isReceivePI)
                            {
                                if (RS232Connect.mRecevieData != null)
                                {
                                    setPrinterInfotoUI(RS232Connect.mRecevieData);
                                    break;
                                }
                            }
                            break;
                        case "ReadStatus": //打印機溫度電壓等狀態
                            bool isReceiveStatus = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            while (!isReceiveStatus)
                            {
                                if (RS232Connect.mRecevieData != null)
                                {
                                    setPrinterStatus(RS232Connect.mRecevieData);
                                    break;
                                }
                            }
                            break;
                        case "ReadNowStatus": //即時狀態
                            bool isReceiveNowStatus = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            while (!isReceiveNowStatus)
                            {
                                if (RS232Connect.mRecevieData != null)
                                {
                                    switch (QueryNowStatusPosition)
                                    {
                                        case "bottom":
                                            showPrinteNowStatus(RS232Connect.mRecevieData, StatusMonitorLabel);
                                            break;
                                        case "maintain":
                                            showPrinteNowStatus(RS232Connect.mRecevieData, PrinterStatusText);
                                            break;
                                    }

                                    break;
                                }
                            }
                            break;
                        case "CommunicationTest": //通訊測試
                            RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
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

        //===================USB 傳送指令==================

        private void USBConnectAndSendCmd(string dataType, byte[] data, int receiveLength)
        {
            if (USBpath != null)
            {
                int result = USBConnect.ConnectUSBDevice(USBpath);

                if (result == 1)
                {
                    USBConnectImage.Source = new BitmapImage(new Uri("Images/green_circle.png", UriKind.Relative));
                    switch (dataType)
                    {
                        case "ReadPara":
                        //bool isReceiveData = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        //while (!isReceiveData)
                        //{
                        //    if (RS232Connect.mRecevieData != null)
                        //    {
                        //        setParaColumn(RS232Connect.mRecevieData);
                        //        break;
                        //    }
                        //}
                        //break;
                        case "ReadSN":
                            //bool isReceiveSN = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            //while (!isReceiveSN)
                            //{
                            //    if (RS232Connect.mRecevieData != null)
                            //    {
                            //        SetPrinterInfo(RS232Connect.mRecevieData);
                            //        break;
                            //    }
                            //}
                            break;
                        case "ReadPrinterInfo": //打印機統計信息
                            //bool isReceivePI = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            //while (!isReceivePI)
                            //{
                            //    if (RS232Connect.mRecevieData != null)
                            //    {
                            //        setPrinterInfotoUI(RS232Connect.mRecevieData);
                            //        break;
                            //    }
                            //}
                            break;
                        case "ReadStatus": //打印機溫度電壓等狀態
                                           //    bool isReceiveStatus = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                                           //    while (!isReceiveStatus)
                                           //    {
                                           //        if (RS232Connect.mRecevieData != null)
                                           //        {
                                           //            setPrinterStatus(RS232Connect.mRecevieData);
                                           //            break;
                                           //        }
                                           //    }
                            break;
                        case "ReadNowStatus":
                            //bool isReceiveNowStatus = RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                            //while (!isReceiveNowStatus)
                            //{
                            //    if (RS232Connect.mRecevieData != null)
                            //    {
                            //switch (QueryNowStatusPosition)
                            //{
                            //    case "bottom":
                            //        showPrinteNowStatus(RS232Connect.mRecevieData, StatusMonitorLabel);
                            //        break;
                            //    case "maintain":
                            //        showPrinteNowStatus(RS232Connect.mRecevieData, PrinterStatusText);
                            //        break;
                            //}
                            //        break;
                            //    }
                            //}
                            break;
                        case "CommunicationTest": //通訊測試
                            USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                            break;
                        case "BeepOrSetting":
                            USBConnect.USBSendCMD("NoReceive", data, null, 0);
                            break;
                    }
                }
                else
                {
                    MessageBox.Show(FindResource("NotSettingUSBport") as string);
                    USBConnect.closeHandle();
                    USBConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //連線失敗時
                }
            }
            else
            {
                MessageBox.Show(FindResource("NotSettingUSBport") as string);
                USBConnect.closeHandle();
                USBConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //連線失敗時
            }
        }

        //===================Ethernet 傳送指令==================

        private void EthernetConnectAndSendCmd(string dataType, byte[] data, int receiveLength)
        {
            bool isConnect = EthernetConnect.connectToPrinter();
            if (isConnect)
            {
                switch (dataType)
                {
                    case "ReadPara":
                        bool isReceiveData = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setParaColumn(EthernetConnect.mRecevieData);
                                EthernetConnect.disconnect();
                                break;
                            }
                        }
                        break;
                    case "ReadSN":
                        bool isReceiveSN = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!isReceiveSN)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                SetPrinterInfo(EthernetConnect.mRecevieData);
                                EthernetConnect.disconnect();
                                break;
                            }
                        }
                        break;
                    case "ReadPrinterInfo": //打印機統計信息
                        bool isReceivePI = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!isReceivePI)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setPrinterInfotoUI(EthernetConnect.mRecevieData);
                                EthernetConnect.disconnect();
                                break;
                            }
                        }
                        break;
                    case "ReadStatus": //打印機溫度電壓等狀態
                        bool isReceiveStatus = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!isReceiveStatus)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setPrinterStatus(EthernetConnect.mRecevieData);
                                EthernetConnect.disconnect();
                                break;
                            }
                        }
                        break;
                    case "ReadNowStatus":
                        bool isReceiveNowStatus = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!isReceiveNowStatus)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                switch (QueryNowStatusPosition)
                                {
                                    case "bottom":
                                        showPrinteNowStatus(EthernetConnect.mRecevieData, StatusMonitorLabel);
                                        break;
                                    case "maintain":
                                        showPrinteNowStatus(EthernetConnect.mRecevieData, PrinterStatusText);
                                        break;
                                }
                                EthernetConnect.disconnect();
                                break;
                            }
                        }
                        break;
                    case "CommunicationTest": //通訊測試
                        EthernetConnect.EthernetSendCmd("NeedReceive", data, FindResource("GNSettingComplete") as string, receiveLength);
                        EthernetConnect.disconnect();
                        break;
                    case "BeepOrSetting":
                        EthernetConnect.EthernetSendCmd("NoReceive", data, null, 0);
                        EthernetConnect.disconnect();
                        break;
                }
            }
        }

        #region 網口欄位是否輸入檢查
        private bool chekckEthernetIPText()
        {
            bool isOK = false;
            if (EhternetIPTxt.Text != "")
            {
                EthernetIPAddress = EhternetIPTxt.Text;
                if (checkIPFormat(EthernetIPAddress))
                {
                    EthernetConnect.EthernetIPAddress = EthernetIPAddress;
                    isOK = true;
                }
                else
                {
                    EthernetIPAddress = null;
                    MessageBox.Show(FindResource("ErrorFormat") as string);
                }
            }
            else
            {
                MessageBox.Show(FindResource("ColumnEmpty") as string);
            }
            return isOK;
        }
        #endregion

        #region ip格式檢查
        private bool checkIPFormat(string ipString)
        {
            //create our Regular Expression object
            string pattern = @"^([1-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])(\.([0-9]|[1-9][0-9]|1[0-9][0-9]|2[0-4][0-9]|25[0-5])){3}$";
            Regex check = new Regex(pattern);
            bool isCorrect = check.IsMatch(ipString, 0);

            return isCorrect;
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

        #region byte array to hex string(須設定startIndex)
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

        #region byte array to hex string then to int
        private int byteArraytoHexStringtoInt(byte[] data)
        {
            int result = 0;
            string hexString = BitConverter.ToString(data).Replace("-", "");
            result = Convert.ToInt32(hexString, 16);
            return result;
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
                    USBConnectAndSendCmd(sendType, sendArray, length);
                    break;
                case "Ethernet":
                    bool isOK = chekckEthernetIPText(); //避免開啟預設空欄位不放在一開始檢查
                    if (isOK)
                    {
                        EthernetConnectAndSendCmd(sendType, sendArray, length);
                    }
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
                //RS232 device = new RS232() { RS232PortName = ports[i] };
                deviceList.Add(device);

                viewmodel.addDevice(device);
            }
        }
        #endregion

        //=========================USB port search/get=====================

        #region 取得usb sn and port description
        public void getUSBSNandDescription()
        {
            RegistryKey rkUsbPrint = rkLocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\DeviceClasses\\{28d78fad-5a12-11d1-ae5b-0000f803a8c2}");
            if (rkUsbPrint != null)
            {
                foreach (String usbTypePath in rkUsbPrint.GetSubKeyNames())
                {
                    //取得device instance
                    string deviceinstance = (string)rkUsbPrint.OpenSubKey(usbTypePath).GetValue("DeviceInstance");
                    //deviceinstance:USB\VID_0471&PID_0055\0003D0000000
                    // deviceinstance: USB\VID_0471 & PID_8E00\001
                    //要把vid和pid取出來，後面開usb要用
                    string sn = null;
                    string vidpid = null;
                    vidpid = deviceinstance.Split('\\')[1];
                    sn = deviceinstance.Split('\\')[2];
                    RegistryKey Device = rkUsbPrint.OpenSubKey(usbTypePath + "\\#\\Device Parameters");
                    string portdescription = null;

                    //取得port number 
                    //win10沒有連線時會沒有portNumber，此數字會=0
                    int portnumber = 0;
                    if (Device.GetValue("Port Number") != null)
                    {
                        portnumber = (int)Device.GetValue("Port Number");

                    }

                    //for winxp
                    bool isLinked = false;
                    if (OSVersion == "5.1")
                    {
                        RegistryKey Linked = rkUsbPrint.OpenSubKey(usbTypePath + "\\#\\Control");
                        if (Linked != null) //沒有抓到usb時會沒#\\Control
                        {
                            if (Linked.GetValue("Linked").ToString() == "1")
                            {
                                isLinked = true;
                            }
                            else
                            {
                                isLinked = false;
                            }
                            portdescription = (string)Device.GetValue("Port Description");//取得port description
                            Device device = new Device() { DeviceType = "usb", USBSN = sn, USBPortDescritption = portdescription, USBDeviceInstance = deviceinstance, USBVIDPID = vidpid, USBisLinked = isLinked, USBPortName = "USB" + portnumber.ToString("D3") };
                            deviceList.Add(device);
                        }
                        else
                        {
                            MessageBox.Show(FindResource("USBNotRegistYet") as string);
                        }
                    }
                    else
                    {
                        //判斷如果註冊碼資料中"Port Description"不加入此項資料
                        // win7的註冊碼中沒有這項資料
                        if (Device.GetValue("Port Description") + "" == "")
                        {   //win7
                            //ToString("D3")，代表這是三位數的格式不到三位數者自動補0.
                            Device device = new Device() { DeviceType = "usb", USBSN = sn, USBDeviceInstance = deviceinstance, USBVIDPID = vidpid, USBPortName = "USB" + portnumber.ToString("D3") };
                            deviceList.Add(device);
                        }
                        else
                        {
                            portdescription = (string)Device.GetValue("Port Description");//取得port description
                            Device device = new Device() { DeviceType = "usb", USBSN = sn, USBPortDescritption = portdescription, USBDeviceInstance = deviceinstance, USBVIDPID = vidpid, USBPortName = "USB" + portnumber.ToString("D3") };
                            deviceList.Add(device);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(FindResource("USBNotRegistYet") as string);
            }
        }
        #endregion

        #region 取得usb port name
        public void getUSBPortName()
        {
            RegistryKey rkUSBPRINT = rkLocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Enum\\USBPRINT");
            foreach (string typeName in rkUSBPRINT.GetSubKeyNames())
            {
                RegistryKey rkDevice = rkUSBPRINT.OpenSubKey(typeName);
                foreach (string portInfo in rkDevice.GetSubKeyNames())
                {
                    RegistryKey rkDevicePara = rkDevice.OpenSubKey(portInfo + "\\Device Parameters");

                    //SBARCO T4ep2  SBARCO_T4ep2_ 
                    //Gprinter GP-1624T Gprinter_GP-1624T
                    foreach (Device device in deviceList)
                    {
                        //portDescription中的空白會在typename中用_代替
                        string tmp = device.USBPortDescritption.Replace(" ", "_");
                        if (tmp.Contains(typeName))
                        {
                            device.USBPortName = (string)rkDevicePara.GetValue("PortName");
                        }
                    }
                }

            }

        }
        #endregion

        #region 判斷usb printer是否連線
        public void getUSBConnectStatus()
        {
            RegistryKey rkUsbPrint = rkLocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Services\\usbprint\\Enum");
            if (rkUsbPrint != null && (int)rkUsbPrint.GetValue("Count") != 0)
            {

                for (int i = 0; i < (int)rkUsbPrint.GetValue("Count"); i++)
                { //當Count>0代表有usb印表機連線
                  //rkUsbPrint.GetValue($"{i}")取得的資料格式為USB\VID_0999&PID_0011\001180400307 及deviceInstance的資料
                    Console.WriteLine(rkUsbPrint.GetValue(i.ToString()));
                    foreach (Device device in deviceList)
                    {

                        if (device.USBDeviceInstance == (string)rkUsbPrint.GetValue(i.ToString()))
                        {
                            device.USBisLinked = true;
                        }
                    }
                }
            }
        }
        #endregion

        #region 取得usb device所有資訊並設定選取item
        public void getUSBInfoandUpdateView()
        {

            getUSBSNandDescription();
            if (OSVersion != "5.1")
            {
                getUSBConnectStatus();
            }
            //getUSBPortName();

            foreach (Device device in deviceList)
            {
                if (device.USBisLinked)
                {
                    //如果要預設放第一筆要另外加判斷
                    try
                    {
                        //if (usbConnection.SelectedIndex == -1 && OSVersion == "5.1") //避免在XP中程式CRASH
                        //{
                        //    MessageBox.Show("USB裝置搜尋中請稍後");
                        //}

                        viewmodel.addDevice(device);

                    }
                    catch (Exception)
                    {

                        //MessageBox.Show(e.ToString());
                    }

                }
            }
        }
        #endregion

        #region 選取usb裝置
        private void DeviceSelectUSB_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Device seletedItem = DeviceSelectUSB.SelectedItem as Device;
            string usbpath = null;
            RegistryKey rkUsbPrint = rkLocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\DeviceClasses\\{28d78fad-5a12-11d1-ae5b-0000f803a8c2}");
            if (rkUsbPrint != null)
            {
                foreach (String usbTypePath in rkUsbPrint.GetSubKeyNames())
                {
                    //seletedItem!=null,要加判斷選取項目不為null不然找不到時會出現錯誤
                    //因為rkUsbPrint.GetSubKeyNames()撈出來的資料是所有有註冊的印表機資料，要比對選取項目的sn
                    //因為傳送路徑時前面##?#要改為\\?\，並且除了sn全部小寫
                    if (seletedItem != null && usbTypePath.Contains(seletedItem.USBVIDPID))
                    {
                        usbpath = "\\\\?\\" + usbTypePath.ToLower().Substring(4);
                        break;
                    }
                }
                if (usbpath != null)
                {
                    //將sn恢復為大小寫正常顯示
                    USBpath = usbpath.Replace(seletedItem.USBSN.ToLower(), seletedItem.USBSN);
                    //USBConnect.USBpath = USBpath;
                }
            }
            else
            {
                MessageBox.Show(FindResource("USBNotRegistYet") as string);
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
                viewmodel.removePort(deviceList);
                deviceList.Clear(); //清空避免重複
                getSerialPort();
                getUSBInfoandUpdateView();
                viewmodel.getDeviceObserve("rs232");
                DeviceSelectRS232.SelectedIndex = viewmodel.RS232Device.Count - 1;
                viewmodel.getDeviceObserve("usb");
                DeviceSelectUSB.SelectedIndex = viewmodel.USBDevice.Count - 1;//設定選取第一筆

            }
            handled = false;
            return IntPtr.Zero;
        }
        #endregion

        //=======================================裝置的選取==================================

        #region 選取傳輸通道
        private void ConnectType_SelectionChanged(object sender, RoutedEventArgs e)
        {
            //因為通訊測試所以全部傳輸通道要一次設定好
            viewmodel.removePort(deviceList);
            deviceList.Clear(); //清空避免重複
            getSerialPort();
            getUSBInfoandUpdateView();
            viewmodel.getDeviceObserve("rs232");
            DeviceSelectRS232.SelectedIndex = viewmodel.RS232Device.Count - 1;
            viewmodel.getDeviceObserve("usb");
            DeviceSelectUSB.SelectedIndex = viewmodel.USBDevice.Count - 1;//設定選取第一筆      
            if (USBRadio.IsChecked == true)
            {
                DeviceType = "USB";
            }
            else if (EhernetRadio.IsChecked == true)
            {
                DeviceType = "Ethernet";
            }
            else if (RS232Radio.IsChecked == true)
            {
                DeviceType = "RS232";
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
            QueryNowStatusPosition = "bottom";
            PrinterNowStatus();
        }

    }
}
