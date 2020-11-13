using G80Utility.HID;
using G80Utility.Models;
using G80Utility.Tool;
using G80Utility.ViewModels;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;
using Timer = System.Timers.Timer;

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

        //判斷通訊狀態，用來確定維修維護是否要立刻查詢
        bool isUSBConnected;
        bool isRS232Connected;
        bool isEthernetConnected;

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
        StringBuilder nvLogo_full_hex = new StringBuilder();

        //打印機實時狀態計時器
        Timer statusMonitorTimer;

        //定時發送命另計時器
        Timer sendCmdTimer;

        //app閒置Timer
        DispatcherTimer idleTimer = new DispatcherTimer();
        DateTime? lostFocusTime;

        //判斷admin是否已登入
        bool isLoginAdmin;

        //判斷SN是否已登入
        bool isLoginSN;

        // iap object
        IAP_download iap_download;

        //文件解析與下載計時 
        byte download_time = 0;
        byte hex_to_bin_time = 0;

        //委派IAP UI文字更新功能
        public delegate void setTextValueCallBack(byte index, string text);
        //IAP 文件解析與下載UI文字更新
        public setTextValueCallBack setCallBack;

        //委派IAP 設備連接狀態UI文字更新
        public delegate void setConnectStatusCallBack(bool con);

        //IAP 設備連接狀態UI文字更新
        public setConnectStatusCallBack set_connect_status;

        //文件解析與下載計時器 
        Timer timer;

        //測試連線命令(用自動斷線命令來測試是否通,，改用時實命令印表機有狀況時卡住)
        string TEST_SEND_CMD = "1F 1B 10 04 00";
        string TEST_RECEIVE_CMD = "1F-1B-1F-48-46-10-04-00";

        //判斷sn設定位置
        string SNTxtSettingPosition;

        //暫存上次選定通道選項
        int lastRS232SelectedIndex = -1;

        //傳送目前語系給bitmaptoolclass
        string nowLanguage;

        //是否更新成功
        public static bool isLoadHexSuccess,isLoadBinSuccess;
        //儲存檔案路徑
        public static string hex_file_name;
        //儲存檔案內容
        public static byte[] code_array;
        //判斷是否為 bin檔
        public static bool isBin;
        #endregion

        public G80MainWindow()
        {

            InitializeComponent();

            this.Closing += Window_Closing;

            //預設參數設定與導入參數等不可使用
            isParaSettingBtnEnabled(false);

            //語系選單default設定
            setDefaultLanguage();

            // 取得作業系統版本
            OperatingSystem os = Environment.OSVersion;
            OSVersion = os.Version.Major.ToString() + "." + os.Version.Minor.ToString();

            UIInitial();

            //頁面內容產生後註冊usb device plugin notify
            this.ContentRendered += WindowThd_ContentRendered;

            this.Closing += Window_Closing;

        }

        //window畫面生成事件
        private void WindowThd_ContentRendered(object sender, EventArgs e)
        {
            registerUSBdetect();
        }

        //window視窗關閉事件
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveLastIP();
            if (timer != null)
            {
                timer.Dispose(); //關閉計時器  
            }
            e.Cancel = false;
        }

        //app在背景事件
        override protected void OnDeactivated(EventArgs e)
        {
            lostFocusTime = DateTime.Now;
            base.OnDeactivated(e);
        }

        //app在前景事件
        protected override void OnActivated(EventArgs e)
        {
            lostFocusTime = null;
            base.OnActivated(e);
        }
        //======================UI的控制與狀態取得=================

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

            //通訊測試沒有過不能使用
            //isTabEnabled(false);

            //讀取最後一次紀錄ip
            LoadLastIP();

            collectMacAddress();

            DipContentChk();
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
            RS232Connect.CloseSerialPort(); //切換完PORT要先關閉，避免判斷已經開啟但仍連線失敗，下次無法正常RUN
            Console.WriteLine(App.Current.Properties["BaudRateSetting"]);
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

            //if (DensityModeCheckbox.IsChecked == true)
            //{
            //    Config.isDensityModeChecked = true;
            //}
            //else
            //{
            //    Config.isDensityModeChecked = false;
            //}

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

            //if (HeadCloseCutCheckbox.IsChecked == true)
            //{
            //    Config.isHeadCloseCutChecked = true;
            //}
            //else
            //{
            //    Config.isHeadCloseCutChecked = false;
            //}

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

            if (CutBeepSettingCheck.IsChecked == true)
            {
                Config.isCutBeepChecked = true;
            }
            else
            {
                Config.isCutBeepChecked = false;
            }
        }
        #endregion

        #region 判斷參數設定的核取框狀態並做設定
        private void SelectAllorClearAll(bool now, bool changeTo)
        {

            //先取得StackPanel下的所有UIElement
            foreach (StackPanel child in CommunicatePanel.Children)
            {   //沒有選擇時，Combobox的SelectedIndex=-1
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("CheckBox") && ((CheckBox)grandChild).IsChecked == now)
                    {
                        ((CheckBox)grandChild).IsChecked = changeTo;
                    }
                }
            }
            foreach (StackPanel child in PropertyColumn1.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("CheckBox") && ((CheckBox)grandChild).IsChecked == now)
                    {
                        ((CheckBox)grandChild).IsChecked = changeTo;
                    }
                }

            }
            foreach (StackPanel child in PropertyColumn2.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("CheckBox") && ((CheckBox)grandChild).IsChecked == now)
                    {
                        ((CheckBox)grandChild).IsChecked = changeTo;
                    }
                }

            }
            if (CutBeepSettingCheck.IsChecked == now)
            {
                CutBeepSettingCheck.IsChecked = changeTo;
            }

            if (CodePageSetCheckbox.IsChecked == now)
            {
                CodePageSetCheckbox.IsChecked = changeTo;
            }

            if (DIPSwitchCheckbox.IsChecked == now)
            {
                DIPSwitchCheckbox.IsChecked = changeTo;
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
            //if (ErrorTimesCheckbox.IsChecked == true)
            //{
            //    Config.iErrorTimesChecked = true;
            //}
            //else
            //{
            //    Config.iErrorTimesChecked = false;
            //}
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

        #region textbox只能輸入數字
        private void NumberValidationTextBox(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            e.Handled = regex.IsMatch(e.Text);
        }
        #endregion

        #region 16進制有勾選時textbox內容轉換
        private void HexModeCheckbox_Checked(object sender, RoutedEventArgs e)
        {
            String dataString = CmdContentTxt.Text;
            Encoding result = convertEncoding();
            dataString = ConvertStringToHex(dataString, result);
            CmdContentTxt.Text = dataString;

        }
        #endregion

        #region  16進制沒有勾選時textbox內容轉換
        private void HexModeCheckbox_Unchecked(object sender, RoutedEventArgs e)
        {
            String dataString = CmdContentTxt.Text;
            if (StringToByteArray(dataString) == null) return;//hex string中包含錯誤返回
            Encoding result = convertEncoding();
            dataString = result.GetString(StringToByteArray(dataString));
            CmdContentTxt.Text = dataString;

        }
        #endregion

        #region 檢查輸入的欄位是否為空白

        //檢查textbox設定空白
        private bool CheckTextBoxEmpty()
        {
            bool isEmpty = false;

            //先取得StackPanel下的所有UIElement
            foreach (StackPanel child in CommunicatePanel.Children) //因為包了兩層，所以要迴圈兩次
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("TextBox") && ((TextBox)grandChild).Text == "")
                    {
                        isEmpty = true;
                        break;
                    }
                }
            }
            return isEmpty;
        }

        //檢查combobox設定空白
        private bool CheckComboBoxEmpty()
        {
            bool isEmpty = false;

            //先取得StackPanel下的所有UIElement
            foreach (StackPanel child in CommunicatePanel.Children)
            {   //沒有選擇時，Combobox的SelectedIndex=-1
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("ComboBox") && ((ComboBox)grandChild).Text == "")
                    {
                        isEmpty = true;
                        break;
                    }
                }
            }
            foreach (StackPanel child in PropertyColumn1.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("ComboBox") && ((ComboBox)grandChild).Text == "")
                    {
                        isEmpty = true;
                        break;
                    }
                }

            }
            foreach (StackPanel child in PropertyColumn2.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("ComboBox") && ((ComboBox)grandChild).Text == "")
                    {
                        isEmpty = true;
                        break;
                    }
                }

            }

            if (CodePageCom.SelectedIndex == -1 || DIPSwitchCom.SelectedIndex == -1)
                isEmpty = true;

            return isEmpty;
        }
        #endregion

        #region 參數設定的btn開關控制
        public void isParaSettingBtnEnabled(bool isEnabled)
        {

            //導入與保存參數tn
            LoadParaSettingFIleBtn.IsEnabled = isEnabled;
            WriteParaSettingFIleBtn.IsEnabled = isEnabled;

            //切刀鳴叫設定btn/com
            EnableCutBeepBtn.IsEnabled = isEnabled;
            EnableCutBeepCom.IsEnabled = isEnabled;
            CutBeepTimesBtn.IsEnabled = isEnabled;
            CutBeepTimesCom.IsEnabled = isEnabled;
            CutBeepDurationgBtn.IsEnabled = isEnabled;
            CutBeepDurationgCom.IsEnabled = isEnabled;

            //代碼頁btn/com
            CodePageSetBtn.IsEnabled = isEnabled;
            CodePagePrintBtn.IsEnabled = isEnabled;
            CodePageCom.IsEnabled = isEnabled;

            //硬/軟體dip開關btn
            DIPSwitchBtn.IsEnabled = isEnabled;
            DIPSwitchCom.IsEnabled = isEnabled;

            //全局操作btn
            foreach (UIElement child in AllSetAndRead.Children)
            {
                ((Button)child).IsEnabled = isEnabled;
            }

            //通讯参数btn
            foreach (StackPanel child in CommunicatePanel.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("Button"))
                    {
                        ((Button)grandChild).IsEnabled = isEnabled;
                    }

                    if (grandChild.GetType().ToString().Contains("TextBox"))
                    {
                        ((TextBox)grandChild).IsEnabled = isEnabled;
                    }


                    if (grandChild.GetType().ToString().Contains("ComboBox"))
                    {
                        ((ComboBox)grandChild).IsEnabled = isEnabled;
                    }
                }
            }

            //打印机属性设置Column1 btn/textbox/combobox
            foreach (StackPanel child in PropertyColumn1.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("Button"))
                    {
                        ((Button)grandChild).IsEnabled = isEnabled;
                    }

                    if (grandChild.GetType().ToString().Contains("TextBox"))
                    {
                        ((TextBox)grandChild).IsEnabled = isEnabled;
                    }


                    if (grandChild.GetType().ToString().Contains("ComboBox"))
                    {
                        ((ComboBox)grandChild).IsEnabled = isEnabled;
                    }
                }

            }

            //打印机属性设置Column2 btn/textbox/combobox
            foreach (StackPanel child in PropertyColumn2.Children)
            {
                foreach (UIElement grandChild in child.Children)
                {
                    if (grandChild.GetType().ToString().Contains("Button"))
                    {
                        ((Button)grandChild).IsEnabled = isEnabled;
                    }

                    if (grandChild.GetType().ToString().Contains("TextBox"))
                    {
                        ((TextBox)grandChild).IsEnabled = isEnabled;
                    }


                    if (grandChild.GetType().ToString().Contains("ComboBox"))
                    {
                        ((ComboBox)grandChild).IsEnabled = isEnabled;
                    }
                }

            }

            //軟體dip設置
            foreach (UIElement child in DipSettingStackPanel.Children)
            {
                if (child.GetType().ToString().Contains("StackPanel")) //因為其他child並非StackPanel所以要加判斷
                {
                    foreach (UIElement grandChild in ((StackPanel)child).Children)
                    {

                        if (grandChild.GetType().ToString().Contains("Button"))
                        {
                            ((Button)grandChild).IsEnabled = isEnabled;
                        }

                        if (grandChild.GetType().ToString().Contains("CheckBox"))
                        {
                            ((CheckBox)grandChild).IsEnabled = isEnabled;
                        }
                        if (grandChild.GetType().ToString().Contains("ComboBox"))
                        {
                            ((ComboBox)grandChild).IsEnabled = isEnabled;
                        }
                    }

                }

            }

        }
        #endregion

        #region MacAddress的生成
        public byte[] collectMacAddress()
        {
            Random random = new Random();
            int mac4 = random.Next(0, 255);
            int mac5 = random.Next(0, 255);
            int mac6 = random.Next(0, 255);

            string hexMac4 = mac4.ToString("X2"); //X:16進位,2:2位數
            string hexMac5 = mac5.ToString("X2");
            string hexMac6 = mac6.ToString("X2");

            //寫入MAC Address
            byte[] sendArray = StringToByteArray(Command.MAC_ADDRESS_SETTING_HEADER + "00 47 50" + hexMac4 + hexMac5 + hexMac6);
            SetMACText.Text = "00:47:50:" + hexMac4 + ":" + hexMac5 + ":" + hexMac6;
            return sendArray;
        }
        #endregion

        #region 軟體dip值checkbox切換事件
        private void CheckBoxChanged(object sender, RoutedEventArgs e)
        {
            DipContentChk();
        }
        #endregion

        #region dip值設定的顯示控制
        private void DipContentChk()
        {
            //切刀
            if (CutterCheckBox.IsChecked == true)
            {
                CutterLabel.Content = FindResource("Disable") as string;
            }
            else if (CutterCheckBox.IsChecked == false)
            {
                CutterLabel.Content = FindResource("Enable") as string;
            }

            //蜂鳴器
            if (BeepCheckBox.IsChecked == true)
            {
                BeepLabel.Content = FindResource("Enable") as string;
            }
            else if (BeepCheckBox.IsChecked == false)
            {
                BeepLabel.Content = FindResource("Disable") as string;
            }

            //濃度
            if (DensityCheckBox.IsChecked == true)
            {
                DensityLabel.Content = FindResource("Dark") as string;
            }
            else if (DensityCheckBox.IsChecked == false)
            {
                DensityLabel.Content = FindResource("Normal") as string;
            }

            //中文禁止
            if (ChineseForbiddenCheckBox.IsChecked == true)
            {
                ChineseForbiddenLabel.Content = FindResource("YES") as string;
            }
            else if (ChineseForbiddenCheckBox.IsChecked == false)
            {
                ChineseForbiddenLabel.Content = FindResource("No") as string;
            }

            //48/42
            if (CharNumberCheckBox.IsChecked == true)
            {
                CharNumberLabel.Content = FindResource("42") as string;
            }
            else if (CharNumberCheckBox.IsChecked == false)
            {
                CharNumberLabel.Content = FindResource("48") as string;
            }

            //錢箱
            if (CashboxCheckBox.IsChecked == true)
            {
                CashboxLabel.Content = FindResource("OpenAndCut") as string;
            }
            else if (CashboxCheckBox.IsChecked == false)
            {
                CashboxLabel.Content = FindResource("OpenNotCut") as string;
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
            PrinterModuleFac.Content = moudle.Replace(" ", "") + sfvesion + "：" + date;
            PrinterModule.Content = moudle.Replace(" ", "") + sfvesion + "：" + date;

        }
        #endregion

        #region 顯示參數設置所有欄位設定內容
        public void setParaColumn(byte[] data)
        {
            string receiveData = BitConverter.ToString(data);
            //Console.WriteLine(receiveData);

            if (receiveData.Contains(Command.RE_IP_CLASSFY))
            {
                checkIsGetData(SetIPText, null, data, FindResource("SetIP") as string, false, 0);
            }

            if (receiveData.Contains(Command.RE_GATEWAY_CLASSFY))
            {
                checkIsGetData(SetGatewayText, null, data, FindResource("SetGateway") as string, false, 0);
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
                    SysStatusText.Text = FindResource("SetMacAddress") as string + FindResource("NotReadParameterYet") as string;
                }
            }

            if (receiveData.Contains(Command.RE_AUTODISCONNECT_CLASSFY))
            {
                checkIsGetData(null, AutoDisconnectCom, data, FindResource("AutomaticDisconnectTime") as string, false, 10);
            }

            if (receiveData.Contains(Command.RE_CLIENTCOUNT_CLASSFY))
            {
                checkIsGetData(null, ConnectClientCom, data, FindResource("NumberConnections") as string, true, 2);
            }

            if (receiveData.Contains(Command.RE_NETWORK_SPEED_CLASSFY))
            {
                checkIsGetData(null, EthernetSpeedCom, data, FindResource("CommunicationSpeed") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_DHCP_MODE_CLASSFY))
            {
                checkIsGetData(null, DHCPModeCom, data, FindResource("DHCP") as string, false, 3);
            }

            if (receiveData.Contains(Command.RE_USB_MODE_CLASSFY))
            {
                checkIsGetData(null, USBModeCom, data, FindResource("USBMode") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_USB_FIX_CLASSFY))
            {
                checkIsGetData(null, USBFixedCom, data, FindResource("USBPort") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_CODEPAGE_CLASSFY))
            {
                string code = receiveData.Substring(receiveData.Length - 2, 2); //取得收到hex string
                int result = Convert.ToInt32(code, 16);
                List<string> codeList = CodePage.getCodePageList();
                int index = 99; //設這個數代表沒有符合的選項就是讀取不到資料
                for (int i = 0; i < codeList.Count; i++)
                { //取得的會是個位數，前面要補0否則比對會有錯
                    string getItemCode = codeList[i].Split(':')[0];
                    if (getItemCode.Contains(result.ToString()))
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
                    setSysStatusColorAndText(FindResource("SetCodePage") as string + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                }
            }

            if (receiveData.Contains(Command.RE_LANGUAGES_CLASSFY))
            {
                checkIsGetData(null, LanguageSetCom, data, FindResource("SetLanguage") as string, true, 6);
            }

            if (receiveData.Contains(Command.RE_FONTB_CLASSFY))
            {
                checkIsGetData(null, FontBSettingCom, data, FindResource("FontB") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_CUSTOMFONT_CLASSFY))
            {
                checkIsGetData(null, CustomziedFontCom, data, FindResource("CustomizeTheFont") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_DIRECTION_CLASSFY))
            {
                checkIsGetData(null, DirectionCombox, data, FindResource("FeedingDirection") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_MOTOR_ACC_CONTROL_CLASSFY))
            {
                checkIsGetData(null, MotorAccControlCom, data, FindResource("MotorSpeed") as string, false, 1);
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
                        setSysStatusColorAndText(FindResource("MotorAcceleration") as string + FindResource("NotReadParameterYet") as string, "#FFEF7171");
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
                        setSysStatusColorAndText(FindResource("Speed") as string + FindResource("NotReadParameterYet") as string, "#FFEF7171");
                        break;
                }

            }

            //if (receiveData.Contains(Command.RE_DENSITY_MODE_CLASSFY))
            //{
            //    checkIsGetData(null, DensityModeCom, data, FindResource("DensityMode") as string, false, 1);
            //}

            if (receiveData.Contains(Command.RE_DENSITY_CLASSFY))
            {
                checkIsGetData(null, DensityCom, data, FindResource("Density") as string, true, 10);
            }

            if (receiveData.Contains(Command.RE_PAPEROUT_CLASSFY))
            {
                checkIsGetData(null, PaperOutReprintCom, data, FindResource("ReprintPaperOut") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_PAPERWIDTH_CLASSFY))
            {
                checkIsGetData(null, PaperWidthCom, data, FindResource("PaperWidth") as string, false, 1);
            }

            //if (receiveData.Contains(Command.RE_HEADCLOSE_CUT_CLASSFY))
            //{
            //    checkIsGetData(null, HeadCloseCutCom, data, FindResource("AutomaticallyCut") as string, false, 1);
            //}

            if (receiveData.Contains(Command.RE_YOFFSET_CLASSFY))
            {
                checkIsGetData(null, YOffsetCom, data, FindResource("VerticalOffsetUnit") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_MACSHOW_CLASSFY))
            {
                checkIsGetData(null, MACShowCom, data, FindResource("MacAddreessDisplay") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_QRCODE_CLASSFY))
            {
                checkIsGetData(null, QRCodeCom, data, FindResource("QRCode") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_LOGOPRINT_CLASSFY))
            {
                checkIsGetData(null, LogoPrintControlCom, data, FindResource("PrintTestLogoSetting") as string, false, 1);
            }

            if (receiveData.Contains(Command.RE_DIPSW_CLASSFY))
            {
                checkIsGetData(null, DIPSwitchCom, data, FindResource("DIPSwitch") as string, false, 1);
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
                    case "10": //115200 10取反 01，但因為這兩個bit高位要在前?所以=>1(第8位)0(第7位)
                        DIPBaudRateCom.SelectedIndex = 2;
                        break;
                    case "01": // 9600 01取反10，但因為這兩個bit高位要在前?所以=>0(第8位)1(第7位)
                        DIPBaudRateCom.SelectedIndex = 1;
                        break;
                    case "00": //38400
                        DIPBaudRateCom.SelectedIndex = 3;
                        break;
                }
            }

            if (receiveData.Contains(Command.RE_CUT_BEEP))
            {
                byte[] oneByteData = new byte[1];
                //is enabled
                oneByteData[0] = data[8];
                checkIsGetData(null, EnableCutBeepCom, oneByteData, FindResource("BeepEnable") as string, false, 1);
                //beep times
                oneByteData[0] = data[9];
                checkIsGetData(null, CutBeepTimesCom, oneByteData, FindResource("BeepTimes") as string, true, 10);
                //beep duration
                oneByteData[0] = data[10];
                checkIsGetData(null, CutBeepDurationgCom, oneByteData, FindResource("BeepDurationg") as string, true, 10);
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
                status.Append(FindResource("HeadOpen") as string + "；");
            }
            if (!bitarray[1]) // bit1 缺纸
            {
                status.Append(FindResource("PaperOut") as string + "；");

            }
            if (!bitarray[2]) // bit2 切刀错误
            {
                status.Append(FindResource("CutterError") as string + "；");

            }
            if (!bitarray[3]) //钱箱状态
            {
                status.Append(FindResource("DrawerOpen") as string + "；");
            }
            if (!bitarray[4]) //打印头超温
            {
                status.Append(FindResource("TPHOverheadted") as string + "；");

            }
            if (!bitarray[5]) //已发生错误
            {
                status.Append(FindResource("ErrorOccured") as string + "；");

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
                    status.Append(FindResource("Ready") as string + "");
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
            for (int i = 12; i < 16; i++) //打印行數
            {
                receiveArray4Bytes[i - 12] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray4Bytes);
            PrintedLinesTxt.Text = receiveInt.ToString();
            for (int i = 16; i < 20; i++) //切紙次數
            {
                receiveArray4Bytes[i - 16] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray4Bytes);
            CutPaperTimesTxt.Text = receiveInt.ToString();
            for (int i = 20; i < 22; i++) //開蓋次數
            {
                receiveArray2Bytes[i - 20] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            HeadOpenTimesTxt.Text = receiveInt.ToString();
            for (int i = 22; i < 24; i++) //缺紙次數
            {
                receiveArray2Bytes[i - 22] = data[i];
            }
            receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            PaperOutTimesTxt.Text = receiveInt.ToString();
            //for (int i = 24; i < 26; i++) //故障次數
            //{
            //    receiveArray2Bytes[i - 24] = data[i];
            //}
            //receiveInt = byteArraytoHexStringtoInt(receiveArray2Bytes);
            //ErrorTimesTxt.Text = receiveInt.ToString();
        }
        #endregion

        #region 藍牙名稱設定至畫面
        private void setBTNametoUI(byte[] data)
        {
            string btname = null;
            ////(0~7)前8個是無意義資料
            for (int i = 8; i < 24; i++)
            {
                btname += Convert.ToChar(data[i]);
            }
            btname.Replace("\0", "");
            BTName_Txt.Text = btname;
            //Console.WriteLine("bt:" + btname);
        }
        #endregion

        #region WIFI名稱設定至畫面
        private void setWIFINametoUI(byte[] data)
        {
            string WIFIname = null;
            //(0~7)前8個是無意義資料
            for (int i = 8; i < 24; i++)
            {
                WIFIname += Convert.ToChar(data[i]);      
            }
            WIFIName_Txt.Text = WIFIname;
            Console.WriteLine("WIFIname:" + WIFIname);
        }
        #endregion

        #region WIFI名稱設定至畫面
        private void setWIFIPwdtoUI(byte[] data)
        {
            string WIFIPWD = null;
            //(0~7)前8個是無意義資料
            for (int i = 8; i < 24; i++)
            {
                WIFIPWD += Convert.ToChar(data[i]);
            }
            WIFIPwd_Txt.Text = WIFIPWD;
            //Console.WriteLine("WIFIPWD:" + WIFIPWD);
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

        //通訊介面與實時查詢按鈕
        #region 通讯接口测试按鈕事件
        private void ConnectTest_Click(object sender, RoutedEventArgs e)
        {
            byte[] sendArray;
            if ((bool)rs232Checkbox.IsChecked)
            {
                if (RS232PortName != null)
                {
                    RS232Connect.CloseSerialPort(); //連線前會先判斷是否已開啟PORT    
                    bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);
                    if (!isError)
                    {
                        sendArray = StringToByteArray(Command.RS232_COMMUNICATION_TEST);
                        RS232Connect.SerialPortSendCMD("NeedReceive", sendArray, null, 8);
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                isRS232CommunicateOK(RS232Connect.mRecevieData, "communication");
                                break;
                            }
                        }
                        SendCmdFail("R");
                        RS232Connect.CloseSerialPort(); //最後關閉
                    }
                    else //serial open port failed
                    {
                        isRS232Connected = false;
                        connectFailUI(RS232ConnectImage, FindResource("CannotOpenComport") as string);
                    }
                }
                else
                {
                    MessageBox.Show(FindResource("NotSettingComport") as string);
                }

            }
            if ((bool)USBCheckbox.IsChecked)
            {
                if (USBpath != null) //先判斷是否有USBpath
                {    //已斷線，重新測試連線
                    if (USBConnect.USBHandle == -1)
                    {
                        int result = USBConnect.ConnectUSBDevice(USBpath);
                        if (result == 1)
                        {

                            byte[] sendArrayUSB = StringToByteArray(Command.USB_COMMUNICATION_TEST);
                            //USBConnectAndSendCmd("CommunicationTest", sendArray, 8);
                            USBConnect.USBSendCMD("NeedReceive", sendArrayUSB, null, 8);
                            while (!USBConnect.isReceiveData)
                            {
                                if (USBConnect.mRecevieData != null)
                                {
                                    isUSBCommunicateOK(USBConnect.mRecevieData, "communication");
                                    USBConnect.closeHandle();
                                    break;
                                }
                            }
                            SendCmdFail("U");
                        }
                        else //USB CreateFile失敗
                        {
                            isUSBConnected = false;
                            connectFailUI(USBConnectImage, FindResource("NotSettingUSBport") as string);
                        }
                    }

                }
                else
                {
                    MessageBox.Show(FindResource("NotSettingUSBport") as string);
                }
            }
            if ((bool)EthernetCheckbox.IsChecked)
            {
                bool isOK = chekckEthernetIPText();
                if (isOK)
                {
                    //checkEthernetCommunitcation();
                    int connectStatus = EthernetConnect.EthernetConnectStatus();
                    switch (connectStatus)
                    {
                        case 0: //fail
                            isEthernetConnected = false;
                            connectFailUI(EthernetConnectImage, FindResource("NotSettingEthernetport") as string);
                            break;
                        case 1: //success
                            byte[] sendArrayEthernet = StringToByteArray(Command.ETHERNET_COMMUNICATION_TEST);
                            //EthernetConnectAndSendCmd("CommunicationTest", sendArray, 8);
                            EthernetConnect.EthernetSendCmd("NeedReceive", sendArrayEthernet, null, 8);
                            while (!EthernetConnect.isReceiveData)
                            {
                                if (EthernetConnect.mRecevieData != null)
                                {
                                    isEthernetCommunicateOK(EthernetConnect.mRecevieData, "communication");
                                    break;
                                }
                            }
                            SendCmdFail("E");
                            EthernetConnect.disconnect();
                            break;
                        case 2: //timeout
                            isEthernetConnected = false;
                            connectFailUI(EthernetConnectImage, FindResource("ConnectTimeout") as string);
                            EthernetConnect.disconnect();
                            break;

                    }
                }
            }
            //一樣傳送命令前確認開通通道再關閉，避免前面測試完通道已經關閉造成錯誤
            //DifferInterfaceConnectChkAndSend("Send3Empty");
            //因為機器狀態為故障時不能傳送正常命令
            send3empty();
        }
        #endregion

        #region 重啟印表機按鈕事件
        private void Restart_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("Restart");
        }
        #endregion

        #region 查詢實時狀態按鈕事件
        private void StatusMonitorBtn_Click(object sender, RoutedEventArgs e)
        {
            string btnName = StatusMonitorBtn.Content.ToString();
            if (btnName.Contains("启动") || btnName.Contains("開啟") || btnName.Contains("Start"))
            {
                startStatusMonitorTimer();
            }
            else
            {
                stopStatusMonitorTimer();
                StatusMonitorLabel.Content = "";
            }
        }
        #endregion

        //機器序列號(通訊)
        #region 讀取機器序列號(通訊)按鈕事件
        private void ReadSNBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("RoadPrinterSN");
        }
        #endregion

        #region 設置機器序列號(通訊)按鈕事件
        private void SetSNBtn_Click(object sender, RoutedEventArgs e)
        {
            SNTxtSettingPosition = "communication";
            editSNAuthority();
        }
        #endregion

        //管理員介面
        #region 導入參數按鈕事件
        private void LoadParaSettingFIleBtn_Click(object sender, RoutedEventArgs e)
        {
            IFormatter formatter = new BinaryFormatter();
            //default讀取位置在跟exe檔同目錄夾
            try
            {
                Stream stream = new FileStream("SettingPara.sp", FileMode.Open, FileAccess.Read, FileShare.Read);
                //判斷檔案是否為空
                if (stream.Length != 0)
                {
                    ParaSettings parasetting = (ParaSettings)formatter.Deserialize(stream);
                    stream.Close();
                    readParafromFile(parasetting);
                }
                else
                {
                    MessageBox.Show(FindResource("GNSettingEmpty") as string);

                }

            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(FindResource("NotFoundSettingFile") as string);
            }
        }
        #endregion

        #region 儲存參數按鈕事件
        private void WriteParaSettingFIleBtn_Click(object sender, RoutedEventArgs e)
        {
            bool isTxtEmpty = CheckTextBoxEmpty();
            bool isComEmpty = CheckComboBoxEmpty();
            if (!isTxtEmpty && !isComEmpty) //isFormatok && 
            {
                IFormatter formatter = new BinaryFormatter();
                //default儲存位置在跟exe檔同目錄夾
                Stream stream = new FileStream("SettingPara.sp", FileMode.Create, FileAccess.Write, FileShare.None);
                formatter.Serialize(stream, saveParatoFile());
                stream.Close();
            }
            else if (isTxtEmpty || isComEmpty)
            {
                MessageBox.Show(FindResource("ColumnEmpty") as string);
            }
        }
        #endregion

        #region 管理員登入按鈕事件
        private void AdminLoginBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!isLoginAdmin)
            {
                var dialog = new Dialog();
                dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                dialog.Owner = this;
                dialog.DiaglogLabel.Content = FindResource("EnterAdminPwd") as string;
                if (dialog.ShowDialog() == true)
                {
                    if (dialog.PwdText == Config.ADMIN_PWD)
                    {
                        IdleAndLogoutTimerStart();
                        isLoginAdmin = true;
                        MessageBox.Show(FindResource("LoginSuccess") as string);
                        isParaSettingBtnEnabled(true);
                    }
                    else
                    {
                        MessageBox.Show(FindResource("PwdError") as string);
                    }
                }
            }
        }
        #endregion

        //數據傳輸按鈕
        #region 發送命令按鈕事件
        private void SendCmdBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SendCmd");
        }
        #endregion

        #region 發送換行命令按鈕事件
        private void SendEnterBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SendEmpty");
        }
        #endregion

        #region 清空命令按鈕事件
        private void ClearCmdBtn_Click(object sender, RoutedEventArgs e)
        {
            CmdContentTxt.Text = "";
        }
        #endregion

        #region 打開文件按鈕事件
        private void OepnFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog();
            openFileDlg.Multiselect = false;//该值确定是否可以选择多个文件
            openFileDlg.Title = FindResource("SelectFolder") as string;
            openFileDlg.Filter = FindResource("AllFiles") as string + "(*.txt)|*.txt";
            Nullable<bool> openDlgResult = openFileDlg.ShowDialog();
            string filepath;
            if (openDlgResult == true)
            {
                filepath = openFileDlg.FileName;
                CmdContentTxt.Text = filepath;
                CmdContentTxt.Text = System.IO.File.ReadAllText(filepath);
            }

        }
        #endregion

        #region 定時發送按鈕事件
        private void SendCmdItervalBtn_Click(object sender, RoutedEventArgs e)
        {
            string btnName = SendCmdItervalBtn.Content.ToString();
            int interval;

            if (SendCmdItervalTxt.Text == "")
            {
                interval = 0;
            }
            else
            {
                interval = Int32.Parse(SendCmdItervalTxt.Text);
            }
            if (btnName.Contains("开始") || btnName.Contains("開始") || btnName.Contains("Start"))
            {
                startSendCmdTimer(interval);
            }
            else
            {
                stopSendCmdTimer();
            }

        }
        #endregion

        //參數設置按鈕
        #region 讀取所有參數設定按鈕事件
        private void ReadAllBtn_Click(object sender, RoutedEventArgs e)
        {
            IsParaSettingChecked();
            DifferInterfaceConnectChkAndSend("readALL");
        }
        #endregion

        #region 寫入所有參數設定按鈕事件
        private void WriteAllBtn_Click(object sender, RoutedEventArgs e)
        {
            IsParaSettingChecked();
            DifferInterfaceConnectChkAndSend("sendALL");
        }
        #endregion

        #region 全選與清除按鈕事件
        //全選
        private void SelectBtn_Click(object sender, RoutedEventArgs e)
        {
            SelectAllorClearAll(false, true);
        }

        //清除全選
        private void ClearAllBtn_Click(object sender, RoutedEventArgs e)
        {
            SelectAllorClearAll(true, false);
        }
        #endregion

        #region 設定IP Address按鈕事件
        private void SetIPBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SetIP");
        }
        #endregion

        #region  設定Gateway按鈕事件
        private void SetGatewayBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SetGateway");
        }
        #endregion

        #region 設定MAC按鈕事件
        private void SetMACBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SetMAC");
        }
        #endregion

        #region 設定自動斷線時間按鈕事件
        private void AutoDisconnectBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("AutoDisconnect");
        }
        #endregion

        #region 設定網路連接數量按鈕事件
        private void ConnectClientBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("ConnectClient");
        }
        #endregion

        #region 設定網口通訊速度按鈕事件
        private void EthernetSpeedBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("EthernetSpeed");
        }
        #endregion

        #region 設定DHCP模式按鈕事件
        private void DHCPModeBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("DHCPMode");
        }
        #endregion

        #region 設定USB模式按鈕事件
        private void USBModeBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("USBMode");
        }
        #endregion

        #region 設定USB端口值按鈕事件
        private void USBFixedBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("USBFixed");
        }
        #endregion

        #region 設定代碼頁按鈕事件
        private void CodePageSetBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CodePageSet");
        }
        #endregion

        #region 打印代碼頁按鈕事件
        private void CodePagePrintBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CodePagePrint");
        }
        #endregion

        #region 設定語言按鈕事件
        private void LanguageSetBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("LanguageSet");
        }
        #endregion

        #region FontB設定按鈕事件
        private void FontBSettingBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("FontBSetting");
        }
        #endregion

        #region 設定定制字體按鈕事件
        private void CustomziedFontBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CustomziedFont");
        }
        #endregion

        #region 走紙方向按鈕事件
        private void Direction_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SetDirection");
        }
        #endregion

        #region 設定馬達加速開關按鈕事件
        private void MotorAccControlBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("MotorAccControl");
        }
        #endregion

        #region 設定馬達加速度按鈕事件
        private void AccMotorBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("AccMotor");
        }
        #endregion

        #region 設定打印速度按鈕事件
        private void PrintSpeedBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintSpeed");
        }
        #endregion

        #region 設定濃度模式按鈕事件
        private void DensityModeBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("DensityMode");
        }
        #endregion

        #region 設定濃度調節按鈕事件
        private void DensityBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("Density");
        }
        #endregion

        #region 設定紙盡重打按鈕事件
        private void PaperOutReprintBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PaperOutReprint");
        }
        #endregion

        #region 設定打印紙寬按鈕事件
        private void PaperWidthBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PaperWidth");
        }
        #endregion

        #region 設定合蓋自動切紙按鈕事件
        private void HeadCloseCutBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("HeadCloseCut");
        }
        #endregion

        #region 設定垂直移動單位按鈕事件
        private void YOffsetBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("YOffset");
        }
        #endregion

        #region 設定MAC顯示按鈕事件
        private void MACShowBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("MACShow");
        }
        #endregion

        #region 設定二維碼按鈕事件
        private void QRCodeBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("QRCode");
        }
        #endregion

        #region 設定自檢頁logo按鈕事件
        private void LogoPrintControlBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("LogoPrintControl");
        }
        #endregion

        #region 設定DIP開關按鈕事件
        private void DIPSwitchBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("DIPSwitch");
        }
        #endregion

        #region DIP值設定按鈕事件
        private void DIPSettingBtn_Click(object sender, RoutedEventArgs e)
        {
            if (DIPSwitchCom.SelectedIndex == 0) //設定為軟體dip時才寫入
            {
                DifferInterfaceConnectChkAndSend("DIPSetting");
            }
        }
        #endregion

        #region 切刀鳴叫開關按鈕事件
        private void EnableCutBeepBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CutBeepSettings");
        }
        #endregion

        #region 切刀鳴叫次數按鈕事件
        private void CutBeepTimesBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CutBeepSettings");
        }
        #endregion

        #region 切刀鳴叫時間按鈕事件
        private void CutBeepDurationgBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CutBeepSettings");
        }
        #endregion

        //維護維修按鈕
        #region 打印機維護維修tab按鈕事件
        private void MaintainTab_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {   //判斷連線成功才執行命令避免重複跳出錯誤訊息
            QueryNowStatusPosition = "maintain";
            DifferInterfaceConnectChkAndSend("PrinterInfoRead"); //第一筆確認正常以後才執行第二筆，避免重複跳出錯誤訊息
            switch (DeviceType)
            {
                case "RS232":
                    if (isRS232Connected)
                    {
                        PrinterNowStatus();
                        DifferInterfaceConnectChkAndSend("queryPrinterStatus");
                    }
                    break;
                case "USB":
                    if (isUSBConnected)
                    {

                        PrinterNowStatus();
                        DifferInterfaceConnectChkAndSend("queryPrinterStatus");
                    }
                    break;
                case "Ethernet":
                    if (isEthernetConnected)
                    {

                        PrinterNowStatus();
                        DifferInterfaceConnectChkAndSend("queryPrinterStatus");
                    }
                    break;
            }

        }
        #endregion

        #region 打印機信息查詢按鈕事件
        private void PrinterInfoReadBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrinterInfoRead");
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
            //if (!Config.iErrorTimesChecked)
            //{
            //    ErrorTimesCheckbox.IsChecked = true;
            //}

        }
        #endregion

        #region 打印機清除所有信息按鈕事件
        private void CLeanPrinterInfoBtn_Click(object sender, RoutedEventArgs e)
        {
            if (IsPrinterInfoAllChecked())
            {
                //清除所有的打印机统计信息
                DifferInterfaceConnectChkAndSend("cleanPrinterInfo");
            }
            else
            {
                DifferInterfaceConnectChkAndSend("CleanPrinterInfoOneByOne");
            }

            //清除完就要讀取打印機信息
            DifferInterfaceConnectChkAndSend("PrinterInfoRead");
        }
        #endregion

        #region 打印機狀態信息查詢按鈕事件
        private void PrinterStatusQueryBtn_Click(object sender, RoutedEventArgs e)
        {
            QueryNowStatusPosition = "maintain";
            PrinterNowStatus();
            DifferInterfaceConnectChkAndSend("queryPrinterStatus");
        }
        #endregion

        #region 打印自檢頁(短)-維護-按鈕事件
        private void PrintTest_S_Maintanin_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintTestShort");
        }
        #endregion

        #region 打印自檢頁(長)-維護-按鈕事件
        private void PrintTest_L_Maintanin_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintTestLong");
        }
        #endregion

        #region 打印均勻測試-維護-按鈕事件
        private void PrintTest_EVEN_Maintanin_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintEvenTest");
        }
        #endregion

        #region 蜂鳴器測試-維護-按鈕事件
        private void BeepTest_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("BeepTest");
        }
        #endregion

        #region 下踢錢箱-維護-按鈕事件
        private void OpenCashBox_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("OpenCashBox");
        }
        #endregion

        #region 連續切紙-維護-按鈕事件
        private void CutTimes_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CutTimesMaintain");
        }
        #endregion

        #region 指令測試-維護-按鈕事件
        private void CMDTest_Maintanin_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CMDTestMaintain");
        }
        #endregion

        #region SDRAM測試按鈕事件
        private void SDRAMTestBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("SDRAMTest");
        }
        #endregion

        #region Flash測試按鈕事件
        private void FlashTestBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("FlashTest");
        }
        #endregion

        #region 取得藍牙名稱按鈕事件
        private void BTName_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("ReadBT");
        }
        #endregion

        #region 取得WIFI名稱按鈕事件
        private void WIFINameLoad_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("ReadWIFIName");
        }
        #endregion

        #region 寫入WIFI名稱按鈕事件
        private void WIFINameWirte_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("WriteWIFIName");
        }
        #endregion

        #region 取得WIFI密碼按鈕事件
        private void wIFIPwdLoad_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("ReadWIFIPwd");
        }
        #endregion

        #region 寫入WIFI密碼按鈕事件
        private void wIFIPwdWrite_Btn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("WriteWIFIPwd");
        }
        #endregion
        //工廠生產按鈕
        #region 打印自檢頁(短)-工廠-按鈕事件
        private void PrintTest_S_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintTestShort");
        }
        #endregion

        #region 打印自檢頁(長)-工廠-按鈕事件
        private void PrintTest_L_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintTestLong");
        }
        #endregion

        #region 打印均勻測試-工廠-按鈕事件
        private void PrintTest_EVEN_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintEvenTest");
        }
        #endregion

        #region 蜂鳴器測試-工廠-按鈕事件
        private void BeepTestBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("BeepTest");
        }
        #endregion

        #region 下踢錢箱-工廠-按鈕事件
        private void OpenCashBoxBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("OpenCashBox");
        }
        #endregion

        #region 連續切紙-工廠-按鈕事件
        private void CutTimesBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CutTimesFactory");
        }
        #endregion

        #region 指令測試-工廠-按鈕事件
        private void CMDTestBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("CMDTestFactory");
        }
        #endregion

        #region 讀取機器序列號(工廠)按鈕事件
        private void RoadPrinterSNFacBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("RoadPrinterSN");
        }
        #endregion

        #region 設置機器序列號(工廠)按鈕事件
        private void SetPrinterSNFacBtn_Click(object sender, RoutedEventArgs e)
        {
            SNTxtSettingPosition = "factory";
            editSNAuthority();
        }
        #endregion

        #region 出廠設置按鈕事件
        private void FactoryDefaultBtn_Click(object sender, RoutedEventArgs e)
        {
            //清除所有的打印机统计信息
            DifferInterfaceConnectChkAndSend("cleanPrinterInfo");

            //根据参数设置界面的复选框进行所有参数的下载
            IsParaSettingChecked();
            DifferInterfaceConnectChkAndSend("sendALL");
            Thread.Sleep(1000); //要等一下不然命令會來不及送

            //发送打印自检页（长）命令
            DifferInterfaceConnectChkAndSend("PrintTestLong");
        }
        #endregion

        //NVLogo按鈕
        #region 打印logo按鈕事件
        private void PrintLogoBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("PrintLogo");
        }
        #endregion

        #region 清除logo下載按鈕事件
        private void ClearLogoBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("ClearLogo");
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
            nvLogo_full_hex.Clear(); //每次開啟要清空，不然下載時會一直重複
            nvLogo_full_hex.Append("1C71");
            if ((bool)openFileDialog.ShowDialog())
            {
                if (fileNameArray == null) //第一次開啟
                {
                    fileNameArray = openFileDialog.FileNames;
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
                    if (!BitmapTool.checkBitmapRange(fileNameArray[i], bmp.Width, bmp.Height, nowLanguage)) //判斷寬高是否超過標準
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
                    Bitmap bmp = BitmapTool.ToGray(BitmapTool.BitmapImageToBitmap(bmpImg), 1);

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
            OpenImgNoTxt.Text = "";
        }
        #endregion

        #region 下載圖片集按鈕事件
        private void DonwaldLogoBtn_Click(object sender, RoutedEventArgs e)
        {
            DifferInterfaceConnectChkAndSend("DonwaldLogo");
        }
        #endregion

        //升級程序按鈕
        #region 升級程序tab按鈕事件
        private void FWUpdateTab_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {   //初始化
            UIintial();

            //iap初始化
            if (iap_download != null)
            {
                iap_download.Close();
            }
            iap_download = new IAP_download();
            iap_download.Initial();
            iap_download.isConnectedFunc = device_connect_status;
            set_connect_status += device_connect_status_ui;

            //啟動ipa計時器   
            timer = new Timer();
            timer.Interval = 1000;
            timer.Elapsed += timer_1s_Tick;
            timer.Start();

            //未連接設備不可以下載
            if (DeviceStatusTxt.Text == "" || DeviceStatusTxt.Text == FindResource("DeviceDisconnected") as string) 
            {
                openfileAndDownloadUIControl(false);
            }
            else {
                openfileAndDownloadUIControl(true);
            }
        }
        #endregion

        #region 重新連接打印機按鈕事件
        private void ReconnectBtn_Click(object sender, RoutedEventArgs e)
        {
            UIintial();
            //iap初始化
            if (iap_download != null) {
                iap_download.Close();
            }
            iap_download = new IAP_download(); 
            iap_download.Initial();
            iap_download.isConnectedFunc = device_connect_status;        

            //啟動ipa計時器   
            if (timer != null)
            {
                timer.Dispose();  //避免重複設定
                timer.Close();                
            }
            timer = new Timer();
            timer.Interval = 1000;
            timer.Elapsed += timer_1s_Tick;
            timer.Start();

            if (!isLoadBinSuccess && !isLoadHexSuccess)
            {
                set_connect_status += device_connect_status_ui;
            }

            //未連接設備不可以下載
            if (DeviceStatusTxt.Text == "" || DeviceStatusTxt.Text == FindResource("DeviceDisconnected") as string)
            {
                openfileAndDownloadUIControl(false);
            }
            else
            {
                openfileAndDownloadUIControl(true);
            }

        }
        #endregion

        #region 開啟FW檔案按鈕事件
        private void OpenFWfileBtn_Click(object sender, EventArgs e)
        {
            string file_name = "";
            string ext_name = "";
            isBin = false;
            isLoadHexSuccess = false;//打開文件需要重新解析，這邊把此變數恢復預設
            isLoadBinSuccess = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = FindResource("SelectFolder") as string;
            dialog.Filter = FindResource("AllFiles") as string + "(*.hex)|*.hex|" + FindResource("AllFiles") as string + "(*.bin)|*.bin";
            if (dialog.ShowDialog() == true)
            {
                file_name = dialog.FileName;
                ext_name = Path.GetExtension(file_name);
                if (ext_name.Equals(".bin")) {
                    isBin = true;
                }
                FilePathTxt.Text = file_name;
                iap_download.hex_file_name = file_name;
                updata_file_button_Click(sender, e);
                DownloadFWBtn.IsEnabled = true;
            }
            else
            {
                FilePathTxt.Text = FindResource("Status") as string + FindResource("FailedOpen") as string;
                DownloadFWBtn.IsEnabled = false;

            }
        }

        #endregion

        #region 更新FW程序按鈕事件
        private void DownloadFWBtn_Click(object sender, RoutedEventArgs e)
        {
            download_time = 0;
            hex_to_bin_time = 0;
            //实例化回调
            setCallBack = new setTextValueCallBack(updata_ui_status_text);
            if (isLoadHexSuccess || isLoadBinSuccess || isBin)
                updata_file_button_Click(sender, e);
            //创建一个线程去执行这个方法:创建的线程默认是前台线程
            Thread thread = new Thread(new ParameterizedThreadStart(iap_download.download_code));
            thread.IsBackground = true;
            thread.Start(this);
        }
        #endregion
        //========================傳輸三個空白===========================

        public void send3empty()
        {
            byte[] empty3Array = { 0x0a, 0x0a, 0x0a };
            byte[] cut = StringToByteArray(Command.CUT_TIMES);
            switch (DeviceType)
            {
                case "RS232":
                    if (RS232PortName != null) //先判斷是否有get prot name
                    {
                        bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);
                        if (!isError)
                        {
                            RS232Connect.SerialPortSendCMD("NoReceive", empty3Array, null, 0);                       
                            RS232Connect.SerialPortSendCMD("NoReceive", cut, null, 0);
                            SendCmdFail("R"); //避免傳輸時突然有問題
                            RS232Connect.CloseSerialPort(); //最後關閉
                            Thread.Sleep(1000); //串口要停一下避免等下讀取又誤
                        }
                        else //serial open port failed
                        {
                            stopStatusMonitorTimer();
                            stopSendCmdTimer();
                            connectFailUI(RS232ConnectImage, FindResource("CannotOpenComport") as string);
                        }
                    }
                    else
                    {
                        stopSendCmdTimer();
                        stopStatusMonitorTimer();
                        MessageBox.Show(FindResource("NotSettingComport") as string);
                    }
                    break;
                case "USB":
                    if (USBpath != null) //先判斷是否有USBpath
                    {
                        int result = USBConnect.ConnectUSBDevice(USBpath);
                        if (result == 1)
                        {
                            USBConnect.USBSendCMD("NoReceive", empty3Array, null, 0);
                            USBConnect.USBSendCMD("NoReceive", cut, null, 0);
                            USBConnect.closeHandle();
                            SendCmdFail("U");//避免傳輸時突然有問題
                        }
                        else
                        {
                            stopSendCmdTimer();
                            stopStatusMonitorTimer();
                            connectFailUI(USBConnectImage, FindResource("NotSettingUSBport") as string);
                        }
                    }
                    else
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        MessageBox.Show(FindResource("NotSettingUSBport") as string);
                    }
                    break;
                case "Ethernet":
                    bool isOK = chekckEthernetIPText();
                    if (isOK)
                    {
                        int connectStatus = EthernetConnect.EthernetConnectStatus();
                        switch (connectStatus)
                        {
                            case 0: //fail
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(EthernetConnectImage, FindResource("NotSettingEthernetport") as string);
                                break;
                            case 1: //success                           
                                EthernetConnect.EthernetSendCmd("NoReceive", empty3Array, null, 0);
                                EthernetConnect.EthernetSendCmd("NoReceive", cut, null, 0);
                                EthernetConnect.disconnect();
                                break;
                            case 2: //timeout
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(EthernetConnectImage, FindResource("ConnectTimeout") as string);
                                break;
                        }
                    }
                    break;
            }
        }

        //========================紀錄和讀取最後一次設定IP=================
        #region 紀錄最後一次輸入IP
        public void SaveLastIP()
        {
            string IP = EhternetIPTxt.Text;
            //建立註冊機碼目錄
            createSNRegistry();
            setRegistry("IP", IP);
        }
        #endregion

        #region 讀取最後一次輸入IP
        public void LoadLastIP()
        {   //首次使用要建立註冊機碼目錄
            createSNRegistry();
            string IP = getRegistry("IP");
            EhternetIPTxt.Text = IP;

        }
        #endregion
        //============================判斷sn設定權限======================

        public void editSNAuthority()
        {
            if (isLoginSN)
            {
                DifferInterfaceConnectChkAndSend("SetPrinterSN");
            }
            else
            {
                var dialog = new Dialog();
                dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                dialog.Owner = this;
                dialog.DiaglogLabel.Content = FindResource("EnterSnPwd") as string;
                if (dialog.ShowDialog() == true)
                {
                    if (dialog.PwdText == Config.EDIT_SN_PWD)
                    {
                        isLoginSN = true;
                        MessageBox.Show(FindResource("LoginSuccess") as string);
                        DifferInterfaceConnectChkAndSend("SetPrinterSN");
                        //SetPrinterSN();
                    }
                    else
                    {
                        MessageBox.Show(FindResource("PwdError") as string);
                    }

                }
            }
        }

        //============================儲存和寫入參數檔====================

        #region 儲存參數至檔案
        private ParaSettings saveParatoFile()
        {
            ParaSettings parasetting = new ParaSettings();
            parasetting.IpAddress = SetIPText.Text;
            parasetting.Gateway = SetGatewayText.Text;
            parasetting.MacAddress = SetMACText.Text;
            parasetting.AutoDisconnectIndex = AutoDisconnectCom.SelectedIndex;
            parasetting.ConnectClientIndex = ConnectClientCom.SelectedIndex;
            parasetting.EthernetSpeedIndex = EthernetSpeedCom.SelectedIndex;
            parasetting.DHCPModeIndex = DHCPModeCom.SelectedIndex;
            parasetting.USBModeIndex = USBModeCom.SelectedIndex;
            parasetting.USBFixedIndex = USBFixedCom.SelectedIndex;
            parasetting.CodePageSetIndex = CodePageCom.SelectedIndex;
            parasetting.LanguageSetIndex = LanguageSetCom.SelectedIndex;
            parasetting.FontBSettingtIndex = FontBSettingCom.SelectedIndex;
            parasetting.CustomziedFontIndex = CustomziedFontCom.SelectedIndex;
            parasetting.SetDirectionIndex = DirectionCombox.SelectedIndex;
            parasetting.MotorAccControlIndex = MotorAccControlCom.SelectedIndex;
            parasetting.AccMotorIndex = AccMotorCom.SelectedIndex;
            parasetting.PrintSpeedIndex = PrintSpeedCom.SelectedIndex;
            //parasetting.DensityModeIndex = DensityModeCom.SelectedIndex;
            parasetting.DensityIndex = DensityCom.SelectedIndex;
            parasetting.PaperOutReprintIndex = PaperOutReprintCom.SelectedIndex;
            parasetting.PaperWidthIndex = PaperWidthCom.SelectedIndex;
            //parasetting.HeadCloseCutIndex = HeadCloseCutCom.SelectedIndex;
            parasetting.YOffsetIndex = YOffsetCom.SelectedIndex;
            parasetting.MACShowIndex = MACShowCom.SelectedIndex;
            parasetting.QRCodeIndex = QRCodeCom.SelectedIndex;
            parasetting.LogoPrintControlIndex = LogoPrintControlCom.SelectedIndex;

            parasetting.CutBeepEnable = EnableCutBeepCom.SelectedIndex;
            parasetting.CutBeepTimes = CutBeepTimesCom.SelectedIndex;
            parasetting.CutBeepDuration = CutBeepDurationgCom.SelectedIndex;

            parasetting.DIPSwitchIndex = DIPSwitchCom.SelectedIndex;
            if (CutterCheckBox.IsChecked == true)
            {
                parasetting.CutterCheck = false;
            }
            else
            {
                parasetting.CutterCheck = true;
            }
            if (BeepCheckBox.IsChecked == true)
            {
                parasetting.BeepCheck = false;
            }
            else
            {
                parasetting.BeepCheck = true;
            }
            if (DensityCheckBox.IsChecked == true)
            {
                parasetting.DensityCheck = false;
            }
            else
            {
                parasetting.DensityCheck = true;
            }
            if (ChineseForbiddenCheckBox.IsChecked == true)
            {
                parasetting.ChineseForbiddenCheck = false;
            }
            else
            {
                parasetting.ChineseForbiddenCheck = true;
            }
            if (CharNumberCheckBox.IsChecked == true)
            {
                parasetting.CharNumberCheck = false;
            }
            else
            {
                parasetting.CharNumberCheck = true;
            }
            if (CashboxCheckBox.IsChecked == true)
            {
                parasetting.CashboxCheck = false;
            }
            else
            {
                parasetting.CashboxCheck = true;
            }

            parasetting.DIPBaudRateComIndex = DIPBaudRateCom.SelectedIndex;

            return parasetting;
        }
        #endregion

        #region 從檔案讀取參數
        private void readParafromFile(ParaSettings parasetting)
        {
            SetIPText.Text = parasetting.IpAddress;
            SetGatewayText.Text = parasetting.Gateway;
            SetMACText.Text = parasetting.MacAddress;
            AutoDisconnectCom.SelectedIndex = parasetting.AutoDisconnectIndex;
            ConnectClientCom.SelectedIndex = parasetting.ConnectClientIndex;
            EthernetSpeedCom.SelectedIndex = parasetting.EthernetSpeedIndex;
            DHCPModeCom.SelectedIndex = parasetting.DHCPModeIndex;
            USBModeCom.SelectedIndex = parasetting.USBModeIndex;
            USBFixedCom.SelectedIndex = parasetting.USBFixedIndex;
            CodePageCom.SelectedIndex = parasetting.CodePageSetIndex;
            LanguageSetCom.SelectedIndex = parasetting.LanguageSetIndex;
            FontBSettingCom.SelectedIndex = parasetting.FontBSettingtIndex;
            CustomziedFontCom.SelectedIndex = parasetting.CustomziedFontIndex;
            DirectionCombox.SelectedIndex = parasetting.SetDirectionIndex;
            MotorAccControlCom.SelectedIndex = parasetting.MotorAccControlIndex;
            AccMotorCom.SelectedIndex = parasetting.AccMotorIndex;
            PrintSpeedCom.SelectedIndex = parasetting.PrintSpeedIndex;
            //DensityModeCom.SelectedIndex = parasetting.DensityModeIndex;
            DensityCom.SelectedIndex = parasetting.DensityIndex;
            PaperOutReprintCom.SelectedIndex = parasetting.PaperOutReprintIndex;
            PaperWidthCom.SelectedIndex = parasetting.PaperWidthIndex;
            //HeadCloseCutCom.SelectedIndex = parasetting.HeadCloseCutIndex;
            YOffsetCom.SelectedIndex = parasetting.YOffsetIndex;
            MACShowCom.SelectedIndex = parasetting.MACShowIndex;
            QRCodeCom.SelectedIndex = parasetting.QRCodeIndex;
            LogoPrintControlCom.SelectedIndex = parasetting.LogoPrintControlIndex;

            EnableCutBeepCom.SelectedIndex = parasetting.CutBeepEnable;
            CutBeepTimesCom.SelectedIndex = parasetting.CutBeepTimes;
            CutBeepDurationgCom.SelectedIndex = parasetting.CutBeepDuration;

            DIPSwitchCom.SelectedIndex = parasetting.DIPSwitchIndex;
            if (parasetting.CutterCheck == false)
            {
                CutterCheckBox.IsChecked = true;
            }
            else
            {
                CutterCheckBox.IsChecked = false;
            }
            if (parasetting.BeepCheck == false)
            {
                BeepCheckBox.IsChecked = true;
            }
            else
            {
                BeepCheckBox.IsChecked = false;
            }
            if (parasetting.DensityCheck == false)
            {
                DensityCheckBox.IsChecked = true;
            }
            else
            {
                DensityCheckBox.IsChecked = false;
            }
            if (parasetting.ChineseForbiddenCheck == false)
            {
                ChineseForbiddenCheckBox.IsChecked = true;
            }
            else
            {
                ChineseForbiddenCheckBox.IsChecked = false;
            }
            if (parasetting.CharNumberCheck == false)
            {
                CharNumberCheckBox.IsChecked = true;
            }
            else
            {
                CharNumberCheckBox.IsChecked = false;
            }
            if (parasetting.CashboxCheck == false)
            {
                CashboxCheckBox.IsChecked = true;
            }
            else
            {
                CashboxCheckBox.IsChecked = false;
            }
            DIPBaudRateCom.SelectedIndex = parasetting.DIPBaudRateComIndex;
        }
        #endregion

        //============================數據傳輸發送命令功能=================

        #region 取得數據傳輸欄位資料
        public byte[] getTransferData()
        {
            String dataString = CmdContentTxt.Text;
            if (CmdContentTxt.Text == "")
            {
                CmdContentTxt.Text = dataString = FindResource("DefaultText") as string;
            }
            byte[] dataArray = null;
            if (HexModeCheckbox.IsChecked == true)
            {
                dataString = dataString.Replace(" ", "");
                if (StringToByteArray(dataString) == null) return dataArray;//hex string中包含錯誤返回
                dataArray = StringToByteArray(dataString);
            }
            else
            {
                Encoding result = convertEncoding();
                dataArray = result.GetBytes(dataString);
            }
            return dataArray;
        }
        #endregion

        #region 只發送一次命令
        public void SendCmd()
        {
            byte[] sendArray = getTransferData();

            switch (DeviceType)
            {
                case "RS232":
                    if (isRS232Connected)
                    {
                        SendCmd(sendArray, "BeepOrSetting", 0);
                    }
                    else
                    {
                        stopSendCmdTimer();
                    }
                    break;
                case "USB":
                    if (isUSBConnected)
                    {
                        SendCmd(sendArray, "BeepOrSetting", 0);
                    }
                    else
                    {
                        stopSendCmdTimer();
                    }
                    break;
                case "Ethernet":
                    if (isEthernetConnected)
                    {
                        SendCmd(sendArray, "BeepOrSetting", 0);
                    }
                    else
                    {
                        stopSendCmdTimer();
                    }
                    break;
            }
        }
        #endregion

        #region 定時發送命令 
        //避免卡住不做通訊測試
        public void SendCmdWithInterval()
        {
            byte[] sendArray = getTransferData();

            switch (DeviceType)
            {
                case "RS232":
                    if (RS232PortName != null) //先判斷是否有get prot name
                    {    //已斷線，重新測試連線
                        if (!RS232Connect.RS232ConnectStatus())
                        {
                            bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);
                            if (!isError)
                            {
                                RS232Connect.SerialPortSendCMD("NoReceive", sendArray, null, 0);
                                SendCmdFail("R"); //避免傳輸時突然有問題
                            }
                            else //serial open port failed
                            {
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(RS232ConnectImage, FindResource("CannotOpenComport") as string);
                            }
                        }
                        else
                        {
                            RS232Connect.SerialPortSendCMD("NoReceive", sendArray, null, 0);
                            SendCmdFail("R");//避免傳輸時突然有問題
                        }
                        //SendCmdFail("R");
                    }
                    else
                    {
                        stopSendCmdTimer();
                        stopStatusMonitorTimer();
                        MessageBox.Show(FindResource("NotSettingComport") as string);
                    }
                    break;
                case "USB":
                    if (USBpath != null) //先判斷是否有USBpath
                    {    //已斷線，重新測試連線
                        if (USBConnect.USBHandle == -1)
                        {
                            int result = USBConnect.ConnectUSBDevice(USBpath);
                            if (result == 1)
                            {
                                USBConnect.USBSendCMD("NoReceive", sendArray, null, 0);
                                SendCmdFail("U");//避免傳輸時突然有問題
                            }
                            else
                            {
                                stopSendCmdTimer();
                                stopStatusMonitorTimer();
                                connectFailUI(USBConnectImage, FindResource("NotSettingUSBport") as string);
                            }
                        }
                        else
                        {
                            USBConnect.USBSendCMD("NoReceive", sendArray, null, 0);
                            SendCmdFail("U");//避免傳輸時突然有問題
                        }

                    }
                    else
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        MessageBox.Show(FindResource("NotSettingUSBport") as string);
                    }
                    break;
                case "Ethernet":
                    bool isOK = chekckEthernetIPText();
                    if (isOK)
                    {
                        int connectStatus = EthernetConnect.EthernetConnectStatus();
                        switch (connectStatus)
                        {
                            case 0: //fail
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(EthernetConnectImage, FindResource("NotSettingEthernetport") as string);
                                break;
                            case 1: //success
                                EthernetConnect.EthernetSendCmd("NeedReceive", sendArray, null, 9);
                                SendCmdFail("E");
                                break;
                            case 2: //timeout
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(EthernetConnectImage, FindResource("ConnectTimeout") as string);
                                break;
                        }
                    }
                    break;
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
            byte[] sendArray = collectMacAddress();
            //Random random = new Random();
            //int mac4 = random.Next(0, 255);
            //int mac5 = random.Next(0, 255);
            //int mac6 = random.Next(0, 255);

            //string hexMac4 = mac4.ToString("X2"); //X:16進位,2:2位數
            //string hexMac5 = mac5.ToString("X2");
            //string hexMac6 = mac6.ToString("X2");

            ////寫入MAC Address
            //sendArray = StringToByteArray(Command.MAC_ADDRESS_SETTING_HEADER + "00 47 50" + hexMac4 + hexMac5 + hexMac6);
            //SetMACText.Text = "00:47:50:" + hexMac4 + ":" + hexMac5 + ":" + hexMac6;
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
                    case 6: //6min=>33 06
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 06");
                        break;
                    case 7: //7min=>33 07
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 07");
                        break;
                    case 8: //8min=>33 08
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 08");
                        break;
                    case 9: //9min=>33 09
                        sendArray = StringToByteArray(Command.NETWORK_AUTODICONNECTED_SETTING_HEADER + "33 09");
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
                HexCode = Int32.Parse(HexCode).ToString("X2");
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

        /*
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
        */

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

        /*
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
        */

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

                if (DIPSwitchCom.SelectedIndex == 0) //software
                {
                    sendArray = StringToByteArray(Command.DIP_OFF_SETTING);
                    DIPGroupBox.IsEnabled = true;
                }
                else if (DIPSwitchCom.SelectedIndex == 1) //hardware
                {
                    sendArray = StringToByteArray(Command.DIP_ON_SETTING);
                    DIPGroupBox.IsEnabled = false;
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
                case 1: //9600 01取反10，但因為這兩個bit高位要在前?所以=>0(第8位)1(第7位)
                    dipArray.Set(6, false);
                    dipArray.Set(7, true);
                    break;
                case 2: //115200 10取反 01，但因為這兩個bit高位要在前?所以=>1(第8位)0(第7位)
                    dipArray.Set(6, true);
                    dipArray.Set(7, false);
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
            byte[] sendArray = StringToByteArray(Command.DIP_VALUE_SETTING_HEADER);
            Array.Resize(ref sendArray, sendArray.Length + 1);
            sendArray[sendArray.Length - 1] = bytes[0];
            SendCmd(sendArray, "BeepOrSetting", 0);
            MessageBox.Show(FindResource("BaudChangeWarning") as string);
        }
        #endregion

        #region 切刀鳴叫開關/次數/時間
        private void CutBeepSettings()
        {
            if (EnableCutBeepCom.SelectedIndex != -1 || CutBeepTimesCom.SelectedIndex != -1 || CutBeepDurationgCom.SelectedIndex != -1)
            {
                List<byte> tempList = new List<byte>();
                tempList = StringToByteArray(Command.SET_CUT_BEEP).ToList();
                if (EnableCutBeepCom.SelectedIndex == 0)
                {
                    tempList.Add(0x00);
                }
                else if (EnableCutBeepCom.SelectedIndex == 1)
                {
                    tempList.Add(0x01);
                }
                int times = CutBeepTimesCom.SelectedIndex + 1;
                tempList.Add(Convert.ToByte(times));
                int Duration = CutBeepDurationgCom.SelectedIndex + 1;
                tempList.Add(Convert.ToByte(Duration));
                SendCmd(tempList.ToArray(), "BeepOrSetting", 0);

            }
            else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
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

            //if (Config.isDensityModeChecked)
            //{
            //    sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 26 01");
            //    SendCmd(sendArray, "ReadPara", 9);
            //}

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

            //if (Config.isHeadCloseCutChecked)
            //{
            //    sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 17 01");
            //    SendCmd(sendArray, "ReadPara", 9);
            //}

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
            if (DIPSwitchCom.SelectedIndex == 0) //設定為軟體dip時才讀取
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 35 01");
                SendCmd(sendArray, "ReadPara", 9);
            }

            //切刀鳴叫設定讀取
            if (Config.isCutBeepChecked)
            {
                sendArray = StringToByteArray(Command.READ_ALL_HEADER + "31 32 01");
                SendCmd(sendArray, "ReadPara", 11);
            }
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

            //if (Config.isDensityModeChecked)
            //{
            //    DensityMode();
            //}

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

            //if (Config.isHeadCloseCutChecked)
            //{
            //    HeadCloseCut();
            //}

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

            if (DIPSwitchCom.SelectedIndex == 0) //設定為軟體dip時才寫入
            {
                DIPSetting();
            }

            if (Config.isCutBeepChecked)
            {
                CutBeepSettings();
            }
        }
        #endregion

        //=============================維護維修功能=============================

        #region 讀取打印機所有統計信息
        private void PrinterInfoRead()
        {
            byte[] sendArray = StringToByteArray(Command.READ_PRINTINFO);
            SendCmd(sendArray, "ReadPrinterInfo", 30);
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

        #region 讀取BT名稱
        private void LoadBTName()
        {
            byte[] sendArray = StringToByteArray(Command.BT_LAOD);
            SendCmd(sendArray, "ReadBTName", 24);
        }
        #endregion

        #region 讀取WIFI名稱
        private void LoadWIFIName()
        {
            byte[] sendArray = StringToByteArray(Command.WIFI_NAME_LOAD);
            SendCmd(sendArray, "ReadWIFIName", 24);
        }
        #endregion

        #region 讀取WIFI密碼
        private void LoadWIFIPwd()
        {
            byte[] sendArray = StringToByteArray(Command.WIFI_PWD_LOAD);
            SendCmd(sendArray, "ReadWIFIPwd", 24);
        }
        #endregion

        #region 設定WIFI名稱
        private void SetWIFIName()
        {
            byte[] sendArray = null;
            string name;
            if (WIFIName_Txt.Text != "")
            {
                name = WIFIName_Txt.Text.Replace("\0", "");

                String result = ConvertStringToHex(name);
                if (result != null)
                {
                    sendArray = StringToByteArray(Command.WIFI_NAME_SET + result);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }
                else {
                    MessageBox.Show(FindResource("WIFINameTooLarge") as string);
                }

            }
            else
            {
                MessageBox.Show(FindResource("ColumnEmpty") as string);
            }
        }
        #endregion

        #region 設定WIFI密碼
        private void SetWIFIPWD()
        {
            byte[] sendArray = null;
            string pwd;
            if (WIFIPwd_Txt.Text != "")
            {
                pwd = WIFIPwd_Txt.Text.Replace("\0","");

                String result = ConvertStringToHex(pwd);
                if (result != null)
                {
                    sendArray = StringToByteArray(Command.WIFI_PWD_SET + result);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                }
                else
                {
                    MessageBox.Show(FindResource("WIFIPwdTooLarge") as string);
                }

            }
            else
            {
                MessageBox.Show(FindResource("ColumnEmpty") as string);
            }
        }
        #endregion

        //==============================工廠生產功能=============================

        #region 打印自檢頁
        //private void PrintTest(string printType)
        //{
        //    byte[] sendArray = null;

        //    if (printType == "short")
        //    {
        //        sendArray = StringToByteArray(Command.PRINT_TEST_SHORT);
        //    }
        //    else if (printType == "long")
        //    {
        //        sendArray = StringToByteArray(Command.PRINT_TEST_LONG);
        //    }
        //    SendCmd(sendArray, "BeepOrSetting", 0);

        //}
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

        #region 指令測試
        private void CMDTestFactory()
        {
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
        }

        private void CMDTestMaintain()
        {
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
        private void SetPrinterSN()
        {
            string sn_reg = null;
            string sn_input = null;
            //建立註冊機碼目錄
            createSNRegistry();
            switch (SNTxtSettingPosition)
            {
                case "factory":
                    sn_input = PrinterSNFacTxt.Text;
                    break;
                case "communication":
                    sn_input = PrinterSNTxt.Text;
                    break;
            }
            //初次未記錄,抓取user輸入值
            if (getRegistry("SN") == null)
            {
                if (printerSNInputChk())
                {
                    sn_reg = sn_input;
                    snWriteInPrinter(sn_reg);
                }
                else
                {
                    MessageBox.Show(FindResource("SNFormatError") as string);
                }
            }
            else //已經有紀錄
            {
                sn_reg = getRegistry("SN");

                if (sn_input != sn_reg)//註冊與輸入的不同要重新寫入
                {
                    if (printerSNInputChk())
                    {
                        sn_reg = sn_input;
                        snWriteInPrinter(sn_reg);
                    }
                    else
                    {
                        MessageBox.Show(FindResource("SNFormatError") as string);
                    }
                }
                else
                { //註冊與輸入的相同直接+1
                    if (printerSNInputChk())
                    {
                        string number = sn_reg.Substring(10, 6);
                        sn_reg = sn_reg.Remove(10, 6);
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
                        sn_reg = sn_reg + number;
                        snWriteInPrinter(sn_reg);
                    }
                    else
                    {
                        MessageBox.Show(FindResource("SNFormatError") as string);
                    }
                }

            }
        }
        #endregion

        #region 寫入序號到打印機
        private void snWriteInPrinter(string sn_reg)
        {
            if (sn_reg != "" && sn_reg.Length == 16)
            {
                byte[] snArray = Encoding.Default.GetBytes(sn_reg);
                byte[] sendArray = StringToByteArray(Command.SN_SETTING_HEADER);
                int sendLen = sendArray.Length;
                int snLen = snArray.Length;
                Array.Resize(ref sendArray, sendLen + snLen);
                for (int i = sendLen; i < sendLen + snLen; i++)
                {
                    sendArray[i] = snArray[i - sendLen];
                }
                SendCmd(sendArray, "BeepOrSetting", 0);
                setRegistry("SN", sn_reg); //寫入序號到註冊機碼
                //寫入序號到畫面
                PrinterSNFacTxt.Text = sn_reg;
                PrinterSNTxt.Text = sn_reg;
            }
            else if (sn_reg.Length < 16 || sn_reg.Length > 16)
            {
                MessageBox.Show(FindResource("LessLength") as string);
            }
        }
        #endregion

        #region 輸入的機器序列號檢查
        private bool printerSNInputChk()
        {
            bool isOK = true;
            string sn = null;
            switch (SNTxtSettingPosition)
            {
                case "factory":
                    sn = PrinterSNFacTxt.Text;
                    break;
                case "communication":
                    sn = PrinterSNTxt.Text;
                    break;
            }
            if (sn == "")
            { //cloumn is empty
                isOK = false;
            }
            else
            {
                try
                {
                    string area = sn.Substring(0, 2);
                    DateTime date = DateTime.ParseExact(sn.Substring(2, 8), "yyyyMMdd", null); //轉換成date就自動判斷月日是否正確
                    int year = date.Year;
                    string codeString = sn.Substring(10, 6);
                    int codeInt = Int32.Parse(codeString); //如果pare錯誤就是非數字，會丟到catch
                    if ((area == "HW" || area == "GN"))
                    {
                        if (year >= 2020 && year <= 2030)
                        {
                            isOK = true;
                        }
                        else
                        {
                            isOK = false;
                        }
                    }
                    else
                    {
                        isOK = false;
                    }

                }
                catch (Exception) //日期格式或代碼格式錯誤
                {
                    isOK = false;
                }
            }
            return isOK;
        }
        #endregion
        //========================IAP 升級Firmware相關===================

        #region hid裝置連線狀態
        public void device_connect_status(bool con)
        {
            this.Dispatcher.Invoke(set_connect_status, con);
        }
        #endregion

        #region hid裝置連線狀態更新到ui
        public void device_connect_status_ui(bool con)
        {
            if (con)
            {
                DeviceStatusTxt.Text = FindResource("DeviceConnected") as string;
                openfileAndDownloadUIControl(true);
                ReconnectBtn.IsEnabled = false;
            }
            else
            {
                DeviceStatusTxt.Text = FindResource("DeviceDisconnected") as string;
                ReconnectBtn.IsEnabled = true;

                if (isBin)
                {
                     //isBinUpgradeSuccess = true; 
                        isLoadBinSuccess = true;
                }
                else {
                    //isUpgradeSuccess = true;
                    isLoadHexSuccess = true;
                }
                openfileAndDownloadUIControl(false);
            }
        }
        #endregion

        #region fw更新畫面初始化
        private void UIintial()
        {   
            StatusLabel.Content = "";
            CodeSizeLabel.Content = "";
            AddrLabel.Content = "";
            ReadFileProgress.Value = 0;
            WriteStatusLabel.Content = "";
            DownloadProgress.Value = 0;
            DownloadTimeTxt.Text = "";
            ConverTimeTxt.Text = "";
        }
        #endregion

        #region ui內容更新function
        private void updata_ui_status_text(byte index, string text)
        {
            switch (index)
            {
                case 1:
                    StatusLabel.Content = text;
                    break;

                case 2:
                    CodeSizeLabel.Content = text;
                    break;

                case 3:
                    AddrLabel.Content = text;
                    break;

                case 4:
                    ReadFileProgress.Value = int.Parse(text);
                    break;

                case 5:
                    WriteStatusLabel.Content = text;
                    break;

                case 6:
                    DownloadProgress.Value = int.Parse(text);
                    break;
                case 7:
                    MessageBox.Show(text);
                    break;
                case 8://使用show就不會影響後方執行
                    IAPDialog iapDlg = new IAPDialog();
                    iapDlg.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                    iapDlg.Owner = this;
                    iapDlg.Show();
                    break;
                case 9://傳輸資料與完畢判斷關閉開啟button
                    if (text.Equals("finish"))
                    {
                        openfileAndDownloadUIControl(true);
                        ReconnectBtn.IsEnabled = true;
                    }
                    else if (text.Equals("sending")) {          
                        openfileAndDownloadUIControl(false);
                        ReconnectBtn.IsEnabled = false;
                    }
                    break;
            }
        }
        #endregion

        #region 執行檔案解析
        private void updata_file_button_Click(object sender, EventArgs e)
        {
            download_time = 0;
            hex_to_bin_time = 0;
            //实例化回调
            setCallBack = new setTextValueCallBack(updata_ui_status_text);
            if (isBin || isLoadBinSuccess) 
            { 
                iap_download.get_bin_array(sender);
            }
            else
            {
                Thread thread = new Thread(new ParameterizedThreadStart(iap_download.hex_file_to_bin_array));
                thread.IsBackground = true;
                thread.Start(this);
            }
        }
        #endregion

        #region 解析與更新計時器
        private void timer_1s_Tick(object sender, EventArgs e)
        {
            switch (iap_download.run_step)
            {
                case 2:
                    download_time++;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        DownloadTimeTxt.Text = download_time.ToString();
                    }), null);
                    break;
                case 1:
                    hex_to_bin_time++;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ConverTimeTxt.Text = hex_to_bin_time.ToString();
                    }), null);

                    break;
                case 3: //下載秒數歸0
                    download_time=0;
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        DownloadTimeTxt.Text = "";
                    }), null);

                    break;
            }
        }
        #endregion

        #region 是否開啟打開文件與下載button
        public void openfileAndDownloadUIControl(bool isEnable) {
            DownloadFWBtn.IsEnabled = isEnable;
            OpenFWfileBtn.IsEnabled = isEnable;
           
        }
        #endregion
        //==============================打印機實時狀態============================

        #region 打印機即時狀態
        //因為機器故障時只能收到實時狀態命令，故移除確認通訊的命令傳輸
        public void PrinterNowStatus()
        {
            byte[] sendArray = StringToByteArray(Command.STATUS_MONITOR);

            switch (DeviceType)
            {
                case "RS232":
                    if (RS232PortName != null) //先判斷是否有get prot name
                    {
                        bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);
                        if (!isError)
                        {
                            RS232Connect.SerialPortSendCMD("NeedReceive", sendArray, null, 9);
                            while (!RS232Connect.isReceiveData)
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
                            SendCmdFail("R"); //避免傳輸時突然有問題
                        }
                        else //serial open port failed
                        {
                            stopStatusMonitorTimer();
                            stopSendCmdTimer();
                            connectFailUI(RS232ConnectImage, FindResource("CannotOpenComport") as string);
                        }
                    }
                    else
                    {
                        stopSendCmdTimer();
                        stopStatusMonitorTimer();
                        MessageBox.Show(FindResource("NotSettingComport") as string);
                    }
                    break;
                case "USB":
                    if (USBpath != null) //先判斷是否有USBpath
                    {
                        int result = USBConnect.ConnectUSBDevice(USBpath);
                        if (result == 1)
                        {
                            USBConnect.USBSendCMD("NeedReceive", sendArray, null, 9);
                            while (!USBConnect.isReceiveData)
                            {
                                if (USBConnect.mRecevieData != null)
                                {
                                    switch (QueryNowStatusPosition)
                                    {
                                        case "bottom":
                                            showPrinteNowStatus(USBConnect.mRecevieData, StatusMonitorLabel);
                                            break;
                                        case "maintain":
                                            showPrinteNowStatus(USBConnect.mRecevieData, PrinterStatusText);
                                            break;
                                    }
                                    break;
                                }
                            }
                            SendCmdFail("U");//避免傳輸時突然有問題
                        }
                        else
                        {
                            stopSendCmdTimer();
                            stopStatusMonitorTimer();
                            connectFailUI(USBConnectImage, FindResource("NotSettingUSBport") as string);
                        }
                    }
                    else
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        MessageBox.Show(FindResource("NotSettingUSBport") as string);
                    }
                    break;
                case "Ethernet":
                    bool isOK = chekckEthernetIPText();
                    if (isOK)
                    {
                        int connectStatus = EthernetConnect.EthernetConnectStatus();
                        switch (connectStatus)
                        {
                            case 0: //fail
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(EthernetConnectImage, FindResource("NotSettingEthernetport") as string);
                                break;
                            case 1: //success                           
                                bool isReceiveNowStatus = EthernetConnect.EthernetSendCmd("NeedReceive", sendArray, null, 9);
                                while (!EthernetConnect.isReceiveData)
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
                                        break;
                                    }
                                }
                                break;
                            case 2: //timeout
                                stopStatusMonitorTimer();
                                stopSendCmdTimer();
                                connectFailUI(EthernetConnectImage, FindResource("ConnectTimeout") as string);
                                break;
                        }
                    }
                    break;
            }
        }
        #endregion

        //===============================定時登出timer功能========================

        #region 啟動閒置檢查timer
        private void IdleAndLogoutTimerStart()
        {
            if (idleTimer.IsEnabled)
            {
                idleTimer.Stop();
            }
            idleTimer.Interval = TimeSpan.FromSeconds(10); //每10秒檢查一次閒置狀態

            // 加入callback function
            idleTimer.Tick += idle_timer_tick;

            idleTimer.Start();
        }
        #endregion

        #region 閒置時間計算並登出
        private void idle_timer_tick(object sender, EventArgs e)
        {
            //TimeSpan? 代表TimeSpan可為null
            //app閒置時間,可判斷app非在前景的時間
            TimeSpan? appIdle = lostFocusTime == null ? null : (TimeSpan?)DateTime.Now.Subtract((DateTime)lostFocusTime);

            //系統閒置時間,可判斷開啟視窗完全沒有移動滑鼠或使用鍵盤的狀態
            TimeSpan machineIdle = IdleCheck.GetLastInputTime();

            //設定閒置多少時間要登出
            TimeSpan idleTimeSpan = new TimeSpan(0, 10, 0); //設定閒置時間為10分鐘

            //app在前景且滑鼠或鍵盤無使用超過設定之閒置時間
            bool isMachineIdle = machineIdle > idleTimeSpan;

            //app在背景超過設定之閒置時間
            bool isAppIdle = appIdle != null && appIdle > idleTimeSpan;

            if (isAppIdle || isMachineIdle)
            {
                if (idleTimer.IsEnabled)
                {
                    idleTimer.Stop();
                }
                isLoginAdmin = false;
                MessageBox.Show(FindResource("Logout") as string);
                isParaSettingBtnEnabled(false);

            }

        }
        #endregion

        //===============================實時狀態timer功能========================

        #region 啟動實時狀態查詢
        private void startStatusMonitorTimer()
        {
            statusMonitorTimer = new Timer();
            statusMonitorTimer.Interval = 1000;
            statusMonitorTimer.Elapsed += timer_Elapsed;
            statusMonitorTimer.Start();
            QueryNowStatusPosition = "bottom";
            StatusMonitorBtn.Content = FindResource("StopStatusMonitor") as string;
        }
        #endregion

        #region timer的ui處理
        private void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                PrinterNowStatus();
                //DifferInterfaceConnectChkAndSend("PrinterNowStatus");
            }), null);
        }
        #endregion

        #region 關閉實時狀態查詢
        private void stopStatusMonitorTimer()
        {
            if (statusMonitorTimer != null)
            {
                statusMonitorTimer.Dispose();
                StatusMonitorBtn.Content = FindResource("StartStatusMonitor") as string;
            }
        }
        #endregion

        //=========================定時發送命令timer功能===========================

        #region 啟動定時發送命令
        private void startSendCmdTimer(int intervel)
        {
            sendCmdTimer = new Timer();
            sendCmdTimer.Interval = intervel;
            sendCmdTimer.Elapsed += timer_Send;
            sendCmdTimer.Start();
            SendCmdItervalBtn.Content = FindResource("StopSendCmd") as string;
        }
        #endregion

        #region timer的ui處理
        private void timer_Send(object sender, ElapsedEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SendCmdWithInterval();
                //DifferInterfaceConnectChkAndSend("SendCmd");
            }), null);
        }
        #endregion

        #region 關閉定時發送命令
        private void stopSendCmdTimer()
        {
            if (sendCmdTimer != null)
            {
                sendCmdTimer.Dispose();
                SendCmdItervalBtn.Content = FindResource("StartSendCmd") as string;
                switch (DeviceType)
                {
                    case "RS232":
                        RS232Connect.CloseSerialPort(); //最後關閉
                        break;
                    case "USB":
                        USBConnect.closeHandle();
                        break;
                    case "Ethernet":
                        EthernetConnect.disconnect();
                        break;
                }
            }
        }
        #endregion

        //==========================註冊機碼的寫入與讀取=============================

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

        #region 註冊機碼的寫入
        private void setRegistry(string valueName, string value)
        {
            RegistryKey registryKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\ZLPPT"); //修改也要使用create不能用open
            registryKey.SetValue(valueName, value);
        }
        #endregion

        #region 註冊機碼的讀取
        private string getRegistry(string valueName)
        {
            string lastValue = null;
            if (Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ZLPPT").GetValue(valueName) == null)
            {
                lastValue = "";
            }
            else
            {
                lastValue = (string)Registry.LocalMachine.OpenSubKey(@"SOFTWARE\ZLPPT").GetValue(valueName);
            }
            return (string)lastValue;
        }
        #endregion

        //============================所有功能整合===================================

        #region 整合所有功能
        public void AllFunctionCollection(string cmdType)
        {
            switch (cmdType)
            {
                case "Restart":
                    byte[] sendArray = StringToByteArray(Command.RESTART);
                    SendCmd(sendArray, "BeepOrSetting", 0);
                    USBConnect.IsConnect = false;
                    RS232Connect.IsConnect = false;
                    EthernetConnect.connectStatus = 0;
                    EthernetConnectImage.Source = new BitmapImage(new Uri("Images/grey_circle.png", UriKind.Relative));
                    USBConnectImage.Source = new BitmapImage(new Uri("Images/grey_circle.png", UriKind.Relative));
                    RS232ConnectImage.Source = new BitmapImage(new Uri("Images/grey_circle.png", UriKind.Relative));
                    //if (isRS232Connected != false && isUSBConnected != false && isEthernetConnected != false) {
                    //    MessageBox.Show(FindResource("Reconnect") as string); //通訊失敗就不用跳訊息
                    //}
                    break;
                case "CleanPrinterInfoOneByOne":
                    IsPrinterInfoChecked();//先確認選取狀態
                    if (Config.isFeedLinesChecked)
                    {
                        byte[] sendArrayFeed = StringToByteArray(Command.CLEAN_PRINTINFO_FEED_LINES);
                        SendCmd(sendArrayFeed, "BeepOrSetting", 0);
                    }

                    if (Config.isPrintedLinesChecked)
                    {
                        byte[] sendArrayPrint = StringToByteArray(Command.CLEAN_PRINTINFO_PRINTED_LINES);
                        SendCmd(sendArrayPrint, "BeepOrSetting", 0);
                    }

                    if (Config.isCutPaperTimesChecked)
                    {
                        byte[] sendArrayCut = StringToByteArray(Command.CLEAN_PRINTINFO_CUTPAPER_TIMES);
                        SendCmd(sendArrayCut, "BeepOrSetting", 0);
                    }

                    if (Config.isHeadOpenTimesChecked)
                    {
                        byte[] sendArrayHead = StringToByteArray(Command.CLEAN_PRINTINFO_HEADOPEN_TIMES);
                        SendCmd(sendArrayHead, "BeepOrSetting", 0);
                    }

                    if (Config.isPaperOutTimesChecked)
                    {
                        byte[] sendArrayOut = StringToByteArray(Command.CLEAN_PRINTINFO_PAPEROUT_TIMES);
                        SendCmd(sendArrayOut, "BeepOrSetting", 0);
                    }
                    if (Config.iErrorTimesChecked)
                    {
                        byte[] sendArrayError = StringToByteArray(Command.CLEAN_PRINTINFO_ERROR_TIMES);
                        SendCmd(sendArrayError, "BeepOrSetting", 0);
                    }
                    break;
                case "SendEmpty":
                    byte[] emptyArray = { 0x0a };
                    SendCmd(emptyArray, "BeepOrSetting", 0);
                    break;
                //case "Send3Empty": //通訊測試後連送三個0x0a
                //    byte[] empty3Array = { 0x0a, 0x0a, 0x0a };
                //    SendCmd(empty3Array, "BeepOrSetting", 0);
                //    break;
                case "SendCmd":
                    SendCmd();
                    break;
                case "SetIP":
                    SetIP();
                    break;
                case "SetGateway":
                    SetGateway();
                    break;
                case "SetMAC":
                    SetMAC();
                    break;
                case "AutoDisconnect":
                    AutoDisconnect();
                    break;
                case "ConnectClient":
                    ConnectClient();
                    break;
                case "EthernetSpeed":
                    EthernetSpeed();
                    break;
                case "DHCPMode":
                    DHCPMode();
                    break;
                case "USBMode":
                    USBMode();
                    break;
                case "USBFixed":
                    USBFixed();
                    break;
                case "CodePageSet":
                    CodePageSet();
                    break;
                case "CodePagePrint":
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
                        byte[] sendArrayCode = codePage.ToArray();
                        SendCmd(sendArrayCode, "BeepOrSetting", 0);
                    }
                    else { MessageBox.Show(FindResource("ColumnEmpty") as string); }
                    break;
                case "LanguageSet":
                    LanguageSet();
                    break;
                case "FontBSetting":
                    FontBSetting();
                    break;
                case "CustomziedFont":
                    CustomziedFont();
                    break;
                case "SetDirection":
                    SetDirection();
                    break;
                case "MotorAccControl":
                    MotorAccControl();
                    break;
                case "AccMotor":
                    AccMotor();
                    break;
                case "PrintSpeed":
                    PrintSpeed();
                    break;
                //case "DensityMode":
                //    DensityMode();
                //    break;
                case "Density":
                    Density();
                    break;
                case "PaperOutReprint":
                    PaperOutReprint();
                    break;
                case "PaperWidth":
                    PaperWidth();
                    break;
                //case "HeadCloseCut":
                //    HeadCloseCut();
                //    break;
                case "YOffset":
                    YOffset();
                    break;
                case "MACShow":
                    MACShow();
                    break;
                case "QRCode":
                    QRCode();
                    break;
                case "LogoPrintControl":
                    LogoPrintControl();
                    break;
                case "DIPSwitch":
                    DIPSwitch();
                    break;
                case "DIPSetting":
                    DIPSetting();
                    break;
                case "readALL":
                    readALL();
                    break;
                case "sendALL":
                    sendALL();
                    break;
                case "PrinterInfoRead":
                    PrinterInfoRead();
                    break;
                case "cleanPrinterInfo":
                    cleanPrinterInfo();
                    break;
                case "queryPrinterStatus":
                    queryPrinterStatus();
                    break;
                case "PrintTestShort":
                    byte[] sendArrayShort = StringToByteArray(Command.PRINT_TEST_SHORT);
                    SendCmd(sendArrayShort, "BeepOrSetting", 0);
                    break;
                case "PrintTestLong":
                    byte[] sendArrayLong = StringToByteArray(Command.PRINT_TEST_LONG);
                    SendCmd(sendArrayLong, "BeepOrSetting", 0);
                    break;
                case "PrintEvenTest":
                    PrintEvenTest();
                    break;
                case "BeepTest":
                    BeepTest();
                    break;
                case "OpenCashBox":
                    OpenCashBox();
                    break;
                case "CutTimesFactory":
                    int resultF;
                    string timesF = CutTimesTxt.Text;
                    if (timesF == "" || !Int32.TryParse(timesF, out resultF))
                    {
                        MessageBox.Show(FindResource("ErrorFormat") as string);
                    }
                    else
                    {
                        int contiunue = Int32.Parse(timesF);
                        byte[] sendArrayCutF = StringToByteArray(Command.CUT_TIMES);
                        for (int i = 0; i < contiunue; i++)
                        {
                            SendCmd(sendArrayCutF, "BeepOrSetting", 0);
                            if (DeviceType == "Ethernet")
                            {
                                Thread.Sleep(100);
                            }
                        }
                    }
                    break;
                case "CutTimesMaintain":
                    int result;
                    string times = null;
                    times = CutTimes_Maintanin_Txt.Text;
                    if (times == "" || !Int32.TryParse(times, out result))
                    {
                        MessageBox.Show(FindResource("ErrorFormat") as string);
                    }
                    else
                    {
                        int contiunue = Int32.Parse(times);
                        byte[] sendArrayCutM = StringToByteArray(Command.CUT_TIMES);
                        for (int i = 0; i < contiunue; i++)
                        {
                            SendCmd(sendArrayCutM, "BeepOrSetting", 0);
                            if (DeviceType == "Ethernet")
                            {
                                Thread.Sleep(100);
                            }
                        }
                    }
                    break;
                case "SDRAMTest":
                    byte[] SDRAMArray = StringToByteArray(Command.SDRAM_TEST);
                    SendCmd(SDRAMArray, "BeepOrSetting", 0);
                    break;
                case "FlashTest":
                    byte[] FlashArray = StringToByteArray(Command.FLASH_TEST);
                    SendCmd(FlashArray, "BeepOrSetting", 0);
                    break;
                case "CMDTestFactory":
                    CMDTestFactory();
                    break;
                case "CMDTestMaintain":
                    CMDTestMaintain();
                    break;
                case "RoadPrinterSN":
                    RoadPrinterSN();
                    break;
                case "SetPrinterSN":
                    SetPrinterSN();
                    break;
                //case "PrinterNowStatus":
                //    PrinterNowStatus();
                //    break;
                case "PrintLogo":
                    nvLogoRadioBtnChecked();
                    int number;
                    if (NVLogoPieceTXT.Text != null && Int32.TryParse(NVLogoPieceTXT.Text, out number))
                    {
                        string numberHex = number.ToString("X2");
                        byte[] sendArrayPrint = StringToByteArray(Command.PRINT_LOGOS_HEADER + numberHex + nvLogo_m_hex);
                        SendCmd(sendArrayPrint, "BeepOrSetting", 0);
                    }
                    else
                    {
                        MessageBox.Show(FindResource("PrintPieceEmpty") as string);
                    }
                    break;
                case "ClearLogo":
                    byte[] sendArrayClear = StringToByteArray(Command.CLEAN_LOGOS_INPRINTER);
                    SendCmd(sendArrayClear, "BeepOrSetting", 0);
                    break;
                case "DonwaldLogo":
                    if (nvLogo_full_hex.Length == 0)
                    {
                        MessageBox.Show(FindResource("GalleryEmpty") as string);
                    }
                    else
                    {
                        nvLogo_n_hex = fileNameArray.Length.ToString("X2");
                        byte[] insertBtye = StringToByteArray(nvLogo_n_hex);
                        byte[] sendArrayDownload = StringToByteArray(nvLogo_full_hex.ToString());
                        List<byte> sendList = sendArrayDownload.ToList();
                        sendList.Insert(2, insertBtye[0]);
                        sendArrayDownload = sendList.ToArray();
                        SendCmd(sendArrayDownload, "BeepOrSetting", 0);
                        Console.WriteLine(BitConverter.ToString(sendArrayDownload).Replace("-", ""));
                        MessageBox.Show(FindResource("WaitforRedLight") as string);
                    }
                    break;
                case "ReadBT":
                    LoadBTName();
                    break;
                case "ReadWIFIName":
                    LoadWIFIName();
                    break;
                case "WriteWIFIName":
                    SetWIFIName();
                    break;
                case "ReadWIFIPwd":
                    LoadWIFIPwd();
                    break;
                case "WriteWIFIPwd":
                    SetWIFIPWD();
                    break;
                case "CutBeepSettings":
                    CutBeepSettings();
                    break;
            }
        }
        #endregion

        //======================通訊介面連線狀態檢查並執行命令傳送=======================

        #region 檢查通訊介面並傳送命令
        public void DifferInterfaceConnectChkAndSend(string cmdType)
        {
            switch (DeviceType)
            {
                case "RS232":
                    if (RS232PortName != null) //先判斷是否有get prot name
                    {    //已斷線，重新測試連線
                        if (!RS232Connect.RS232ConnectStatus())
                        {
                            checkRS232Communitcation();
                        }
                        if (isRS232Connected)
                        {
                            AllFunctionCollection(cmdType);
                            RS232Connect.CloseSerialPort(); //最後關閉
                        }

                    }
                    else
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        MessageBox.Show(FindResource("NotSettingComport") as string);
                    }
                    break;
                case "USB":
                    if (USBpath != null) //先判斷是否有USBpath
                    {    //已斷線，重新測試連線
                        if (USBConnect.USBHandle == -1)
                        {   //usb在打印logo時如果確認通訊會來不及反應，這邊bypass掉
                            //主要是因為部分命令必須執行完成後才能回應後續命令，如切刀命令和打印nvlogo命令有這個問題
                            if (cmdType.Equals("PrintLogo") || cmdType.Equals("CodePagePrint") || cmdType.Equals("CutTimesFactory") || cmdType.Equals("CutTimesMaintain"))
                            {
                                USBConnect.ConnectUSBDevice(USBpath);
                                isUSBConnected = true;
                            }
                            else
                            {
                                checkUSBCommunitcation();
                            }
                        }
                        if (isUSBConnected)
                        {
                            AllFunctionCollection(cmdType);
                            USBConnect.closeHandle(); //最後關閉
                        }
                    }
                    else
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        MessageBox.Show(FindResource("NotSettingUSBport") as string);
                    }
                    break;
                case "Ethernet":
                    bool isOK = chekckEthernetIPText();
                    if (isOK)
                    {
                        int connectStatus = EthernetConnect.EthernetConnectStatus();
                        if (connectStatus != 1)
                        {
                            checkEthernetCommunitcation(connectStatus);
                        }
                        if (isEthernetConnected)
                        {
                            AllFunctionCollection(cmdType);
                            EthernetConnect.disconnect(); //最後關閉
                        }
                    }
                    break;
            }
        }
        #endregion

        #region RS232通訊測試
        public void checkRS232Communitcation()
        {
            bool isError = RS232Connect.OpenSerialPort(RS232PortName, FindResource("CannotOpenComport") as string);
            if (!isError)
            {

                byte[] sendArray = StringToByteArray(TEST_SEND_CMD);
                RS232Connect.SerialPortSendCMD("NeedReceive", sendArray, null, 9);
                while (!RS232Connect.isReceiveData)
                {
                    if (RS232Connect.mRecevieData != null)
                    {
                        isRS232CommunicateOK(RS232Connect.mRecevieData, "cmdSend");
                        break;
                    }
                }
                SendCmdFail("R");
            }
            else //serial open port failed
            {
                stopStatusMonitorTimer();
                stopSendCmdTimer();
                isRS232Connected = false;
                connectFailUI(RS232ConnectImage, FindResource("CannotOpenComport") as string);
            }

        }
        #endregion

        #region USB通訊測試
        //有開啟TIMER都要關閉，否則會一直跳出錯誤訊息導致卡住UI畫面
        public void checkUSBCommunitcation()
        {
            int result = USBConnect.ConnectUSBDevice(USBpath);
            if (result == 1)
            {

                byte[] sendArray = StringToByteArray(TEST_SEND_CMD);
                USBConnect.USBSendCMD("NeedReceive", sendArray, null, 9);
                while (!USBConnect.isReceiveData)
                {
                    if (USBConnect.mRecevieData != null)
                    {
                        isUSBCommunicateOK(USBConnect.mRecevieData, "cmdSend");
                        break;
                    }
                }
                SendCmdFail("U");
            }
            else //USB CreateFile失敗
            {
                stopSendCmdTimer();
                stopStatusMonitorTimer();
                isUSBConnected = false;
                connectFailUI(USBConnectImage, FindResource("NotSettingUSBport") as string);
            }

        }
        #endregion

        #region Ethernet通訊測試
        public void checkEthernetCommunitcation(int connectStatus)
        {
            switch (connectStatus)
            {
                case 0: //fail
                    stopStatusMonitorTimer();
                    stopSendCmdTimer();
                    isEthernetConnected = false;
                    connectFailUI(EthernetConnectImage, FindResource("NotSettingEthernetport") as string);
                    break;
                case 1: //success
                    byte[] sendArray = StringToByteArray(TEST_SEND_CMD);
                    EthernetConnect.EthernetSendCmd("NeedReceive", sendArray, null, 9);
                    while (!EthernetConnect.isReceiveData)
                    {
                        if (EthernetConnect.mRecevieData != null)
                        {
                            isEthernetCommunicateOK(EthernetConnect.mRecevieData, "cmdSend");
                            break;
                        }
                    }
                    SendCmdFail("E");
                    break;
                case 2: //timeout
                    stopStatusMonitorTimer();
                    stopSendCmdTimer();
                    isEthernetConnected = false;
                    connectFailUI(EthernetConnectImage, FindResource("ConnectTimeout") as string);
                    break;

            }
        }
        #endregion

        #region RS232判斷是否正確收到通訊測試回傳訊息
        public void isRS232CommunicateOK(byte[] data, string testSource)
        {
            string receiveData = BitConverter.ToString(data);
            Console.WriteLine("RS232:" + receiveData);
            if (testSource == "cmdSend")
            {
                if (receiveData.Contains(TEST_RECEIVE_CMD))
                {
                    isRS232Connected = true;
                    connectSuccessUI(RS232ConnectImage);
                }
                else //回傳資料錯誤時
                {
                    isRS232Connected = false;
                    connectFailUI(RS232ConnectImage, FindResource("FailToSend") as string);
                }

            }
            else if (testSource == "communication")
            {
                if (receiveData.Contains(Command.RS232_COMMUNICATION_RECEIVE))
                {
                    isRS232Connected = true;
                    connectSuccessUI(RS232ConnectImage);
                }
                else //回傳資料錯誤時
                {
                    isRS232Connected = false;
                    connectFailUI(RS232ConnectImage, FindResource("FailToSend") as string);
                }

            }

        }
        #endregion

        #region USB判斷是否正確收到通訊測試回傳訊息
        public void isUSBCommunicateOK(byte[] data, string testSource)
        {
            string receiveData = BitConverter.ToString(data);
            Console.WriteLine("USB:" + receiveData);
            if (testSource == "cmdSend")
            {
                if (receiveData.Contains(TEST_RECEIVE_CMD))
                {
                    isUSBConnected = true;
                    connectSuccessUI(USBConnectImage);
                }
                else //回傳資料錯誤時
                {
                    isUSBConnected = false;
                    connectFailUI(USBConnectImage, FindResource("FailToSend") as string);
                }

            }
            else if (testSource == "communication")
            {
                if (receiveData.Contains(Command.USB_COMMUNICATION_RECEIVE))
                {
                    isUSBConnected = true;
                    connectSuccessUI(USBConnectImage);
                }
                else //回傳資料錯誤時
                {
                    isUSBConnected = false;
                    connectFailUI(USBConnectImage, FindResource("FailToSend") as string);
                }

            }
        }
        #endregion

        #region Ethernet判斷是否正確收到通訊測試回傳訊息
        public void isEthernetCommunicateOK(byte[] data, string testSource)
        {
            string receiveData = BitConverter.ToString(data);
            Console.WriteLine("Ethernet:" + receiveData);
            if (testSource == "cmdSend")
            {
                if (receiveData.Contains(TEST_RECEIVE_CMD))
                {
                    isEthernetConnected = true;
                    connectSuccessUI(EthernetConnectImage);
                }
                else //回傳資料錯誤時
                {
                    isEthernetConnected = false;
                    connectFailUI(EthernetConnectImage, FindResource("FailToSend") as string);
                }
            }
            else if (testSource == "communication")
            {
                if (receiveData.Contains(Command.ETHERNET_COMMUNICATION_RECEIVE))
                {
                    isEthernetConnected = true;
                    connectSuccessUI(EthernetConnectImage);
                }
                else //回傳資料錯誤時
                {
                    isEthernetConnected = false;
                    connectFailUI(EthernetConnectImage, FindResource("FailToSend") as string);
                }

            }
        }
        #endregion

        #region 連線成功的變數設定與UI處理
        public void connectSuccessUI(System.Windows.Controls.Image circleImage)
        {
            circleImage.Source = new BitmapImage(new Uri("Images/green_circle.png", UriKind.Relative));
            // isTabEnabled(true);
        }
        #endregion

        #region 連線失敗的變數設定與UI處理
        public void connectFailUI(System.Windows.Controls.Image circleImage, string errorMessage)
        {
            circleImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative));
            MessageBox.Show(errorMessage);
            setSysStatusColorAndText(errorMessage, "#FFEF7171");
            // isTabEnabled(false);

        }
        #endregion

        #region 命令傳送失敗的處理
        public void SendCmdFail(string connectType)
        {
            switch (connectType)
            {
                case "R": //RS232
                    if (!RS232Connect.IsConnect)
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        isRS232Connected = false;
                        connectFailUI(RS232ConnectImage, FindResource("ConnectTimeout") as string);
                    }
                    break;
                case "U": //USB
                    if (!USBConnect.IsConnect)
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        isUSBConnected = false;
                        connectFailUI(USBConnectImage, FindResource("ConnectTimeout") as string);
                        USBConnect.closeHandle();
                    }
                    break;
                case "E":
                    if (EthernetConnect.connectStatus != 1)
                    {
                        stopStatusMonitorTimer();
                        stopSendCmdTimer();
                        isEthernetConnected = false;
                        connectFailUI(EthernetConnectImage, FindResource("ConnectTimeout") as string);
                    }
                    break;
            }
        }
        #endregion

        //===========================命令傳送與接收資料==============================
        //連線失敗時isReceiveData為true
        #region RS232傳送與接收資料
        private void SerialPortConnect(string dataType, byte[] data, int receiveLength)
        {
            //避免傳輸中斷時重複判斷連線狀態造成訊息一直跳出
            if (isRS232Connected)
            {
                switch (dataType)
                {
                    case "ReadPara":
                        RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                setParaColumn(RS232Connect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadSN":
                        RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                SetPrinterInfo(RS232Connect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadPrinterInfo": //打印機統計信息
                        RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        while (!RS232Connect.isReceiveData)
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
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                setPrinterStatus(RS232Connect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadBTName": //藍牙名稱
                        RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                setBTNametoUI(RS232Connect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadWIFIName": //wifi名稱
                        RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                setWIFINametoUI(RS232Connect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadWIFIPwd": //wifi密碼
                        RS232Connect.SerialPortSendCMD("NeedReceive", data, null, receiveLength);
                        while (!RS232Connect.isReceiveData)
                        {
                            if (RS232Connect.mRecevieData != null)
                            {
                                setWIFIPwdtoUI(RS232Connect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "BeepOrSetting":
                        RS232Connect.SerialPortSendCMD("NoReceive", data, null, 0);
                        break;
                }
            }
        }
        #endregion

        #region USB傳送命令與接收資料
        private void USBConnectAndSendCmd(string dataType, byte[] data, int receiveLength)
        {
            if (isUSBConnected)
            {
                switch (dataType)
                {
                    case "ReadPara":
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                setParaColumn(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadSN":
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                SetPrinterInfo(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadPrinterInfo": //打印機統計信息
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                setPrinterInfotoUI(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadStatus": //打印機溫度電壓等狀態
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                setPrinterStatus(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadBTName": //藍牙名稱
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                setBTNametoUI(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadWIFIName": //wifi名稱
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                setWIFINametoUI(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadWIFIPwd": //wifi密碼
                        USBConnect.USBSendCMD("NeedReceive", data, null, receiveLength);
                        while (!USBConnect.isReceiveData)
                        {
                            if (USBConnect.mRecevieData != null)
                            {
                                setWIFIPwdtoUI(USBConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "BeepOrSetting":
                        USBConnect.USBSendCMD("NoReceive", data, null, 0);
                        break;
                }
            }
        }
        #endregion

        #region 網口傳送命令與接收資料 
        private void EthernetConnectAndSendCmd(string dataType, byte[] data, int receiveLength)
        {
            if (isEthernetConnected)
            {
                switch (dataType)
                {
                    case "ReadPara":
                        bool isReceiveData = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setParaColumn(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadSN":
                        bool isReceiveSN = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                SetPrinterInfo(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadPrinterInfo": //打印機統計信息
                        bool isReceivePI = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setPrinterInfotoUI(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadStatus": //打印機溫度電壓等狀態
                        bool isReceiveStatus = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setPrinterStatus(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadBTName": //藍牙名稱
                        bool isReceivebt = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setBTNametoUI(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadWIFIName": //wifi名稱
                        bool isReceiveWifiName = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setWIFINametoUI(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "ReadWIFIPwd": //wifi密碼
                        bool isReceiveWifiPwd = EthernetConnect.EthernetSendCmd("NeedReceive", data, null, receiveLength);
                        while (!EthernetConnect.isReceiveData)
                        {
                            if (EthernetConnect.mRecevieData != null)
                            {
                                setWIFIPwdtoUI(EthernetConnect.mRecevieData);
                                break;
                            }
                        }
                        break;
                    case "BeepOrSetting":
                        EthernetConnect.EthernetSendCmd("NoReceive", data, null, 0);
                        break;
                }
            }
        }
        #endregion

        #region 不同通道傳送命令
        public void SendCmd(byte[] sendArray, string sendType, int length)
        {
            //stopTimer();
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

        //===========================Ethernet 相關檢查=============================

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
                else //有定時傳送時，若有錯誤會重複跳出訊息，要先關閉定時傳送
                {
                    stopStatusMonitorTimer();
                    stopSendCmdTimer();
                    EthernetConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //欄位錯誤連線也會錯誤
                    EthernetIPAddress = null;
                    MessageBox.Show(FindResource("ErrorFormat") as string);
                }
            }
            else
            {
                stopStatusMonitorTimer();
                stopSendCmdTimer();
                EthernetConnectImage.Source = new BitmapImage(new Uri("Images/red_circle.png", UriKind.Relative)); //欄位空白連線也會錯誤
                MessageBox.Show(FindResource("IPEmpty") as string);
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
        public byte[] StringToByteArray(string hex)
        {
            string afterConvert = hex.Replace(" ", "");
            byte[] data = null;
            try
            {
                data = Enumerable.Range(0, afterConvert.Length)
                                  .Where(x => x % 2 == 0)
                                  .Select(x => Convert.ToByte(afterConvert.Substring(x, 2), 16))
                                  .ToArray();
            }
            catch (Exception)
            {
                MessageBox.Show(FindResource("HexStringError") as string);
            }
            return data;
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

        #region string to hex string
        public string ConvertStringToHex(String input, Encoding encoding)
        {
            Byte[] stringBytes = encoding.GetBytes(input);
            StringBuilder sbBytes = new StringBuilder(stringBytes.Length * 2);
            foreach (byte b in stringBytes)
            {
                sbBytes.AppendFormat("{0:X2}", b);
            }
            return sbBytes.ToString();
        }
        #endregion

        #region string to hex string without endoding
        public string ConvertStringToHex(String input)
        {
            Byte[] stringBytes = Encoding.Default.GetBytes(input);
            int stringLength = stringBytes.Length;
            if (stringLength > 16)
            {

                return null;
            }
            else {
                StringBuilder sbBytes = new StringBuilder(16);
                for (int i = 0; i < 16; i++) {
                    
                    if (i > stringLength - 1)
                    {
                        sbBytes.AppendFormat("{0:X2}", 0);
                    }
                    else {
                        sbBytes.AppendFormat("{0:X2}", stringBytes.ElementAt(i));
                    }
                }
                Console.WriteLine(sbBytes.Length);
                return sbBytes.ToString();
            }
           
        }
        #endregion

        #region 編碼轉換
        public Encoding convertEncoding()
        {   //編碼  
            Encoding result;
            switch (LanguageSetCom.SelectedIndex)
            {
                case 0:
                    result = Encoding.GetEncoding("gb18030");
                    break;
                case 1:
                    result = Encoding.GetEncoding("big5");
                    break;
                default: //預設沒有讀取打印機參數時，編碼為gb18030
                    result = Encoding.GetEncoding("gb18030");
                    break;

            }
            return result;
        }
        #endregion

        //=========================RS232 取得port並設定到UI=====================

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

        #region 取得RS232 Port後設定到UI
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
                lastRS232SelectedIndex = DeviceSelectRS232.SelectedIndex; //紀錄上次使用選項
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
                    if (seletedItem != null && usbTypePath.Contains(seletedItem.USBVIDPID) && usbTypePath.Contains(seletedItem.USBSN))
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
                viewmodel.removePort(deviceList); //msg == USBDetector.UsbDevicechange (int)wparam == USBDetector.NewUsbDeviceConnected || (int)wparam == USBDetector.UsbDeviceRemoved
                deviceList.Clear(); //清空避免重複
                getSerialPort();
                getUSBInfoandUpdateView();
                viewmodel.getDeviceObserve("rs232");
                if (lastRS232SelectedIndex != -1)
                {
                    DeviceSelectRS232.SelectedIndex = lastRS232SelectedIndex;
                }
                else if (lastRS232SelectedIndex == -1)
                { //預設
                    DeviceSelectRS232.SelectedIndex = viewmodel.RS232Device.Count - 1;
                }
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
            if (lastRS232SelectedIndex != -1)
            {
                DeviceSelectRS232.SelectedIndex = lastRS232SelectedIndex;
            }
            else if (lastRS232SelectedIndex == -1)
            { //預設
                DeviceSelectRS232.SelectedIndex = viewmodel.RS232Device.Count - 1;
            }
            viewmodel.getDeviceObserve("usb");
            DeviceSelectUSB.SelectedIndex = viewmodel.USBDevice.Count - 1;//設定選取第一筆      

            //增加判斷如果切換通道時未連線進行提醒
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
                //case 0:
                //    LoadLanguage("zh-TW");
                //    break;
                case 0:
                    LoadLanguage("zh-CN");
                    nowLanguage = "zh-CN";
                    break;
                    //case 2:
                    //    LoadLanguage("en-US");
                    //    nowLanguage = "en-US";
                    //    break;
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
            switch (lan)
            {
                //case "zh-TW":
                //    language.SelectedIndex = 0;
                //    break;
                case "zh-CN":
                    language.SelectedIndex = 0;
                    break;
                    //case "en-US":
                    //    language.SelectedIndex = 1;
                    //    break;
            }

        }
        #endregion


    }
}
