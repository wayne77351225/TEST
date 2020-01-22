using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace G80Utility.Tool
{
    class Command
    {
        //RS232接口測試
        public static string RS232_COMMUNICATION_TEST = "1F 1B 10 01 00";

        //USB接口測試
        public static string USB_COMMUNICATION_TEST = "1F 1B 10 02 00";

        //Ethernet接口測試
        public static string ETHERNET_COMMUNICATION_TEST = "1F 1B 10 03 00";

        //即時狀態
        public static string STATUS_MONITOR = "1F 1B 10 04 00";

        //軟體日期
        public static string SW_DATE = "1F 1B 10 05 00";

        //設定序號命令開頭
        public static string SN_SETTING_HEADER = "1F 1B 1F 53 5A 4A 42 5A 46 12 01 02 ";

        //讀取機器訊息
        public static string DEVICE_INFO_READING = "1F 1B 1F 53 5A 4A 42 5A 46 12 01 01";

        //重啟打印機
        public static string RESTART = "1F 1B 1F 53 5A 4A 42 5A 46 11 00 00";

        //設定ip address HEADER
        public static string IP_SETTING_HEADER = "1F 1B 1F 91 00 49 50";

        //設定gateway
        public static string GATEWAY_SETTING_HEADER = "1F 1B 1F 53 5A 4A 42 5A 46 30 11 02";

        //設定mac address Header
        public static string MAC_ADDRESS_SETTING_HEADER = "1F 1B 1F 91 00 49 44";

        //設定自動斷線時間HEADER
        public static string NETWORK_AUTODICONNECTED_SETTING_HEADER = "1F 1B 1F 95 01 02 03 11 12 13";

        //設定網路連接數量
        public static string CONNECT_CLIENT_1_SETTING = "1F 1B 1F 99 01 02 03 11 12 13 55";
        public static string CONNECT_CLIENT_2_SETTING = "1F 1B 1F 99 01 02 03 11 12 13 33";

        //設定網口通訊速度
        public static string ETHERNET_SPEED_SETTING_10MHZ = "1F 1B 1F 9D 11 12 13 01 02 03 33";
        public static string ETHERNET_SPEED_SETTING_100MHZ = "1F 1B 1F 9D 11 12 13 01 02 03 55";

        //設定DHCP模式HEADER
        public static string DHCP_MODE_SETTING_HEADER = "1F 1B 1F 10 13 14 15 19 18 17";


        //設定USB模式
        public static string USB_UTP_SETTING = "1F 1B 1F A3 16 17 18 19 18 17 76";
        public static string USB_VCOM_SETTING = "1F 1B 1F A3 16 17 18 19 18 17 67";

        //設定USB端口值
        public static string USB_FIXED_SETTING = "1F 1B 1F 9C 11 12 13 01 02 03 33";
        public static string USB_UNFIXED_SETTING = "1F 1B 1F 9C 11 12 13 01 02 03 55";


        //設定語言
        public static string LANGUAGE_SETTING_GB18030 = "1F 1B 1F EE 11 12 13 55 0A";
        public static string LANGUAGE_SETTING_BIG5 = "1F 1B 1F EE 11 12 13 33 0A";
        public static string LANGUAGE_SETTING_KOREAN = "1F 1B 1F EE 11 12 13 66 0A ";
        public static string LANGUAGE_SETTING_JAPANESE = "1F 1B 1F EE 11 12 13 44 0A";
        public static string LANGUAGE_SETTING_SHIFT_JIS = "1F 1B 1F 9E 11 12 13 01 02 03 55";
        public static string LANGUAGE_SETTING_JIS = "1F 1B 1F 9E 11 12 13 01 02 03 33";

        //設定FontB字體
        public static string FONTB_ON_SETTING = "1F 1B 1F E5 09 06 03 AA";
        public static string FONTB_OFF_SETTING = "1F 1B 1F E5 09 06 03 77";

        //設定走紙方向
        public static string DIRECTION_80250N_SETTING = "1F 1B 1F 92 10 11 12 15 16 17 33";
        public static string DIRECTION_H80250N_SETTING = "1F 1B 1F 92 10 11 12 15 16 17 55";

        //設定濃度命令開頭
        public static string DENSITY_SETTING_HEADER = "1F 1B 1F A9 06 05 04 03 02 01 ";

        //設定代碼頁命令開頭
        public static string CODEPAGE_SETTING_HEADER = "1F 1B 1F FF ";

        //打印代碼頁命令集
        public static byte[] CODEPAGE_PRINT_HEADER = { 0x1b, 0x40, 0x1c, 0x2e, 0x1b, 0x33, 0x6b, 0x1b, 0x74, };
        public static byte[] CODEPAGE_PRINT_SEPARATE = { 0x1d, 0x21, 0x11, };
        public static byte[] CODEPAGE_PRINT_CHAR1 = { 0x0a, 0x20, 0x20, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x0a, 0x38, 0x20,
                                                      0x80, 0x81, 0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8a, 0x8b, 0x8c, 0x8d, 0x8e, 0x8f, 0x0a, 0x39, 0x20,
                                                      0x90, 0x91, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, 0x99, 0x9a, 0x9b, 0x9c, 0x9d, 0x9e, 0x9f, 0x0a,
                                                      0x41, 0x20, 0xa0, 0xa1, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6, 0xa7, 0xa8, 0xa9, 0xaa, 0xab, 0xac, 0xad, 0xae, 0xaf, 0x0a, 0x42, 0x20, 0xb0, 0xb1, 0xb2,
                                                      0xb3, 0xb4, 0xb5, 0xb6, 0xb7, 0xb8, 0xb9, 0xba, 0xbb, 0xbc, 0xbd, 0xbe, 0xbf, 0x0a, 0x43, 0x20, 0xc0, 0xc1,
                                                      0xc2, 0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xcb, 0xcc, 0xcd, 0xce, 0xcf, 0x0a, 0x44, 0x20, 0xd0, 0xd1, 0xd2, 0xd3, 0xd4,
                                                      0xd5, 0xd6, 0xd7, 0xd8, 0xd9, 0xda, 0xdb, 0xdc, 0xdd, 0xde, 0xdf, 0x0a, 0x45, 0x20, 0xe0, 0xe1, 0xe2, 0xe3,
                                                      0xe4, 0xe5, 0xe6, 0xe7, 0xe8, 0xe9, 0xea, 0xeb, 0xec, 0xed, 0xee, 0xef, 0x0a, 0x46, 0x20, 0xf0, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6,
                                                      0xf7, 0xf8, 0xf9, 0xfa, 0xfb, 0xfc, 0xfd, 0xfe, 0xff, 0x0a, 0x1b, 0x40, 0x1c, 0x2e, 0x1b, 0x33, 0x50, 0x1b, 0x21, 0x01, 0x1d, 0x21,0x11};
        public static byte[] CODEPAGE_PRINT_CHAR2 = { 0x0a, 0x20, 0x20, 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x0a, 0x38, 0x20, 0x80, 0x81,
                                                      0x82, 0x83, 0x84, 0x85, 0x86, 0x87, 0x88, 0x89, 0x8a, 0x8b, 0x8c, 0x8d, 0x8e, 0x8f, 0x0a, 0x39, 0x20, 0x90, 0x91, 0x92, 0x93, 0x94, 0x95, 0x96,
                                                      0x97, 0x98, 0x99, 0x9a, 0x9b, 0x9c, 0x9d, 0x9e, 0x9f, 0x0a, 0x41, 0x20, 0xa0, 0xa1, 0xa2, 0xa3, 0xa4, 0xa5, 0xa6, 0xa7, 0xa8, 0xa9, 0xaa, 0xab,
                                                      0xac, 0xad, 0xae, 0xaf, 0x0a, 0x42, 0x20, 0xb0, 0xb1, 0xb2, 0xb3, 0xb4, 0xb5, 0xb6, 0xb7, 0xb8, 0xb9, 0xba, 0xbb, 0xbc, 0xbd, 0xbe, 0xbf, 0x0a,
                                                      0x43, 0x20, 0xc0, 0xc1, 0xc2, 0xc3, 0xc4, 0xc5, 0xc6, 0xc7, 0xc8, 0xc9, 0xca, 0xcb, 0xcc, 0xcd, 0xce, 0xcf, 0x0a, 0x44, 0x20, 0xd0, 0xd1, 0xd2,
                                                      0xd3, 0xd4, 0xd5, 0xd6, 0xd7, 0xd8, 0xd9, 0xda, 0xdb, 0xdc, 0xdd, 0xde, 0xdf, 0x0a, 0x45, 0x20, 0xe0, 0xe1, 0xe2, 0xe3, 0xe4, 0xe5, 0xe6, 0xe7,
                                                      0xe8, 0xe9, 0xea, 0xeb, 0xec, 0xed, 0xee, 0xef, 0x0a, 0x46, 0x20, 0xf0, 0xf1, 0xf2, 0xf3, 0xf4, 0xf5, 0xf6, 0xf7, 0xf8, 0xf9, 0xfa, 0xfb, 0xfc, 0xfd,
                                                      0xfe, 0xff, 0x0a, 0x1d, 0x56, 0x42, 0x00, 0x1b, 0x40 };

        //設定MAC地址顯示模式
        public static string MAC_SHOW_HEX_SETTING = "1F 1B 1F A5 04 05 06 09 08 07 22";
        public static string MAC_SHOW_DEC_SETTING = "1F 1B 1F A5 04 05 06 09 08 07 23";

        //設定二維碼功能
        public static string QRCODE_ON_SETTING = "1F 1B 1F 94 01 02 03 11 12 13 33";
        public static string QRCODE_OFF_SETTING = "1F 1B 1F 94 01 02 03 11 12 13 55";

        //設定DIP開關
        public static string DIP_ON_SETTING = "1F 1B 1F C0 15 14 13 12 11 10 AA";
        public static string DIP_OFF_SETTING = "1F 1B 1F C0 15 14 13 12 11 10 77";

        //設定紙寬
        public static string PAPER_WIDTH_80MM_SETTING = "1F 1B 1F A4 01 02 03 11 12 13 55";
        public static string PAPER_WIDTH_58MM_SETTING = "1F 1B 1F A4 01 02 03 11 12 13 33";

        //設定馬達加速度命令開頭
        public static string ACCELERATION_OF_MOTOR_SETTING = "1F 1B 1F 16 19 16 18 15 17 14 ";

        //設定馬達加速開關
        public static string ACCELERATION_OF_MOTOR_OFF_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 14 02 00";
        public static string ACCELERATION_OF_MOTOR_ON_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 14 02 01";

        //設定打印速度
        public static string PRINT_SPEED_200_SETTING = "1F 1B 1F 96 01 02 03 11 12 13 66";
        public static string PRINT_SPEED_250_SETTING = "1F 1B 1F 96 01 02 03 11 12 13 55";
        public static string PRINT_SPEED_300_SETTING = "1F 1B 1F 96 01 02 03 11 12 13 33";


        //設定定制字體
        public static string CUSTOMIZED_FONT_OFF_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 29 02 00";
        public static string CUSTOMIZED_FONT_ON_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 29 02 01";

        //設定濃度模式
        public static string DENSITY_MODE_LOW_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 26 02 00";
        public static string DENSITY_MODE_HIGH_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 26 02 01";

        //設定紙盡重打
        public static string PAPEROUT_REPRINT_OFF_SETTING = "1F 1B 1F F4 11 22 33 33";
        public static string PAPEROUT_REPRINT_ON_SETTING = "1F 1B 1F F4 11 22 33 55";

        //設定合蓋後自動切紙
        public static string HEADCLOSE_AUTOCUT_OFF_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 17 02 00";
        public static string HEADCLOSE_AUTOCUT_ON_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 17 02 01";

        //設定垂直移動單位
        public static string Y_OFFSET_1_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 18 02 00";
        public static string Y_OFFSET_05_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 18 02 01";

        //設定自檢頁logo是否打印
        public static string LOGO_PRINT_OFF_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 20 02 00";
        public static string LOGO_PRINT_ON_SETTING = "1F 1B 1F 53 5A 4A 42 5A 46 31 20 02 01";

        //設定DIP值HEAER
        public static string DIP_VALUE_SETTING_HEADER = "1F 1B 1F A1 10 11 12 13 14 15";

        //讀取所有欄位header
        public static string READ_ALL_HEADER = "1F 1B 1F 53 5A 4A 42 5A 46";

        //收到IP的命令分類
        public static string RE_IP_CLASSFY = "30-10";

        //收到Gateway的命令分類
        public static string RE_GATEWAY_CLASSFY = "30-11";

        //收到Mac的命令分類
        public static string RE_MAC_CLASSFY = "30-12";

        //收到自動斷線的命令分類
        public static string RE_AUTODISCONNECT_CLASSFY = "30-13";

        //收到連線數的命令分類
        public static string RE_CLIENTCOUNT_CLASSFY = "30-14";

        //收到網速的命令分類
        public static string RE_NETWORK_SPEED_CLASSFY = "30-15";

        //收到DHCP模式的命令分類
        public static string RE_DHCP_MODE_CLASSFY = "30-16";

        //收到USB MODE的命令分類
        public static string RE_USB_MODE_CLASSFY = "30-17";

        //收到USB端口的命令分類
        public static string RE_USB_FIX_CLASSFY = "30-18";

        //收到設置代碼頁的命令分類
        public static string RE_CODEPAGE_CLASSFY = "31-36";

        //收到設置語言的命令分類
        public static string RE_LANGUAGES_CLASSFY = "31-28";

        //收到設置FONTB的命令分類
        public static string RE_FONTB_CLASSFY = "31-23";

        //收到設置定制字型的命令分類
        public static string RE_CUSTOMFONT_CLASSFY = "31-29";

        //收到設置走紙方向的命令分類
        public static string RE_DIRECTION_CLASSFY = "31-13";

        //收到設置馬達加速與否的命令分類
        public static string RE_MOTOR_ACC_CONTROL_CLASSFY = "31-14";

        //收到設置馬達加速度的命令分類
        public static string RE_MOTOR_ACC_CLASSFY = "31-15";

        //收到設置打印速度的命令分類
        public static string RE_PRINT_SPEED_CLASSFY = "31-31";

        //收到設置濃度模式的命令分類
        public static string RE_DENSITY_MODE_CLASSFY = "31-26";

        //收到設置濃度調節的命令分類
        public static string RE_DENSITY_CLASSFY = "31-27";

        //收到紙盡重打的命令分類
        public static string RE_PAPEROUT_CLASSFY = "31-21";

        //收到打印紙寬的命令分類
        public static string RE_PAPERWIDTH_CLASSFY = "31-30";

        //收到合蓋後自動切紙的命令分類
        public static string RE_HEADCLOSE_CUT_CLASSFY = "31-17";

        //收到垂直移動單位的命令分類
        public static string RE_YOFFSET_CLASSFY = "31-18";

        //收到mac地址顯示的命令分類
        public static string RE_MACSHOW_CLASSFY = "31-24";

        //收到二維碼功能的命令分類
        public static string RE_QRCODE_CLASSFY = "31-25";

        //收到自檢頁logo設置的命令分類
        public static string RE_LOGOPRINT_CLASSFY = "31-20";

        //收到dip開關的命令分類
        public static string RE_DIPSW_CLASSFY = "31-34";

        //收到dip值的命令分類
        public static string RE_DIPVALUE_CLASSFY = "31-35";
    }
}

