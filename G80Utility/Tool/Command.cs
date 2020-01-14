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

        //設定MAC地址顯示模式
        public static string MAC_SHOW_HEX_SETTING = "1F 1B 1F A5 04 05 06 09 08 07 22";
        public static string MAC_SHOW_DEC_SETTING = "1F 1B 1F A5 04 05 06 09 08 07 23";

        //設定二維碼功能
        public static string QRCODE_ON_SETTING = "1F 1B 1F 94 01 02 03 11 12 13 33";
        public static string QRCODE_OFF_SETTING = "1F 1B 1F 94 01 02 03 11 12 13 55";

        //設定DIP開關
        public static string DIP_ON_SETTING = "1F 1B 1F 94 01 02 03 11 12 13 33";
        public static string DIP_OFF_SETTING = "1F 1B 1F 94 01 02 03 11 12 13 55";

        //設定紙寬
        public static string PAPER_WIDTH_80MM_SETTING = "1F 1B 1F A4 01 02 03 11 12 13 55";
        public static string PAPER_WIDTH_58MM_SETTING = "1F 1B 1F A4 01 02 03 11 12 13 33";

        //設定馬達加速度命令開頭
        public static string ACCELERATION_OF_MOTOR_SETTING = "1F 1B 1F 16 19 16 18 15 17 14 ";
    
        //設定馬達加速開關
        //public static string ACCELERATION_OF_MOTOR_OEPN_SETTING = "1F 1B 1F 16 19 16 18 15 17 14 ";
    }
}

