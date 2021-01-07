using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace G80Utility.Tool
{
    class Config
    {
        public static bool isSetIPChecked;
        public static bool isSetGatewayChecked;
        public static bool isSetMacChecked;
        public static bool isAutoDisconnectChecked;
        public static bool isConnectClientChecked;
        public static bool isEthernetSpeedChecked;
        public static bool isDHCPModeChecked;
        public static bool isUSBModeChecked;
        public static bool isUSBFixedChecked;
        public static bool isCodePageSetChecked;

        public static bool isLanguageSetChecked;
        public static bool isFontBSettingChecked;
        public static bool isCustomziedFontChecked;
        public static bool isDirectionChecked;
        public static bool isMotorAccControlChecked;
        public static bool isAccMotorChecked;
        public static bool isPrintSpeedChecked;
        //public static bool isDensityModeChecked;

        public static bool isDensityChecked;
        public static bool isPaperOutReprintChecked;
        public static bool isPaperWidthChecked;
        public static bool isHeadCloseCutChecked;
        public static bool isYOffsetChecked;
        //public static bool isMACShowChecked;
        public static bool isLEDChecked;
        public static bool isQRCodeChecked;
        public static bool isLogoPrintControlhecked;
        public static bool isDIPSwitchChecked;
        public static bool isCutBeepChecked;

        //for maintail
        public static bool isCMDQRCodeMaintainChecked;
        public static bool isCMDGeneralMaintainChecked;
        public static bool isCMDPageMaintainChecked;
        public static bool isFeedLinesChecked;
        public static bool isPrintedLinesChecked;
        public static bool isCutPaperTimesChecked;
        public static bool isHeadOpenTimesChecked;
        public static bool isPaperOutTimesChecked;
        public static bool iErrorTimesChecked;

        //for factory
        public static bool isCMDQRCodeChecked;
        public static bool isCMDGeneralChecked;
        public static bool isCMDPageChecked;

        //for authority
        public static string ADMIN_PWD = ""; //shjbzfkj123
        public static string EDIT_SN_PWD = ""; //szxlh123
    }
}
