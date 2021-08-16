using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace G80Utility.Models
{
    [Serializable]
    class ParaSettings
    {
        public string IpAddress;
        public string Gateway;
        public string MacAddress;
        public int AutoDisconnectIndex;
        public int ConnectClientIndex;
        public int EthernetSpeedIndex;
        public int DHCPModeIndex;
        public int USBModeIndex;
        public int USBFixedIndex;
        public int CodePageSetIndex;
        public int LanguageSetIndex;
        public int FontBSettingtIndex;
        public int CustomziedFontIndex;
        public int SetDirectionIndex;
        public int MotorAccControlIndex;
        public int AccMotorIndex;
        public int PrintSpeedIndex;
        //public int DensityModeIndex;
        public int DensityIndex;
        public int PaperOutReprintIndex;
        //public int PaperWidthIndex;
        public int DrawerIndex;
        public int HeadCloseCutIndex;
        public int YOffsetIndex;
        //public int MACShowIndex;
        public int LEDIndex;
        public int QRCodeIndex;
        public int LogoPrintControlIndex;
        public int DIPSwitchIndex;
        public bool CutterCheck;
        public bool BeepCheck;
        public bool DensityCheck;
        public bool ChineseForbiddenCheck;
        public bool CharNumberCheck;
        public bool CashboxCheck;
        public int DIPBaudRateComIndex;
        public int CutBeepEnable;
        public int CutBeepTimes;
        public int CutBeepDuration;
        public bool PaperWidthCheck;
    }      
    
}
