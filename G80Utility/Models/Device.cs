using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace G80Utility.Models
{
    class Device
    {
        public string DeviceType { get; set; } //設定目前選定種類
        public string RS232PortName { get; set; }
        public string USBSN { get; set; }
        public string USBPortNumber { get; set; }
        public string USBPortDescritption { get; set; }
        public string USBPortName { get; set; }
        public string USBVIDPID { get; set; }
        public Boolean USBisLinked { get; set; }
        public string USBDeviceInstance { get; set; }
        public string WIFIIP { get; set; }
        public string WIFIPort { get; set; }
        public string DisplayName
        {
            get
            {
                string display=null;
                switch (DeviceType) //依據不同傳輸類別取得相關資料
                {
                    case "usb":
                        if (USBisLinked)
                        {
                            //display = $"{USBPortName} [{USBPortDescritption} SN:{USBSN}] Linked";
                            //display = $"{USBPortDescritption} [{USBPortName}]";
                            //display = USBPortDescritption+" ["+USBPortName+"]";
                            display = USBPortName;
                        }
                       
                        break;
                    case "wifi"://目前沒有
                        display = WIFIIP; //$"{WIFIIP}:{WIFIPort}
                        break;
                    case "rs232":
                        display = RS232PortName;
                        break;
                    default:
                        break;
                }

                return display;


            }

        }
    }
}
