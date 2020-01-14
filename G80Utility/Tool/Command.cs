using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace G80Utility.Tool
{
    class Command
    {
        //RS232接口測試
        public static string RS232_COMMUNICATION_TEST = "1F1B100100";

        //重啟打印機
        public static string RESTART = "1F1B1F535A4A425A46110000";

        //走紙方向
        public static string Direction_80250N = "1F1B1F9210111215161733";
        public static string Direction_H80250N = "1F1B1F9210111215161755";

    }
}
