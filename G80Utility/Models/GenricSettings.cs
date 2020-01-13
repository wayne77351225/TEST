using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PirnterUtility.Models
{
    [Serializable]
    class GenricSettings
    {
        public int Speed;
        public int Density;
        public int ThermalMode;
        public int LabelType;
        public double LabelWidth;
        public double LabelHeight;
        public double GapDistance;
        public double GapOffset;
        public double BLineThickness;
        public double BLineFeedLen;
        public double ContinueOffset;
        public int Direction;
        public int Mirror;
        public double ReferenceX;
        public double ReferenceY;
        public int DrawReverse;
        public int PostPrintAction;
        public int OnDemandCom;
        public int CutNumber;
        public int CutAction;
        public int CutMode;
        //public int BackAfterCut;
        public int SensorType;
        public int CoverClose;
    }

}
