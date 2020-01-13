using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PirnterUtility.Tool
{
    class UnitChange
    {
        public static int PrinterDPI { get; set; }

        public static double PixeltoMillim(double pix)
        {
            double min = 0;
            switch (PrinterDPI)
            {
                case 203:
                    min = Math.Round(pix / 8.0, 1, MidpointRounding.AwayFromZero);
                    break;
                case 300:
                    min = Math.Round(pix / 12.0, 1, MidpointRounding.AwayFromZero);
                    break;
                case 600:
                    min = Math.Round(pix / 24.0, 1, MidpointRounding.AwayFromZero);
                    break;
            }


            return min;
        }

        public static double PixeltoCm(double pix)
        {
            double cm = 0;
            switch (PrinterDPI)
            {
                case 203:
                    cm = Math.Round(pix / 80.0, 2, MidpointRounding.AwayFromZero);
                    break;
                case 300:
                    cm = Math.Round(pix / 120.0, 2, MidpointRounding.AwayFromZero);
                    break;
                case 600:
                    cm = Math.Round(pix / 240.0, 2, MidpointRounding.AwayFromZero);
                    break;
            }

            return cm;
        }

        public static double PixeltoInch(double pix)
        {
            double inch = 0;
            switch (PrinterDPI)
            {
                case 203:
                    inch = Math.Round(pix / 203.0, 3, MidpointRounding.AwayFromZero);
                    break;
                case 300:
                    inch = Math.Round(pix / 300.0, 3, MidpointRounding.AwayFromZero);
                    break;
                case 600:
                    inch = Math.Round(pix / 600.0, 3, MidpointRounding.AwayFromZero);
                    break;
            }
            return inch;
        }

        public static int MillimtoPixel(double mm)
        {
            int pixel = 0;
            switch (PrinterDPI)
            {
                case 203:
                    pixel = (int)Math.Round(mm * 8, 0, MidpointRounding.AwayFromZero);
                    break;
                case 300:
                    pixel = (int)Math.Round(mm * 12, 0, MidpointRounding.AwayFromZero);
                    break;
                case 600:
                    pixel = (int)Math.Round(mm * 24, 0, MidpointRounding.AwayFromZero);
                    break;
            }
            Console.WriteLine("mm:" + pixel);
            return pixel;
        }

        public static int CmtoPixel(double cm)
        {
            int pixel = 0;
            switch (PrinterDPI)
            {
                case 203:
                    pixel = (int)Math.Round(cm * 80, 0, MidpointRounding.AwayFromZero);
                    break;
                case 300:
                    pixel = (int)Math.Round(cm * 120, 0, MidpointRounding.AwayFromZero);
                    break;
                case 600:
                    pixel = (int)Math.Round(cm * 240, 0, MidpointRounding.AwayFromZero);
                    break;
            }
            Console.WriteLine("cm:" + pixel);
            return pixel;
        }

        public static int InchtoPixel(double inch)
        {
            int pixel = 0;
            switch (PrinterDPI)
            {
                case 203:
                    pixel = (int)Math.Round(inch * 203, 0, MidpointRounding.AwayFromZero);
                    break;
                case 300:
                    pixel = (int)Math.Round(inch * 300, 0, MidpointRounding.AwayFromZero);
                    break;
                case 600:
                    pixel = (int)Math.Round(inch * 600, 0, MidpointRounding.AwayFromZero);
                    break;
            }
            Console.WriteLine("inch:" + pixel);
            return pixel;
        }

        public static string OtherstoPixel(double others, string unitType)
        {
            double result = 0;
            switch (unitType)
            {
                case "mm":
                    result = MillimtoPixel(others);
                    break;
                case "inch":
                    result = InchtoPixel(others);
                    break;
                case "cm":
                    result = CmtoPixel(others);
                    break;
            }
            Console.WriteLine("result:" + result);
            return result.ToString();
        }

        public static string OtherstoMin(double others, string unitType)
        {
            double result = 0;
            switch (unitType)
            {
                case "dot":
                    result = PixeltoMillim(others);
                    break;
                case "cm":                  
                    result = PixeltoMillim(CmtoPixel(others));
                    break;
                case "inch":
                    result = PixeltoMillim(InchtoPixel(others));
                    break;
            }
            Console.WriteLine("result:" + result);
            return result.ToString();
        }

        public static string OtherstoCm(double others, string unitType)
        {
            double result = 0;
            switch (unitType)
            {
                case "dot":
                    result = PixeltoCm(others);
                    break;
                case "mm":
                    result = PixeltoCm(MillimtoPixel(others));
                    break;
                case "inch":
                    result = PixeltoCm(InchtoPixel(others));
                    break;
            }
            Console.WriteLine("result:" + result);
            return result.ToString();
        }

        public static string OtherstoInch(double others, string unitType)
        {
            double result = 0;
            switch (unitType)
            {
                case "dot":
                    result = PixeltoInch(others);
                    break;
                case "mm":
                    result = PixeltoInch(MillimtoPixel(others));
                    break;
                case "cm":
                    result = PixeltoInch(CmtoPixel(others));
                    break;
            }
            Console.WriteLine("result:" + result);
            return result.ToString();
        }
    }
}
