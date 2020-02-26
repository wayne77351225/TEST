using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Media.Imaging;

namespace G80Utility.Tool
{
    partial class BitmapTool
    {
        #region 取得bitmap的xL/xH/yL/yH HexString
        public static string getBmpHighandLowHex(int bmpWidth, int bmpHeight)
        {
            int widthDot = bmpWidth / 8;
            int heightDot = bmpHeight / 8;
            byte[] data = new byte[4];
            data[0] = (byte)(widthDot & 0xFF);
            data[1] = (byte)((widthDot >> 8) & 0xff);
            data[2] = (byte)(heightDot & 0xFF);
            data[3] = (byte)((heightDot >> 8) & 0xff);
            string hexResult = BitConverter.ToString(data).Replace("-", "");
            return hexResult;
        }
        #endregion

        #region 判斷bitmap範圍 
        //80mm的紙寬，200dpi的解析度=>1mm=8 dots=>max width=80*8=640
        public static bool checkBitmapRange(string filename ,int bmpWidth, int bmpHeight)
        {
            if (bmpHeight > 288 * 8)
            { //高超過 
              // 高 = (yL + yH × 256)*8 ,   (yL + yH × 256)≤288  => 高 ≤288*8
                MessageBox.Show(Application.Current.FindResource("ExceedMaxHeight") as string);
                return false;
            }
            if (bmpWidth > 640)
            { //寬超過
                MessageBox.Show(Application.Current.FindResource("ExceedMaxWidth") as string);
                return false;
            }
            //FileInfo fi = new FileInfo(filename);
            int totalPixs = bmpWidth * bmpHeight;
            if (totalPixs > 64000) { //大小超過flash下載限制
                MessageBox.Show(Application.Current.FindResource("ExceedMaxSize") as string);
                return false;
            }

            return true;
        }
        #endregion

        #region bitmap換算成hex string
        public static string bitmapToHexString(Bitmap bmp)
        {
            //epson command 公式:k=(xL+xH*256)*(yL+yH*256)*8
            int length = ((bmp.Width + 7) / 8) * ((bmp.Height + 7) / 8) * 8;
            int height = (bmp.Height + 7) / 8;
            string[] data = new string[length];

            //參考espson cmd圖片，dot的轉換從上到下，從左到右
            //一個dot有八個pixel
            for (int w = 0; w < bmp.Width; w++) //遍歷寬
            {
                for (int h = 0; h < ((bmp.Height + 7) / 8); h++) //遍歷高/8bits
                {
                    string binary = null;
                    for (int l = 0; l < 8; l++)
                    {
                        if (((h * 8) + l) < bmp.Height)  // if within the BMP size //y軸點
                        {
                            Color test = bmp.GetPixel(w, (h * 8) + l);
                            if (test.Name.Contains("000000"))//黑色
                            {
                                binary += "1";
                            }
                            else//白色
                            {
                                binary += "0";
                            }
                        }
                    }
                    data[w * height + h] = Convert.ToInt16(binary, 2).ToString("X2");
                }
            }
            string result = BitmapHexStringZeroStuffing(data);
            return result;
        }
        #endregion

        #region bitmapHexString為null時補0
        private static String BitmapHexStringZeroStuffing(string[] data) //null要轉00
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == null)
                {
                    data[i] = "00 ";
                    sb.Append(data[i] );
                }
                else
                {
                    sb.Append(data[i]);
                }
            }
            return sb.ToString();
        }
        #endregion

        #region 彩色圖片轉灰階
        public static Bitmap ToGray(Bitmap bmp, int mode)
        {
            if (bmp == null)
            {
                return null;
            }
            int w = bmp.Width;
            int h = bmp.Height;
            try
            {
                byte newColor = 0;
                BitmapData srcData = bmp.LockBits(new System.Drawing.Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
                unsafe
                {
                    byte* p = (byte*)srcData.Scan0.ToPointer();
                    for (int y = 0; y < h; y++)
                    {
                        for (int x = 0; x < w; x++)
                        {

                            if (mode == 0) // 加權平均
                            {
                                newColor = (byte)((float)p[0] * 0.114f + (float)p[1] * 0.587f + (float)p[2] * 0.299f);
                            }
                            else// 算數平均
                            {
                                newColor = (byte)((float)(p[0] + p[1] + p[2]) / 3.0f);
                            }
                            p[0] = newColor;
                            p[1] = newColor;
                            p[2] = newColor;

                            p += 3;
                        }
                        p += srcData.Stride - w * 3;
                    }
                    bmp.UnlockBits(srcData);
                    return bmp;
                }
            }
            catch
            {
                return null;
            }
        }
        #endregion

        #region 彩色圖片轉黑白
        public static Bitmap Thresholding(Bitmap img1)
        {
            int[] histogram = new int[256];
            int minGrayValue = 255, maxGrayValue = 0;
            //求取直方圖
            for (int i = 0; i < img1.Width; i++)
            {
                for (int j = 0; j < img1.Height; j++)
                {
                    System.Drawing.Color pixelColor = img1.GetPixel(i, j);
                    histogram[pixelColor.R]++;
                    if (pixelColor.R > maxGrayValue) maxGrayValue = pixelColor.R;
                    if (pixelColor.R < minGrayValue) minGrayValue = pixelColor.R;
                }
            }
            //叠代計算閥值
            int threshold = -1;
            int newThreshold = (minGrayValue + maxGrayValue) / 2;
            for (int iterationTimes = 0; threshold != newThreshold && iterationTimes < 100; iterationTimes++)
            {
                threshold = newThreshold;
                int lP1 = 0;
                int lP2 = 0;
                int lS1 = 0;
                int lS2 = 0;
                //求兩個區域的灰度的平均值
                for (int i = minGrayValue; i < threshold; i++)
                {
                    lP1 += histogram[i] * i;
                    lS1 += histogram[i];
                }
                // int mean1GrayValue = (lP1 / lS1);
                int mean1GrayValue = 0;
                if (lP1 != 0)
                {
                    mean1GrayValue = (lP1 / lS1);
                }
                else
                {
                    mean1GrayValue = 20;
                }
                for (int i = threshold + 1; i < maxGrayValue; i++)
                {
                    lP2 += histogram[i] * i;
                    lS2 += histogram[i];
                }
                //int mean2GrayValue = (lP2 / lS2);
                int mean2GrayValue = 0;
                if (lP2 != 0)
                {
                    mean2GrayValue = (lP2 / lS2);
                }
                else
                {
                    mean2GrayValue = 180;
                }
                newThreshold = (mean1GrayValue + mean2GrayValue) / 2;
            }
            //計算二值化
            for (int i = 0; i < img1.Width; i++)
            {
                for (int j = 0; j < img1.Height; j++)
                {
                    System.Drawing.Color pixelColor = img1.GetPixel(i, j);
                    if (pixelColor.R > threshold) img1.SetPixel(i, j, System.Drawing.Color.FromArgb(255, 255, 255));
                    else img1.SetPixel(i, j, System.Drawing.Color.FromArgb(0, 0, 0));
                }
            }
            return img1;
        }
        #endregion

        #region 判斷圖片是否為黑白
        public static bool isBlackWhite(Bitmap bmp)
        {
            System.Drawing.Color c = new System.Drawing.Color();

            //檢查每個pixel
            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    c = bmp.GetPixel(x, y);

                    //判斷像素點有包含非(0,0,0)或非(255,255,255)的點就非單色圖片
                    if (!(c.R == 0 && c.G == 0 && c.B == 0) && !(c.R == 255 && c.G == 255 && c.B == 255))
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion

        #region Bitmap --> BitmapImage
        public static BitmapImage BitmapToBitmapImage(Bitmap bitmap)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png); // 坑点：格式选Bmp时，不带透明度

                stream.Position = 0;
                BitmapImage result = new BitmapImage();
                result.BeginInit();
                // According to MSDN, "The default OnDemand cache option retains access to the stream until the image is needed."
                // Force the bitmap to load right now so we can dispose the stream.
                result.CacheOption = BitmapCacheOption.OnLoad;
                result.StreamSource = stream;
                result.EndInit();
                result.Freeze();
                return result;
            }
        }
        #endregion

        #region BitmapImage --> Bitmap
        public static Bitmap BitmapImageToBitmap(BitmapImage bitmapImage)
        {
            // BitmapImage bitmapImage = new BitmapImage(new Uri("../Images/test.png", UriKind.Relative));

            using (MemoryStream outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapImage));
                enc.Save(outStream);
                Bitmap bitmap = new Bitmap(outStream);

                return new Bitmap(bitmap);
            }
        }
        #endregion
    }
}
