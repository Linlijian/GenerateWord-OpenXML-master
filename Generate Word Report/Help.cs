using System;
using System.Drawing;
using System.IO;

namespace Generate_Word_Report
{
    public static class Help
    {
        public static string horizontalAlignmentLeft = "left";
        public static string horizontalAlignmentCenter = "center";
        public static string horizontalAlignmentRight = "right";
        
        public static string paragraphUnderline = "U";
        public static string paragraphItalic = "I";
        public static string paragraphBold = "B";

        public static string wrapSquare = "WrapSquare";
        public static string wrapTopBottom = "WrapTopBottom";
        public static string wrapNone = "WrapNone";
        public static string wrapThrough = "wrapThrough";
        public static string wrapTight = "WrapTight";


        public static Bitmap Base64StringToBitmap(this string base64String)
        {
            Bitmap bmpReturn = null;

            byte[] byteBuffer = Convert.FromBase64String(base64String);
            MemoryStream memoryStream = new MemoryStream(byteBuffer);

            memoryStream.Position = 0;

            bmpReturn = (Bitmap)Bitmap.FromStream(memoryStream);

            memoryStream.Close();
            memoryStream = null;
            byteBuffer = null;


            return bmpReturn;
        }
    }

}
