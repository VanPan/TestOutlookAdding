using System;
using System.Drawing;
using System.Windows.Forms;

namespace TestOutlookAddin
{
    internal class PictureConverter : AxHost
    {
        private PictureConverter() : base(String.Empty) { }

        static public stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }

        static public stdole.Picture ImageToPicture(Image image)
        {
            return (stdole.Picture)GetIPictureFromPicture(image);
        }

        static public stdole.IPictureDisp IconToPictureDisp(Icon icon)
        {
            return ImageToPictureDisp(icon.ToBitmap());
        }

        static public stdole.Picture IconToPicture(Icon icon)
        {
            return ImageToPicture(icon.ToBitmap());
        }

        static public Image PictureDispToImage(stdole.IPictureDisp picture)
        {
            return GetPictureFromIPicture(picture);
        }

    }
}
