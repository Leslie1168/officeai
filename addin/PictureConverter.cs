using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficeAi
{
    /// <summary>
    /// 把 Bitmap 转成 COM 的 IPictureDisp，供 Ribbon getImage 回调返回。
    /// AxHost.GetIPictureDispFromPicture 是 protected static，只能靠派生类暴露。
    /// </summary>
    [ComVisible(false)]
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base(string.Empty) { }

        public static object ToPictureDisp(Image image)
        {
            return GetIPictureDispFromPicture(image);
        }
    }
}
