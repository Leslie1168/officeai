using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace OfficeAi
{
    /// <summary>把图片真身放进剪贴板（CF_DIB / CF_BITMAP），供用户 Ctrl+V 粘贴为图片。</summary>
    internal static class ImageBoard
    {
        public static string CopyToClipboard(string filePath)
        {
            try
            {
                using (var bmp = new Bitmap(filePath))
                {
                    // 转成 32bppArgb 的深层副本，避免剪贴板引用文件句柄导致文件被占用
                    var copy = new Bitmap(bmp.Width, bmp.Height, PixelFormat.Format32bppArgb);
                    using (var g = Graphics.FromImage(copy))
                        g.DrawImage(bmp, 0, 0, bmp.Width, bmp.Height);

                    var dobj = new DataObject();
                    dobj.SetData(DataFormats.Bitmap, true, copy);
                    try { dobj.SetData(DataFormats.Dib, true, copy); } catch { }
                    Clipboard.SetDataObject(dobj, true);
                    // 让剪贴板持有引用，等它真正放进系统剪贴板再释放
                    GC.KeepAlive(copy);
                    return "1";
                }
            }
            catch (Exception e)
            {
                Connect.Log("ImageBoard.CopyToClipboard 异常: " + e.Message);
                return "0";
            }
        }
    }
}
