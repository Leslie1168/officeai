using System;

namespace OfficeAi
{
    /// <summary>
    /// 把图片当作“图片本身”插入到当前光标 / 选区位置。
    /// 返回普通文字消息；失败以 "ERR:" 开头。
    /// </summary>
    internal static class ImageInsert
    {
        public static string Insert(object app, Connect.HostKind kind, string filePath)
        {
            if (app == null) return "ERR:没有宿主程序对象";
            try
            {
                switch (kind)
                {
                    case Connect.HostKind.Word:   return InsertWord(app, filePath);
                    case Connect.HostKind.Excel:  return InsertExcel(app, filePath);
                    case Connect.HostKind.PowerPoint: return InsertPpt(app, filePath);
                    default: return "ERR:当前宿主不支持插入图片";
                }
            }
            catch (Exception e)
            {
                return "ERR:" + e.Message;
            }
        }

        private static string InsertWord(object app, string filePath)
        {
            dynamic a = app;
            dynamic sel = a.Selection;
            if (sel == null) return "ERR:没有光标位置";
            try { sel.InlineShapes.AddPicture(filePath); }
            catch (Exception e)
            {
                // 有些宿主把 Range.InlineShapes 当集合用，试试带 LinkToFile 的签名
                if (e.Message.Contains("AddPicture"))
                {
                    dynamic rng = sel.Range;
                    rng.InlineShapes.AddPicture(filePath);
                }
                else throw;
            }
            return "已将图片插入到光标处";
        }

        private static string InsertExcel(object app, string filePath)
        {
            dynamic a = app;
            dynamic cell = a.ActiveCell;
            dynamic shp = a.ActiveSheet.Shapes;
            // AddPicture(Filename, LinkToFile=0, SaveWithDocument=-1,
            //           Left, Top, Width, Height)
            // 放到活动单元格左上角
            int left = (int)(ActiveLeft(a));
            int top = (int)(ActiveTop(a));
            var pic = shp.AddPicture(filePath, 0 /*msoFalse*/, -1 /*msoTrue*/,
                                     left, top, -1, -1);
            try
            {
                // 等比缩放到列宽/行高附近，避免超大
                double cw = 120, rh = 120;
                try { cw = (double)cell.ColumnWidth * 6.5; } catch { }
                try { rh = (double)cell.RowHeight; } catch { }
                double scale = Math.Min(1.0, Math.Min(cw / pic.Width, 300 / pic.Width));
                if (scale < 1) { pic.Width = pic.Width * scale; pic.Height = pic.Height * scale; }
            }
            catch { }
            return "已将图片插入到当前单元格附近";
        }

        private static string InsertPpt(object app, string filePath)
        {
            dynamic a = app;
            dynamic win = a.ActiveWindow;
            dynamic slide = win.View.Slide;
            var shp = slide.Shapes.AddPicture(filePath, 0, -1, 60, 60, -1, -1);
            // 等比限宽
            try
            {
                double mw = (double)slide.Master.Width;
                if (shp.Width > mw * 0.8)
                {
                    double s = (mw * 0.8) / shp.Width;
                    shp.Width = shp.Width * s;
                    shp.Height = shp.Height * s;
                }
            }
            catch { }
            return "已将图片插入到当前幻灯片";
        }

        private static int ActiveLeft(object app) { try { return (int)((dynamic)app).ActiveCell.Left; } catch { return 20; } }
        private static int ActiveTop(object app) { try { return (int)((dynamic)app).ActiveCell.Top; } catch { return 20; } }
    }
}
