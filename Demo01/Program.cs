using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Demo01
{
    class Program
    {
        /// <summary>
        /// 涉及的内容有：
        ///     excel表的创建
        ///     设置单元格的的背景色颜色
        ///     设置字体的大小，粗体
        ///     锁定单元格不能编辑
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var workbook = new HSSFWorkbook();            
            var table = workbook.CreateSheet("sheetName");
            table.ProtectSheet("nicai");            
            var row = table.CreateRow(0);

            var cell = row.CreateCell(0);
            cell.SetCellValue("编号");
            SetColor(workbook, cell, 128, 128, 128);

            cell = row.CreateCell(1);
            cell.SetCellValue("姓名");
            SetColor(workbook, cell, 128, 128, 128);

            cell = row.CreateCell(2);
            cell.SetCellValue("性别");
            SetColor(workbook, cell, 128, 128, 128);

            cell = row.CreateCell(3);
            cell.SetCellValue("年龄");
            SetColor(workbook, cell, 128, 128, 128);

            var cellStyle = workbook.CreateCellStyle();
            cellStyle.IsLocked = false;
            cellStyle.SetFont(GetCommonFont(workbook));
            table.SetDefaultColumnStyle(0, cellStyle);
            table.SetDefaultColumnStyle(1, cellStyle);
            table.SetDefaultColumnStyle(2, cellStyle);
            table.SetDefaultColumnStyle(3, cellStyle);

            using (var fs = File.OpenWrite(@"d:/demo.xls"))
            {
                workbook.Write(fs);
                Console.WriteLine("ok");
            }
        }
        public static void SetColor(HSSFWorkbook workbook, ICell cell, int red, int green, int blue)
        {
            var cellStyle = workbook.CreateCellStyle();
            cellStyle.FillPattern = FillPattern.SolidForeground;// 老版本可能这样写FillPatternType.SOLID_FOREGROUND;
            HSSFPalette palette = workbook.GetCustomPalette(); //调色板实例
            palette.SetColorAtIndex((short)8, (byte)184, (byte)204, (byte)228);
            var hssFColor = palette.FindColor((byte)red, (byte)green, (byte)blue);
            cellStyle.FillForegroundColor = hssFColor.Indexed;

            cellStyle.SetFont(GetFont(workbook));

            cell.CellStyle = cellStyle;//设置
        }

        private static IFont GetFont(HSSFWorkbook workbook)
        {
            var font = workbook.CreateFont();
            font.IsBold = true;
            font.FontHeightInPoints = 12;
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            //font.Boldweight = 20;

            return font;
        }

        private static IFont GetCommonFont(HSSFWorkbook workbook)
        {
            var font = workbook.CreateFont();            
            font.FontHeightInPoints = 12;
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;
            //font.Boldweight = 20;

            return font;
        }
    }


}
