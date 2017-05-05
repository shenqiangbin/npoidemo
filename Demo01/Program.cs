using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
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
        ///     单元格下拉
        ///     todo:限制输入的长度，只能输入数字
        ///     单元格宽度自动适应
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

            //设置宽度
            table.SetColumnWidth(0, 16 * 256 + 200); // 200为常量，这样即可控制列宽为16
            table.SetColumnWidth(1, 16 * 256 + 200);
            table.SetColumnWidth(2, 16 * 256 + 200);

            AutoCellWidth(table, 4);

            List<string> listData = new List<string>();
            listData.AddRange(new string[] { "男", "女" });
            var tempSheet = workbook.CreateSheet("sexSheet");
            tempSheet.ProtectSheet("haha");

            listData.ForEach(m => tempSheet.CreateRow(listData.IndexOf(m)).CreateCell(0).SetCellValue(m));

            IName range = workbook.CreateName();
            range.RefersToFormula = string.Format("sexSheet !$A$1:$A${0}", listData.Count);
            range.NameName = "dicRange";

            CellRangeAddressList regions = new CellRangeAddressList(1, 65535, 3, 3);
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(range.NameName);
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            dataValidate.CreateErrorBox("错误", "请按右侧下拉箭头选择!");//不符合约束时的提示  
            dataValidate.ShowErrorBox = true;//显示上面提示 = True 
            table.AddValidationData(dataValidate);

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

        private static void AutoCellWidth(ISheet paymentSheet, int rowsCount)
        {
            for (int columnNum = 0; columnNum <= rowsCount; columnNum++)
            {
                int columnWidth = paymentSheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 1; rowNum <= paymentSheet.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //当前行未被使用过
                    if (paymentSheet.GetRow(rowNum) == null)
                    {
                        currentRow = paymentSheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = paymentSheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                paymentSheet.SetColumnWidth(columnNum, columnWidth * 256);
            }
        }
    }


}
