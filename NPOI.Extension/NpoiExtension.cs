using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using NPOI.HPSF;
using NPOI.HSSF.Util;

namespace NPOI.Extension
{
    public static class NpoiExtension
    {
        private static IWorkbook ToWorkebook<T>(this IEnumerable<T> source)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            var properties = typeof(T).GetProperties();
            var attributes = GetColumnAttributes(properties);

            var workbook = InitializeWorkbook();
            var sheet = workbook.CreateSheet();

            var firstRowStyle = CreateFirstRowStyle(workbook);
            sheet.CreateFirstRow(properties, attributes, firstRowStyle);

            sheet.CreateRows(source, properties, attributes);

            return workbook;
        }

        private static ColumnAttribute[] GetColumnAttributes(PropertyInfo[] properties)
        {
            var attributes = new ColumnAttribute[properties.Length];

            for (int i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                var attrs = property.GetCustomAttributes(typeof(ColumnAttribute), true) as ColumnAttribute[];
                attributes[i] = attrs != null && attrs.Length > 0 ? attrs[0] : null;
            }

            return attributes;
        }

        private static HSSFWorkbook InitializeWorkbook()
        {
            var workbook = new HSSFWorkbook();

            workbook.DocumentSummaryInformation = CreateDocumentSummaryInfo();
            workbook.SummaryInformation = CreateSummaryInfo();

            return workbook;
        }

        private static DocumentSummaryInformation CreateDocumentSummaryInfo()
        {
            var documentSummaryInfo = PropertySetFactory.CreateDocumentSummaryInformation();
            documentSummaryInfo.Company = "cnki";
            return documentSummaryInfo;
        }

        private static SummaryInformation CreateSummaryInfo()
        {
            var summaryInfomation = PropertySetFactory.CreateSummaryInformation();
            summaryInfomation.Author = "sqb";
            summaryInfomation.Subject = "npoi extension";
            return summaryInfomation;
        }

        private static void CreateFirstRow(this ISheet sheet, PropertyInfo[] properties, ColumnAttribute[] attributes, ICellStyle cellStyle)
        {
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                var excel = attributes[i];
                if (excel == null)
                    continue;

                var cell = row.CreateCell(excel.Index);
                cell.CellStyle = cellStyle;
                cell.SetCellValue(excel.Title);
            }
        }

        private static ICellStyle CreateFirstRowStyle(IWorkbook workebook)
        {
            var style = workebook.CreateCellStyle();

            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.FillForegroundColor = HSSFColor.White.Index;
            style.FillBackgroundColor = HSSFColor.Grey40Percent.Index;
            style.FillPattern = FillPattern.Bricks;

            return style;
        }

        private static void CreateRows<T>(this ISheet sheet, IEnumerable<T> source, PropertyInfo[] properties, ColumnAttribute[] attributes)
        {
            var rowIndex = 1;
            foreach (var item in source)
            {
                var row = sheet.CreateRow(rowIndex);

                for (int i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];
                    var excel = attributes[i];
                    if (excel == null)
                        continue;

                    row.CreateCell(excel.Index, property, item);

                }
                rowIndex++;
            }
        }

        private static void CreateCell<T>(this IRow row, int columnIndex, PropertyInfo property, T item)
        {
            var cell = row.CreateCell(columnIndex);
            var value = property.GetValue(item, null);

            cell.SetPropertyValue(value, property);
        }

        private static void SetPropertyValue(this ICell cell, object value, PropertyInfo property)
        {
            if (value == null)
                cell.SetCellValue(string.Empty);
            else
            {
                if (value is ValueType)
                {
                    if (property.PropertyType == typeof(bool))
                        cell.SetCellValue((bool)value);
                    else if (property.PropertyType == typeof(DateTime))
                        cell.SetCellValue(Convert.ToDateTime(value).ToString("yyyy-MM-dd"));
                    else
                        cell.SetCellValue(Convert.ToDouble(value));
                }
                else
                    cell.SetCellValue(value + "");
            }

        }
    }
}
