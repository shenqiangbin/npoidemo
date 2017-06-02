using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Utility
{
    public class ExcelHelper
    {
        /// <summary>
        /// 类版本
        /// </summary>
        public string version
        {
            get { return "0.1"; }
        }
        readonly int EXCEL03_MaxRow = 65535;

        /// <summary>
        /// 将DataTable转换为excel2003格式。
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public Stream DataTable2Excel(DataTable dt, string sheetName)
        {
            IWorkbook book = new HSSFWorkbook();
            if (dt.Rows.Count < EXCEL03_MaxRow)
                DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName);
            else
            {
                int page = dt.Rows.Count / EXCEL03_MaxRow;
                for (int i = 0; i < page; i++)
                {
                    int start = i * EXCEL03_MaxRow;
                    int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
                    DataWrite2Sheet(dt, start, end, book, sheetName + i.ToString());
                }
                int lastPageItemCount = dt.Rows.Count % EXCEL03_MaxRow;
                DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName + page.ToString());
            }
            MemoryStream ms = new MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return ms;
        }
        private void DataWrite2Sheet(DataTable dt, int startRow, int endRow, IWorkbook book, string sheetName)
        {
            ISheet sheet = book.CreateSheet(sheetName);
            IRow header = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = header.CreateCell(i);
                string val = dt.Columns[i].Caption ?? dt.Columns[i].ColumnName;
                cell.SetCellValue(val);
            }
            int rowIndex = 1;
            for (int i = startRow; i <= endRow; i++)
            {
                DataRow dtRow = dt.Rows[i];
                IRow excelRow = sheet.CreateRow(rowIndex++);
                for (int j = 0; j < dtRow.ItemArray.Length; j++)
                {
                    SetCellVal(excelRow.CreateCell(j), dtRow[j], dt.Columns[j]);
                }
            }
        }

        private void SetCellVal(ICell cell, object val, DataColumn col)
        {
            if (col.DataType == typeof(DateTime))
                cell.SetCellValue(DateTime.Parse(val.ToString()).ToShortDateString());
            else if (col.DataType == typeof(System.DBNull))
                cell.SetCellValue("");
            else
                cell.SetCellValue(val.ToString());
        }

        public Stream List2Excel<T>(IEnumerable<T> list, string sheetName, string[] titles)
        {
            IWorkbook book = new HSSFWorkbook();
            if (list.Count() < EXCEL03_MaxRow)
                DataWrite2Sheet(list, 0, list.Count() - 1, book, sheetName, titles);
            else
            {
                int page = list.Count() / EXCEL03_MaxRow;
                for (int i = 0; i < page; i++)
                {
                    int start = i * EXCEL03_MaxRow;
                    int end = (i * EXCEL03_MaxRow) + EXCEL03_MaxRow - 1;
                    DataWrite2Sheet(list, start, end, book, sheetName + i.ToString(), titles);
                }
                int lastPageItemCount = list.Count() % EXCEL03_MaxRow;
                DataWrite2Sheet(list, list.Count() - lastPageItemCount, lastPageItemCount, book, sheetName + page.ToString(), titles);
            }
            MemoryStream ms = new MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return ms;
        }

        public void DataWrite2Sheet<T>(IEnumerable<T> list, int startRow, int endRow, IWorkbook book, string sheetName, string[] titles)
        {
            ISheet sheet = book.CreateSheet(sheetName);

            IRow header = sheet.CreateRow(0);
            for (int i = 0; i < titles.Count(); i++)
            {
                ICell cell = header.CreateCell(i);
                cell.SetCellValue(titles[i]);
            }

            int rowIndex = 1;
            for (int i = startRow; i <= endRow; i++)
            {
                var item = list.ElementAt(i);
                IRow excelRow = sheet.CreateRow(rowIndex++);
                SetExcelRowVal<T>(excelRow, item);
            }
        }

        private void SetExcelRowVal<T>(IRow excelRow, T item)
        {
            var properties = typeof(T).GetProperties();
            for (int i = 0; i < properties.Length; i++)
            {
                var property = properties[i];
                var val = property.GetValue(item);

                //ICell cell = ;
                //cell.SetCellValue(val.ToString());
                SetPropertyValue(excelRow.CreateCell(i), val, property);
            }
        }

        private void SetPropertyValue(ICell cell, object value, PropertyInfo property)
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

        /// <summary>读取excel
        /// 默认第一行为标头
        /// </summary>
        /// <param name="strFileName">excel文档路径</param>
        /// <returns></returns>
        public DataTableResult Excel2DataTable(string strFileName)
        {
            DataTableResult result = new DataTableResult();
            DataTable dt = new DataTable();

            HSSFWorkbook hssfworkbook;
            using (FileStream file = new FileStream(strFileName, FileMode.Open, FileAccess.Read))
            {
                hssfworkbook = new HSSFWorkbook(file);
            }
            HSSFSheet sheet = hssfworkbook.GetSheetAt(0) as HSSFSheet;
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            HSSFRow headerRow = sheet.GetRow(0) as HSSFRow;
            int cellCount = headerRow.LastCellNum;

            for (int j = 0; j < cellCount; j++)
            {
                HSSFCell cell = headerRow.GetCell(j) as HSSFCell;
                dt.Columns.Add(cell.ToString());
            }

            for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
            {
                HSSFRow row = sheet.GetRow(i) as HSSFRow;
                DataRow dataRow = dt.NewRow();

                bool isEmpty = true;
                for (int j = row.FirstCellNum; j < cellCount; j++)
                {
                    if (row.GetCell(j) != null && !string.IsNullOrEmpty(row.GetCell(j).ToString()))
                    {
                        dataRow[j] = row.GetCell(j).ToString();
                        isEmpty = false;
                    }

                }
                if (!isEmpty)
                    dt.Rows.Add(dataRow);
            }

            result.DataTable = dt;

            Dictionary<string, string> dataDic = new Dictionary<string, string>();

            var dataDicSheet = hssfworkbook.GetSheet("DataDic") as HSSFSheet;
            if (dataDicSheet != null)
            {
                for (int i = 0; i <= dataDicSheet.LastRowNum; i++)
                {
                    HSSFRow row = dataDicSheet.GetRow(i) as HSSFRow;
                    dataDic.Add(row.GetCell(0).ToString(), row.GetCell(1).ToString());
                }
            }

            result.DataDic = dataDic;

            return result;
        }

        public DataTable GetFakeTable()
        {
            DataTable table = new DataTable();


            table.Columns.Add(new DataColumn("id", typeof(int)).Caption = "编号");
            table.Columns.Add(new DataColumn("name", typeof(string)).Caption = "姓名");
            table.Columns.Add(new DataColumn("birthday", Type.GetType("System.DateTime")));

            DataRow row = table.NewRow();
            row["编号"] = 1;
            row["姓名"] = "小王";
            row["birthday"] = DateTime.Parse("1991-11-14");
            table.Rows.Add(row);

            row = table.NewRow();
            row["编号"] = 2;
            row["姓名"] = "小李";
            row["birthday"] = DateTime.Parse("1991-12-13");
            table.Rows.Add(row);

            return table;
        }

        public IEnumerable<Person> GetFakeList()
        {
            return new List<Person>
            {
                new Person { Id = 1, Name = "小王", Birthday =  DateTime.Parse("1991-11-14")},
                new Person { Id = 3, Name = "小李", Birthday =  DateTime.Parse("1991-12-13")},
            };
        }
    }

    public class DataTableResult
    {
        public DataTable DataTable { get; set; }
        public Dictionary<string,string> DataDic { get; set; }
    }

    public class Person
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public DateTime Birthday { get; set; }
    }
}
