using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Demo01
{
    public class ExcelTemplate
    {
        public string FileName { get; set; }
        public IList<ExcelColumn> Columns { get; set; }
        /// <summary>
        /// 是否启用保护
        /// </summary>

        public bool EnableProtect { get; set; }
        private HSSFWorkbook _workbook;

        public ExcelTemplate(string name)
        {
            //this.FileName = string.Format("{0}_{1}.xls", name, DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss"));
            this.FileName = string.Format("{0}.xls", name);
            this.Columns = new List<ExcelColumn>();
        }

        public void Save(string path)
        {
            Directory.CreateDirectory(path);
            string savePath = Path.Combine(path, this.FileName);

            InitWorkbook();

            using (var fs = File.Open(savePath, FileMode.OpenOrCreate))
            {
                _workbook.Write(fs);
            }
        }

        public HttpResponseMessage ExportToWeb()
        {
            InitWorkbook();

            Stream stream = new MemoryStream();
            _workbook.Write(stream);

            System.Globalization.CultureInfo myCItrad = new System.Globalization.CultureInfo("ZH-CN", true);
            HttpResponseMessage response = new HttpResponseMessage { Content = new StreamContent(stream) };
            //response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            //{
            //    FileName = getFileName(),
            //};
            response.Headers.Add("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(this.FileName, Encoding.UTF8));

            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            response.Content.Headers.ContentLength = stream.Length;

            return response;
        }

        //private string getFileName()
        //{            
        //    var agent = HttpContext.Current.Request.UserAgent;
        //    if (agent.IndexOf("Firefox") == -1)//非火狐浏览器 
        //    {
        //        return HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8);
        //    }

        //    return this.FileName;
        //}


        private void InitWorkbook()
        {
            _workbook = new HSSFWorkbook();

            var table = _workbook.CreateSheet();

            if (this.EnableProtect)
                table.ProtectSheet(Guid.NewGuid().ToString());

            table.CreateFreezePane(0, 1);

            CreateColumnTitles(table);
            SetColumnStyle(table);
        }

        private void CreateColumnTitles(ISheet sheet)
        {
            var row = sheet.CreateRow(0);

            foreach (var col in this.Columns)
            {
                var cell = row.CreateCell(col.Index);
                cell.SetCellValue(col.ColName);
                cell.CellStyle = GetHeaderStyle(128, 128, 128);
            }
        }

        private ICellStyle GetHeaderStyle(int red, int green, int blue)
        {
            var cellStyle = _workbook.CreateCellStyle();
            cellStyle.FillPattern = FillPattern.SolidForeground;// 老版本可能这样写FillPatternType.SOLID_FOREGROUND;
            HSSFPalette palette = _workbook.GetCustomPalette(); //调色板实例
            palette.SetColorAtIndex((short)8, (byte)184, (byte)204, (byte)228);
            var hssFColor = palette.FindColor((byte)red, (byte)green, (byte)blue);
            cellStyle.FillForegroundColor = hssFColor.Indexed;

            cellStyle.SetFont(GetHeaderFont());

            return cellStyle;
        }

        private IFont GetHeaderFont()
        {
            var font = _workbook.CreateFont();
            font.IsBold = true;
            font.FontHeightInPoints = 12;
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            //font.Boldweight = 20;

            return font;
        }

        private void SetColumnStyle(ISheet sheet)
        {
            var cellStyle = GetRowStyle();
            foreach (var col in this.Columns)
            {
                sheet.SetDefaultColumnStyle(col.Index, cellStyle);

                if (col.ColWidth.HasValue)
                    sheet.SetColumnWidth(col.Index, col.ColWidth.Value * 256 + 200);

                if (col.DataSource != null)
                    SetColDataSource(sheet, col);
            }
        }

        private ICellStyle GetRowStyle()
        {
            var cellStyle = _workbook.CreateCellStyle();
            cellStyle.IsLocked = false;
            cellStyle.SetFont(GetCommonFont());

            return cellStyle;
        }

        private IFont GetCommonFont()
        {
            var font = _workbook.CreateFont();
            font.FontHeightInPoints = 12;
            font.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Normal;
            //font.Boldweight = 20;

            return font;
        }

        private void SetColDataSource(ISheet sheet, ExcelColumn col)
        {
            var sheetName = string.Format("{0}DataSource", col.ColName);
            var tempSheet = _workbook.CreateSheet(sheetName);
            _workbook.SetSheetHidden(_workbook.GetSheetIndex(sheetName), true);
            tempSheet.ProtectSheet(Guid.NewGuid().ToString());

            col.DataSource.ForEach(m => tempSheet.CreateRow(col.DataSource.IndexOf(m)).CreateCell(0).SetCellValue(m));

            IName range = _workbook.CreateName();
            range.RefersToFormula = string.Format("{0} !$A$1:$A${1}", sheetName, col.DataSource.Count);
            range.NameName = string.Format("{0}range", col.ColName);

            CellRangeAddressList regions = new CellRangeAddressList(1, 65535, col.Index, col.Index);
            DVConstraint constraint = DVConstraint.CreateFormulaListConstraint(range.NameName);
            HSSFDataValidation dataValidate = new HSSFDataValidation(regions, constraint);
            dataValidate.CreateErrorBox("错误", "请按右侧下拉箭头选择!");//不符合约束时的提示  
            dataValidate.ShowErrorBox = true;//显示上面提示 = True 

            sheet.AddValidationData(dataValidate);
        }
    }

    public class ExcelColumn
    {
        public int Index { get; set; }
        public string ColName { get; set; }
        public int? ColWidth { get; set; }
        public List<string> DataSource { get; set; }
    }
}
