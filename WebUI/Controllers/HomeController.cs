using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Web;
using System.Web.Mvc;
using Utility;

namespace WebUI.Controllers
{
    public class HomeController : Controller
    {
        // GET: Default
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Download()
        {
            ExcelTemplate template = new ExcelTemplate("人员导入模板");
            template.EnableProtect = true;
            template.Columns.Add(new ExcelColumn { Index = 1, ColName = "编号", ColWidth = 20 });
            template.Columns.Add(new ExcelColumn { Index = 0, ColName = "姓名", ColWidth = 20 });
            template.Columns.Add(new ExcelColumn { Index = 3, ColName = "性别", ColWidth = 20, DataSource = new List<string> { "男", "女" } });
            template.Columns.Add(new ExcelColumn { Index = 2, ColName = "年级", ColWidth = 20, DataSource = new List<string> { "1班", "2班", "3班" } });

            return File(template.ExportToStream(), "application/octet-stream", template.FileName);
        }

        public ActionResult DownloadPersonInfo()
        {
            ExcelHelper excelHelper = new ExcelHelper();
            //Stream stream = excelHelper.DataTable2Excel(excelHelper.GetFakeTable(), "sheet0");
            Stream stream = excelHelper.List2Excel(excelHelper.GetFakeList(), "sheet0", new string[] { "编号", "姓名", "出生日期" });

            return File(stream, "application/octet-stream", "人员.xls");
        }

        public ActionResult ImportTest()
        {
            string path = "D:/web测试1.xls";
            ExcelHelper excelHelper = new ExcelHelper();
            var table = excelHelper.Excel2DataTable(path);
            return Content("ok");
        }
    }
}