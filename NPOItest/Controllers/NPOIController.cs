using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOItest.Models.Sevices;

namespace NPOItest.Controllers
{
    public class NPOIController : Controller
    {
        private string fileSavedPath = "~/Content/";
        private ExcelServices excelServices = new ExcelServices();

        // GET: NPOI
        public void EmptyExport()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.xls"), FileMode.Open, FileAccess.ReadWrite);
            HSSFWorkbook templateWorkbook = excelServices.AccountEmpty(fs);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountEmpty.xls"));
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);
            ms.WriteTo(Response.OutputStream);
            Response.End();
        }

        // GET: NPOI
        public void DataExport()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.xls"), FileMode.Open, FileAccess.ReadWrite);
            HSSFWorkbook templateWorkbook = excelServices.AccountData(fs);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountData.xls"));
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);
            ms.WriteTo(Response.OutputStream);
            Response.End();
        }

        // GET: NPOI
        public void PDFExport(int SchoolCtrl001Id = 0)
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.xls"), FileMode.Open, FileAccess.ReadWrite);

            HSSFWorkbook templateWorkbook = excelServices.AccountData(fs);

            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);

            string target = string.Concat(Server.MapPath(fileSavedPath), "/temp/" + System.Guid.NewGuid().ToString() + "暫存EXCEL.xls");//??
            using (var fileStream = new FileStream(target, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                ms.Position = 0;
                ms.CopyTo(fileStream); // fileStream is not populated
            }

            ConvertPDFHelper Convert = new ConvertPDFHelper();
            string pdfPath = string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp/" + System.Guid.NewGuid().ToString() + "轉檔.pdf");
            string PDFfile = Convert.ConvertExcelToPdf(target, pdfPath);

            Stream iStream = new FileStream(PDFfile, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            MemoryStream memoryStream = new MemoryStream();

            iStream.CopyTo(memoryStream);
            iStream.Dispose();

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccounrPDF.pdf"));

            memoryStream.WriteTo(Response.OutputStream);

            Response.End();
        }
    }
}