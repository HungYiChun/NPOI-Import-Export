using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOItest.Models.Sevices;
using NPOI.XWPF.UserModel;

namespace NPOItest.Controllers
{
    public class NPOIController : Controller
    {
        private string fileSavedPath = "~/Content/";
        private NPOIServices NPServices = new NPOIServices();

        // GET: NPOI Excel
        public void EmptyExport_E()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.xls"), FileMode.Open, FileAccess.ReadWrite);
            HSSFWorkbook templateWorkbook = NPServices.AccountEmpty_E(fs);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountEmpty.xls"));
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);
            ms.WriteTo(Response.OutputStream);
            Response.End();
        }

        // GET: NPOI Excel
        public void DataExport_E()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.xls"), FileMode.Open, FileAccess.ReadWrite);
            HSSFWorkbook templateWorkbook = NPServices.AccountData_E(fs);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountData.xls"));
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);
            ms.WriteTo(Response.OutputStream);
            Response.End();
        }

        // GET: NPOI Excel
        public void PDFExport_E()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.xls"), FileMode.Open, FileAccess.ReadWrite);

            HSSFWorkbook templateWorkbook = NPServices.AccountData_E(fs);

            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);

            string target = string.Concat(Server.MapPath(fileSavedPath), "/temp/" + System.Guid.NewGuid().ToString() + "EXCEL.xls");//??
            using (var fileStream = new FileStream(target, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                ms.Position = 0;
                ms.CopyTo(fileStream); // fileStream is not populated
            }

            ConvertPDFHelper Convert = new ConvertPDFHelper();
            string pdfPath = string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp/" + System.Guid.NewGuid().ToString() + ".pdf");
            string PDFfile = Convert.ConvertExcelToPdf(target, pdfPath);

            Stream iStream = new FileStream(PDFfile, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            MemoryStream memoryStream = new MemoryStream();

            iStream.CopyTo(memoryStream);
            iStream.Dispose();

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountPDF.pdf"));

            memoryStream.WriteTo(Response.OutputStream);

            Response.End();
        }

        // POST: NPOI Excel
        [HttpPost]
        public ActionResult Import_E(HttpPostedFileBase file)
        {
            string message;
            if (file != null && file.ContentLength > 0 && file.ContentLength < (10 * 1024 * 1024))
            {
                string filetype = file.FileName.Split('.').Last();
                string fileName = Path.GetFileName(file.FileName);
                string path = Path.Combine(Server.MapPath("~/Content/Imports"), fileName);
                if (filetype == "xls")
                {
                    file.SaveAs(path);
                    FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                    HSSFWorkbook excel = new HSSFWorkbook(fs);

                    message = NPServices.InsertData_E(excel);
                }
                else
                {
                    message = "File format error !";
                }
            }
            else
            {
                message = "Please select file import !";
            }
            ViewBag.Message = message;

            return View();
        }


        // GET: NPOI Word
        public void EmptyExport_W()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.docx"), FileMode.Open, FileAccess.ReadWrite);
            XWPFDocument templateWorkbook = NPServices.AccountEmpty_W(fs);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountEmpty.docx"));
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);
            ms.WriteTo(Response.OutputStream);
            Response.End();
        }

        // GET: NPOI Word
        public void DataExport_W()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.docx"), FileMode.Open, FileAccess.ReadWrite);
            XWPFDocument templateWorkbook = NPServices.AccountData_W(fs);

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountData.docx"));
            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);
            ms.WriteTo(Response.OutputStream);
            Response.End();
        }

        // GET: NPOI Word
        public void PDFExport_W()
        {
            FileStream fs = new FileStream(string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp.docx"), FileMode.Open, FileAccess.ReadWrite);
            XWPFDocument templateWorkbook = NPServices.AccountData_W(fs);

            MemoryStream ms = new MemoryStream();
            templateWorkbook.Write(ms);

            string target = string.Concat(Server.MapPath(fileSavedPath), "/temp/" + System.Guid.NewGuid().ToString() + "Word.docx");//??
            using (var fileStream = new FileStream(target, FileMode.CreateNew, FileAccess.ReadWrite))
            {
                ms.Position = 0;
                ms.CopyTo(fileStream); // fileStream is not populated
            }

            ConvertPDFHelper Convert = new ConvertPDFHelper();
            string pdfPath = string.Concat(Server.MapPath(fileSavedPath), "/Excels/temp/" + System.Guid.NewGuid().ToString() + ".pdf");
            string PDFfile = Convert.ConvertExcelToPdf(target, pdfPath);

            Stream iStream = new FileStream(PDFfile, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            MemoryStream memoryStream = new MemoryStream();

            iStream.CopyTo(memoryStream);
            iStream.Dispose();

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats";
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Server.UrlPathEncode("AccountPDF.pdf"));

            memoryStream.WriteTo(Response.OutputStream);

            Response.End();
        }

        // POST: NPOI Word
        [HttpPost]
        public ActionResult Import_W(HttpPostedFileBase file)
        {
            string message;
            if (file != null && file.ContentLength > 0 && file.ContentLength < (10 * 1024 * 1024))
            {
                string filetype = file.FileName.Split('.').Last();
                string fileName = Path.GetFileName(file.FileName);
                string path = Path.Combine(Server.MapPath("~/Content/Imports"), fileName);
                if (filetype == "docx")
                {
                    file.SaveAs(path);
                    FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
                    XWPFDocument word = new XWPFDocument(fs);

                    message = NPServices.InsertData_W(word);
                }
                else
                {
                    message = "File format error !";
                }
            }
            else
            {
                message = "Please select file import !";
            }
            ViewBag.Message = message;

            return View();
        }
    }
}