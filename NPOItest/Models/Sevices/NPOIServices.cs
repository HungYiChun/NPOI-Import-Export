using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace NPOItest.Models.Sevices
{
    public class NPOIServices
    {
        private NPOIModel db = new NPOIModel();

        public HSSFWorkbook AccountEmpty_E(FileStream fs)
        {
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs);
            HSSFSheet ws = (HSSFSheet)templateWorkbook.GetSheetAt(0);

            ws.GetRow(1).GetCell(2).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));

            return templateWorkbook;
        }

        public HSSFWorkbook AccountData_E(FileStream fs)
        {
            List<Account> data = db.Account.ToList();
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs);
            HSSFSheet ws = (HSSFSheet)templateWorkbook.GetSheetAt(0);

            ws.GetRow(1).GetCell(2).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));

            int startRow = 3;
            int i = 1;
            foreach (Account item in data)
            {
                ws.GetRow(startRow).GetCell(0).SetCellValue(i.ToString());
                ws.GetRow(startRow).GetCell(1).SetCellValue(item.Username);
                ws.GetRow(startRow).GetCell(2).SetCellValue(item.Name);
                ws.GetRow(startRow).GetCell(3).SetCellValue(item.Email);
                ws.GetRow(startRow).GetCell(4).SetCellValue(item.Sex);
                ws.GetRow(startRow).GetCell(5).SetCellValue(item.Company);
                ws.GetRow(startRow).GetCell(6).SetCellValue(item.Position);
                ws.GetRow(startRow).GetCell(7).SetCellValue(item.Phone);
                startRow++;
                i++;
            }
            ws.ShiftRows(ws.LastRowNum + 1, ws.LastRowNum + (ws.LastRowNum - startRow + 1), -(ws.LastRowNum - startRow + 1));

            return templateWorkbook;
        }

        public string InsertData_E(HSSFWorkbook excel)
        {
            HSSFSheet ws = (HSSFSheet)excel.GetSheetAt(0);
            List<Account> newAccounts = new List<Account>();
            int startRow = 3;
            for (int i = startRow; i <= ws.LastRowNum; i++)
            {
                newAccounts.Add(new Account
                {
                    Username = ws.GetRow(startRow).GetCell(1).StringCellValue,
                    Password = "520520",
                    Name = ws.GetRow(startRow).GetCell(2).StringCellValue,
                    Email = ws.GetRow(startRow).GetCell(3).StringCellValue,
                    Sex = ws.GetRow(startRow).GetCell(4).StringCellValue,
                    Company = ws.GetRow(startRow).GetCell(5).StringCellValue,
                    Position = ws.GetRow(startRow).GetCell(6).StringCellValue,
                    Phone = ws.GetRow(startRow).GetCell(7).StringCellValue
                });
                startRow++;
            }
            db.Account.AddRange(newAccounts);
            db.SaveChanges();

            return "Success !";
        }

        public XWPFDocument AccountEmpty_W(FileStream fs)
        {
            XWPFDocument templateWorkbook = new XWPFDocument(fs);
            XWPFTable tb = templateWorkbook.Tables[0];

            tb.GetRow(0).GetCell(1).SetText(DateTime.Now.ToString("yyyy/MM/dd"));

            return templateWorkbook;
        }

        public XWPFDocument AccountData_W(FileStream fs)
        {
            List<Account> data = db.Account.ToList();
            XWPFDocument templateWorkbook = new XWPFDocument(fs);
            XWPFTable tb = templateWorkbook.Tables[0];

            tb.GetRow(0).GetCell(1).SetText(DateTime.Now.ToString("yyyy/MM/dd"));
            int startRow = 3;
            int i = 1;
            foreach (Account item in data)
            {
                tb.CreateRow().CreateCell();
                tb.GetRow(startRow).GetCell(0).SetText(i.ToString());
                tb.GetRow(startRow).GetCell(1).SetText(item.Username);
                tb.GetRow(startRow).GetCell(2).SetText(item.Name);
                tb.GetRow(startRow).CreateCell();
                tb.GetRow(startRow).GetCell(3).SetText(item.Email);
                tb.GetRow(startRow).CreateCell();
                tb.GetRow(startRow).GetCell(4).SetText(item.Sex);
                tb.GetRow(startRow).CreateCell();
                tb.GetRow(startRow).GetCell(5).SetText(item.Company);
                tb.GetRow(startRow).CreateCell();
                tb.GetRow(startRow).GetCell(6).SetText(item.Position);
                tb.GetRow(startRow).CreateCell();
                tb.GetRow(startRow).GetCell(7).SetText(item.Phone);
                startRow++;
                i++;
            }

            return templateWorkbook;
        }

        public string InsertData_W(XWPFDocument word)
        {
            XWPFTable tb = word.Tables[0];
            List<Account> newAccounts = new List<Account>();
            int startRow = 3;
            for (int i = startRow; i <= 5; i++)
            {
                newAccounts.Add(new Account
                {
                    Username = tb.GetRow(startRow).GetCell(1).GetText(),
                    Password = "520520",
                    Name = tb.GetRow(startRow).GetCell(2).GetText(),
                    Email = tb.GetRow(startRow).GetCell(3).GetText(),
                    Sex = tb.GetRow(startRow).GetCell(4).GetText(),
                    Company = tb.GetRow(startRow).GetCell(5).GetText(),
                    Position = tb.GetRow(startRow).GetCell(6).GetText(),
                    Phone = tb.GetRow(startRow).GetCell(7).GetText()
                });
                startRow++;
            }
            db.Account.AddRange(newAccounts);
            db.SaveChanges();

            return "Success !";
        }
    }
}