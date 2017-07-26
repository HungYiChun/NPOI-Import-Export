using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace NPOItest.Models.Sevices
{
    public class ExcelServices
    {
        private NPOIModel db = new NPOIModel();

        public HSSFWorkbook AccountEmpty(FileStream fs)
        {
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs);
            HSSFSheet ws = (HSSFSheet)templateWorkbook.GetSheetAt(0);

            ws.GetRow(1).GetCell(2).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));

            return templateWorkbook;
        }

        public HSSFWorkbook AccountData(FileStream fs)
        {
            List<Account> data = db.Account.ToList();
            HSSFWorkbook templateWorkbook = new HSSFWorkbook(fs);
            HSSFSheet ws = (HSSFSheet)templateWorkbook.GetSheetAt(0);

            ws.GetRow(1).GetCell(2).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd"));

            int startRow = 4;
            int i = 1;
            foreach (Account item in data)
            {
                ws.GetRow(startRow).GetCell(1).SetCellValue(i.ToString());
                ws.GetRow(startRow).GetCell(2).SetCellValue(item.Username);
                ws.GetRow(startRow).GetCell(3).SetCellValue(item.Name);
                ws.GetRow(startRow).GetCell(4).SetCellValue(item.Email);
                ws.GetRow(startRow).GetCell(5).SetCellValue(item.Sex);
                ws.GetRow(startRow).GetCell(6).SetCellValue(item.Company);
                ws.GetRow(startRow).GetCell(7).SetCellValue(item.Position);
                ws.GetRow(startRow).GetCell(8).SetCellValue(item.Phone);
                startRow++;
                i++;
            }
            startRow++;
            while (i < 55)
            {
                ws.ShiftRows(startRow, ws.LastRowNum, -1);
                i++;
            }
            ws.ShiftRows(startRow, ws.LastRowNum, -1);

            return templateWorkbook;
        }
    }
}