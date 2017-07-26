using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace NPOItest.Models.Sevices
{
    public class ConvertPDFHelper
    {
        public string ConvertExcelToPdf(string inputFile, string pdfPath)
        {
            Application excelApp = new Application();
            excelApp.Visible = false;
            Workbook workbook = null;
            Workbooks workbooks = null;


            try
            {
                workbooks = excelApp.Workbooks;
                workbook = workbooks.Open(inputFile);
                workbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                                             pdfPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                             Type.Missing, Type.Missing);
            }
            finally
            {

                if (workbook != null)
                {
                    workbook.Close(XlSaveAction.xlDoNotSaveChanges);
                    while (Marshal.FinalReleaseComObject(workbook) != 0) { };
                    workbook = null;
                }

                if (workbooks != null)
                {
                    workbooks.Close();
                    while (Marshal.FinalReleaseComObject(workbooks) != 0) { };
                    workbooks = null;
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    excelApp.Application.Quit();
                    while (Marshal.FinalReleaseComObject(excelApp) != 0) { };
                    excelApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return pdfPath;          
        }
    }
}