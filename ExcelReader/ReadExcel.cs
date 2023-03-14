using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;


namespace ExcelReader
{
    public class ExcelAPI
    {

        public ExcelAPI()
        {
            Console.WriteLine("Read excel constructor...");
        }

        public void ReadExcel()
        {
            string filePath = "C:\\Workspace\\C#\\ExcelReader\\Book1.xlsx";
            Application excel = new Application();

            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[1];

            excel.Visible = true;

            ws.Range["A1"].Value = 99;
        }
    }

}
