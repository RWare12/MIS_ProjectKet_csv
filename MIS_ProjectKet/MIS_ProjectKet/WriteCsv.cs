using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace MIS_ProjectKet
{
    class WriteCsv
    {
        public WriteCsv(DataTable dt)
        {
            //sets the existing excel file to be written
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(@"C:\Users\suare\Desktop\Project\pivottable.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;


            //selects a specific worksheet to written on
            x = (Excel.Worksheet)sheet.Sheets[2];

            string[,] data = new string[dt.Rows.Count, dt.Columns.Count];
            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                int j = 0;
                foreach (DataColumn col in dt.Columns)
                {
                    data[i, j++] = row[col].ToString();
                }
                i++;
            }

            int topRow = 2;
            int topColumn = 2;
            Excel.Range c1 = (Excel.Range)x.Cells[topRow, topColumn];
            Excel.Range c2 = (Excel.Range)x.Cells[topRow + dt.Rows.Count - 1, topColumn + dt.Columns.Count - 1];
            Excel.Range range = x.Range[c1, c2];

            //Excel.PivotCache cache = sheet.PivotCaches();
            //Excel.PivotCache cache = (Excel.PivotCache)sheet.PivotCaches().Add(dt,range);  // Set the Source data range from First sheet
            //Excel.PivotCache cache = 
            range.Value = data;
            range.EntireRow.WrapText = false;

            //System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            //Marshal.ReleaseComObject(sheet);
            //Marshal.ReleaseComObject(x);
            //Marshal.ReleaseComObject(c1);
            //Marshal.ReleaseComObject(c2);
            //Marshal.ReleaseComObject(range);
            sheet.Close();
            excel.Quit();
            //excel = null;

        }
    }
}
