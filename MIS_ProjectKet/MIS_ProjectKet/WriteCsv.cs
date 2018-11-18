using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MIS_ProjectKet
{
    class WriteCsv
    {
        public WriteCsv(DataTable dt)
        {
            //sets the existing excel file to be written
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(@"C:\Users\suare\Desktop\Project\test.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            x = (Excel.Worksheet)sheet.Sheets[2];

            int rowCount = 1;
            int dataColumns = dt.Columns.Count;

            foreach (DataRow dr in dt.Rows)
            {
                int columnCount = 0;
                while (columnCount < dataColumns)
                {
                    x.Cells[rowCount, columnCount + 1] = dr[columnCount];
                    columnCount++;
                }
                rowCount++;

            }
            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
        }
    }
}
