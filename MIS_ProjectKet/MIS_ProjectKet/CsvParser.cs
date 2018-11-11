using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace MIS_ProjectKet
{
    class CsvParser
    {
        public CsvParser()
        {
            //sets the existing excel file to be written
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(@"C:\Users\suare\Desktop\Project\test.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
             
            //use of library StreamReader to get csv file
            using (var sr = new StreamReader(@"C:\Users\suare\Desktop\Project\test.csv"))
            {
                //variable to be able to read csv file
                var reader = new CsvReader(sr);

                //csv configurations
                reader.Configuration.PrepareHeaderForMatch = header => Regex.Replace(header, @"\s", string.Empty);
                reader.Configuration.RegisterClassMap<DataRecordMap>();


                //CSVReader will now read the whole file into an enumerable
                IEnumerable<DataRecord> records = reader.GetRecords<DataRecord>();

                //CSV file will be printed to the Output Window
                int row = 1;

                
                foreach (DataRecord record in records)
                {
                    int col = 1;

                    while (col <= record.columnCount)
                    {
                        if (col == 1)
                            x.Cells[row, col] = record.CreateDate.Replace(" ", "");
                        else if (col == 2)
                            x.Cells[row, col] = record.IRInitialResponse.Replace("at", "@");
                        else if (col == 3)
                            x.Cells[row, col] = record.FRFixResponse.Replace("at", "@");
                        else if (col == 4)
                            x.Cells[row, col] = record.Summary;
                        else if (col == 5)
                            x.Cells[row, col] = record.TicketNo;
                        else if (col == 6)
                            x.Cells[row, col] = record.Company;
                        else if (col == 7)
                            x.Cells[row, col] = record.Assignedto;
                        else if (col == 8)
                            x.Cells[row, col] = record.Caller;
                        else if (col == 9)
                            x.Cells[row, col] = record.CaseType;
                        else if (col == 10)
                            x.Cells[row, col] = record.CloseDate;
                        else if (col == 11)
                            x.Cells[row, col] = record.DaysOpen;
                        else if (col == 12)
                            x.Cells[row, col] = record.Department;
                        else if (col == 13)
                            x.Cells[row, col] = record.InfrastructureType;
                        else if (col == 14)
                            x.Cells[row, col] = record.Location;
                        else if (col == 15)
                            x.Cells[row, col] = record.Priority;
                        else if (col == 16)
                            x.Cells[row, col] = record.IncidentType;
                        else if (col == 17)
                            x.Cells[row, col] = record.Resolution;
                        else if (col == 18)
                            x.Cells[row, col] = record.Status;
                        else if (col == 19)
                            x.Cells[row, col] = record.Shift;

                        col++;
                    }
                    Console.WriteLine("row: " + row++);
                    Console.WriteLine("Create Date: {0}, IRInitialResponse: {1}, FRFixResponse: {2}, Summary: {3}, TicketNo: {4}, " +
                        "Company: {5}, Assignedto: {6}, Caller: {7}, CaseType: {8}, CloseDate: {9}, DaysOpen: {10}, Department: {11}, InfrastructureType: {12}" +
                        ", Location: {13}, Priority: {14}, IncidentType: {15}, Resolution: {16}, Status: {17}, Shift: {18}", record.CreateDate.Replace(" ",""), record.IRInitialResponse, record.FRFixResponse,
                        record.Summary, record.TicketNo, record.Company, record.Assignedto, record.Caller, record.CaseType,
                        record.CloseDate, record.DaysOpen, record.Department, record.InfrastructureType, record.Location,
                        record.Priority, record.IncidentType, record.Resolution, record.Status, record.Shift);
                }
            }

            sheet.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            Console.WriteLine("FINISHED");
            Console.ReadKey();
        }

        public sealed class DataRecordMap : ClassMap<DataRecord>
        {
            public DataRecordMap()
            {
                AutoMap();
                Map(m => m.TicketNo).Name("Ticket #");
            }
        }
    }
}
