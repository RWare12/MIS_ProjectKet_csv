using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;


namespace MIS_ProjectKet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var sr = new StreamReader(@"C:\Users\suare\Desktop\Project\test.csv"))
            {
                var reader = new CsvReader(sr);

                //CSVReader will now read the whole file into an enumerable
                IEnumerable<DataRecord> records = reader.GetRecords<DataRecord>();

                //CSV file will be printed to the Output Window
                int row = 0;
                foreach (DataRecord record in records)
                {
                    Console.WriteLine("row: " + row++);
                    Console.WriteLine("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}" +
                        ", {13}, {14}, {15}, {16}, {17}", record.CreateDate, record.IRInitialResponse, record.FRFixResponse,
                        record.Summary, record.TicketNo, record.Company, record.Assignedto, record.Caller, record.CaseType,
                        record.CloseDate, record.DaysOpen, record.Department, record.InfrastructureType, record.Location,
                        record.Priority, record.IncidentType, record.Resolution, record.Shift);
                    
                }
            }


            Console.WriteLine("FINISHED");
            Console.ReadKey();
        }

    }
}
