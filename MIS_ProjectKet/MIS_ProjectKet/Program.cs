using System;
using System.Collections.Generic;
using System.IO;
using CsvHelper;
using System.Text.RegularExpressions;
using CsvHelper.Configuration;

namespace MIS_ProjectKet
{
    class Program
    {
        static void Main(string[] args)
        {
            //use of library StreamReader to get csv file
            using (var sr = new StreamReader(@"C:\Users\suare\Desktop\Project\report.csv"))
            {
                //variable to be able to read csv file
                var reader = new CsvReader(sr);

                //csv configurations
                reader.Configuration.PrepareHeaderForMatch = header => Regex.Replace(header, @"\s", string.Empty);
                reader.Configuration.RegisterClassMap<DataRecordMap>();
               

                //CSVReader will now read the whole file into an enumerable
                IEnumerable<DataRecord> records = reader.GetRecords<DataRecord>();

                //CSV file will be printed to the Output Window
                int row = 0;
                foreach (DataRecord record in records)
                {
                    Console.WriteLine("row: " + row++);
                    Console.WriteLine("Create Date: {0}, IRInitialResponse: {1}, FRFixResponse: {2}, Summary: {3}, TicketNo: {4}, " +
                        "Company: {5}, Assignedto: {6}, Caller: {7}, CaseType: {8}, CloseDate: {9}, DaysOpen: {10}, Department: {11}, InfrastructureType: {12}" +
                        ", Location: {13}, Priority: {14}, IncidentType: {15}, Resolution: {16}, Shift: {17}", record.CreateDate, record.IRInitialResponse, record.FRFixResponse,
                        record.Summary, record.TicketNo, record.Company, record.Assignedto, record.Caller, record.CaseType,
                        record.CloseDate, record.DaysOpen, record.Department, record.InfrastructureType, record.Location,
                        record.Priority, record.IncidentType, record.Resolution, record.Shift);
                }
            }

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
