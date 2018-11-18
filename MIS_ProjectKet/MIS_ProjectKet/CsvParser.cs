using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace MIS_ProjectKet
{
    class CsvParser
    {
        public CsvParser()
        {
            

            //use of library StreamReader to get csv file
            using ( var sr = new StreamReader(@"file path"))
            using (DataTable dt = new DataTable("test"))
            {
                //variable to be able to read csv file
                var reader = new CsvReader(sr);
                
                //csv configurations
                reader.Configuration.PrepareHeaderForMatch = header => Regex.Replace(header, @"\s", string.Empty);
                reader.Configuration.RegisterClassMap<DataRecordMap>();


                //CSVReader will now read the whole file into an enumerable
                IEnumerable<DataRecord> records = reader.GetRecords<DataRecord>();

                dt.Columns.Add("Create Date", typeof(string));
                dt.Columns.Add("IR Initial Response", typeof(string));
                dt.Columns.Add("FR Fix Response", typeof(string));
                dt.Columns.Add("Summary", typeof(string));
                dt.Columns.Add("Ticket #", typeof(string));
                dt.Columns.Add("Company", typeof(string));
                dt.Columns.Add("Assigned to", typeof(string));
                dt.Columns.Add("Caller", typeof(string));
                dt.Columns.Add("Case Type", typeof(string));
                dt.Columns.Add("Close Date", typeof(string));
                dt.Columns.Add("Days Open", typeof(string));
                dt.Columns.Add("Department", typeof(string));
                dt.Columns.Add("Infrastructure Type", typeof(string));
                dt.Columns.Add("Location", typeof(string));
                dt.Columns.Add("Priority", typeof(string));
                dt.Columns.Add("IncidentType", typeof(string));
                dt.Columns.Add("Resolution", typeof(string));
                dt.Columns.Add("Status", typeof(string));
                dt.Columns.Add("Shift", typeof(string));

                foreach (DataRecord record in records)
                {
                    dt.Rows.Add(record.CreateDate.Replace(" ","").Replace("\"",""), record.IRInitialResponse.Replace(" ", "").Replace("\"", "").Replace("at","@"), record.FRFixResponse.Replace(" ", "").Replace("\"", "").Replace("at", "@"), record.Summary,
                        record.TicketNo, record.Company, record.Assignedto, record.Caller, record.CaseType,
                        record.CloseDate, record.DaysOpen, record.Department, record.InfrastructureType, record.Location,
                        record.Priority, record.IncidentType, record.Resolution, record.Status, record.Shift);
                }

                WriteCsv wcsv = new WriteCsv(dt);
                sr.Close();
                
            }
            Console.WriteLine("FINISHED");
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
