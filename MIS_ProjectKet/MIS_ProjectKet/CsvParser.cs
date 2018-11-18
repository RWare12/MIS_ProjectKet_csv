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
            using ( var sr = new StreamReader(@"C:\Users\suare\Desktop\Project\test.csv"))
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
                    dt.Rows.Add(record.CreateDate.Replace(" ",""), record.IRInitialResponse.Replace(" ", "").Replace("at","@"), record.FRFixResponse.Replace(" ", "").Replace("at", "@"), record.Summary,
                        record.TicketNo, record.Company, record.Assignedto, record.Caller, record.CaseType,
                        record.CloseDate, record.DaysOpen, record.Department, record.InfrastructureType, record.Location,
                        record.Priority, record.IncidentType, record.Resolution, record.Status, record.Shift);
                }

                int row = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    Console.WriteLine("=====================Row " + row + "========================");
                    Console.WriteLine("Create Date: {0}, \nIRInitialResponse: {1}, \nFR Fix Response: {2}, "
                        + " \nSummary: {3}, \nTicket #: {4}, \nCompany: {5}, \nAssigned to: {6}, \nCaller: {7}, "
                        + " \nCase Type: {8}, \nClose Date: {9}, \nDays Open: {10}, \nDepartment: {11}, \nInfrastructure Type: {12},"
                        + " \nLocation: {13}, \nPriority: {14}, \nIncident Type: {15}, \nResolution: {16}, \nStatus: {17} "
                        + " \nShift: {18}", dr["Create Date"], dr["IR Initial Response"], dr["FR Fix Response"], dr["Summary"],
                        dr["Ticket #"], dr["Company"], dr["Assigned to"], dr["Caller"], dr["Case Type"],
                        dr["Close Date"], dr["Days Open"], dr["Department"], dr["Infrastructure Type"], dr["Location"],
                        dr["Priority"], dr["IncidentType"], dr["Resolution"], dr["Status"], dr["Shift"]);
                    //Console.WriteLine(dr[18]);
                    row++;
                }


                WriteCsv wcsv = new WriteCsv(dt);
                Console.WriteLine(row);
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
