using System;
using System.Diagnostics;

namespace MIS_ProjectKet
{
    class Program
    {
        static void Main(string[] args)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            CsvParser run = new CsvParser();
            sw.Stop();
            Console.WriteLine("Elapsed={0}", sw.Elapsed);
            Console.ReadKey();
        }

        

    }
}
