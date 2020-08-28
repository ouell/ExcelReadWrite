using System;
using System.IO;
using System.Linq;
using ExcelReadWrite.Read;
using ExcelReadWrite.Templates.Enum;
using static ExcelReadWrite.Application.Execute;

namespace ExcelReadWrite
{
    internal static class Program
    {
        private static void Main()
        {
            //var value = ReadFile.ReadExcelFile(TemplateType.RetornoSandoz, @"C:\Temp\sandoz.xlsx");
            //Console.WriteLine(value);
            // var folders = Directory.GetDirectories(@"C:\Temp\");
            // folders.ToList().ForEach(Console.WriteLine);

            RunTask();
        }
    }
}