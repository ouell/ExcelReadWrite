using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Templates;

namespace ExcelReadWrite.Application
{
    public static class Execute
    {
        private const string Path = @"C:\Temp\";
        public static readonly CultureInfo CulturaPtBr = CultureInfo.GetCultureInfo("pt-BR");
        public static async Task RunTask()
        {
             var folders = Directory.GetDirectories(Path);

             if(folders.Length != 8)
                 throw new ArgumentException("Invalid number of folders, 8 folders are required");

             Tuple<SandozNota, SandozRetorno> sandoz;
             
             folders.ToList().ForEach(folder =>
             {
                 if(folder.Contains("sandoz"))
                 {
                     sandoz = Sandoz.ReadSandoz(folder);
                 }
             });
             // Parallel.ForEach(folders,
             //                  folder =>
             //                  {
             //                    if(folder.Contains("sandoz"))
             //                    {
             //                        sandoz = Sandoz.ReadSandoz(folder);
             //                    }
             //                  });
        }
    }
}