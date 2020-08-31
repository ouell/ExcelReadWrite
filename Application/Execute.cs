using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using AutoMapper;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;
using ExcelReadWrite.Templates;

namespace ExcelReadWrite.Application
{
    public static class Execute
    {
        private const string Path = @"C:\Temp\";
        public static readonly CultureInfo CulturaPtBr = CultureInfo.GetCultureInfo("pt-BR");
        public static async Task RunTask()
        {
            var init = DateTime.Now;
            var folders = Directory.GetDirectories(Path);

             if(folders.Length != 8)
                 throw new ArgumentException("Invalid number of folders, 8 folders are required");

             Tuple<AcheNota, AcheRetorno> ache = null;
             Tuple<BiolabNota, BiolabRetorno> biolab = null;
             Tuple<SandozNota, SandozRetorno> sandoz = null;
             Tuple<SanofiNota, SanofiRetorno> sanofi = null;
             Tuple<BoehringerNota, BoehringerRetorno> boehringer = null;
             Tuple<AstrazenecaNota, AstrazenecaRetorno> astrazeneca = null;
             Tuple<HyperaPharmaNota, HyperaPharmaRetorno> hyperaPharma = null;
             Tuple<HyperaRunningNota, HyperaRunningRetorno> hyperaRunning = null;
             foreach (var folder in folders)
             {
                 if(folder.Contains("sandoz"))
                 {
                     sandoz = Sandoz.ReadSandoz(folder);
                     continue;
                 }

                 if (folder.Contains("ache"))
                 {
                     ache = Ache.ReadAche(folder);
                     continue;
                 }

                 if (folder.Contains("biolab"))
                 {
                     biolab = Biolab.ReadBiolab(folder);
                     continue;
                 }
                 
                 if (folder.Contains("astrazeneca"))
                 {
                     astrazeneca = Astrazeneca.ReadAztrazeneca(folder);
                     continue;
                 }

                 if (folder.Contains("boehringer"))
                 {
                     boehringer = Boehringer.ReadBoehringer(folder);
                     continue;
                 }

                 if (folder.Contains("hypera_plk"))
                 {
                     hyperaPharma = HyperaPharma.ReadHyperaPharma(folder);
                     continue;;
                 }

                 if (folder.Contains("hypera_rng"))
                 {
                     hyperaRunning = HyperaRunning.ReadHyperaRunning(folder);
                     continue;
                 }

                 if (folder.Contains("sanofi"))
                 {
                     sanofi = Sanofi.ReadSanofi(folder);
                 }
             };
             
             ExcelServices.CreateControle(ache, biolab, sandoz, sanofi, boehringer, astrazeneca, hyperaPharma, hyperaRunning);
             
             var end = DateTime.Now;
             Console.WriteLine($"Inicio: {init}");
             Console.WriteLine($"Fim: {end}");
             Console.WriteLine($"Tempo: {end - init}");
        }
    }
}