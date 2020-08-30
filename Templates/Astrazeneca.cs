using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class Astrazeneca
    {
        private const string DateNotaColumnLetter = "F";
        private const string DateRetornoColumnLetter = "D";

        public static Tuple<AstrazenecaNota, AstrazenecaRetorno> ReadAztrazeneca(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var files = Directory.GetFiles(path).ToList();

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, AstrazenecaNota>();
                cfg.CreateMap<BaseDomain, AstrazenecaRetorno>();
            });
            var mapper = config.CreateMapper();
            
            AstrazenecaNota astrazenecaNota = null;
            AstrazenecaRetorno astrazenecaRetorno = null;
            files.ForEach(file =>
            {
                if (file.Contains("nota.xlsx"))
                {
                    astrazenecaNota = mapper.Map<AstrazenecaNota>(ExcelServices.ReadData(file, DateNotaColumnLetter));
                }

                if (file.Contains("retorno.xlsx"))
                {
                    astrazenecaRetorno = mapper.Map<AstrazenecaRetorno>(ExcelServices.ReadData(file, DateRetornoColumnLetter));
                }
            });

            return new Tuple<AstrazenecaNota, AstrazenecaRetorno>(astrazenecaNota, astrazenecaRetorno);
        }
    }
}