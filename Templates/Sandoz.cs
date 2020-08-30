using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public static class Sandoz
    {
        private const string ColumnLetter = "B";

        public static Tuple<SandozNota, SandozRetorno> ReadSandoz(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var files = Directory.GetFiles(path).ToList();

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, SandozNota>();
                cfg.CreateMap<BaseDomain, SandozRetorno>();
            });
            var mapper = config.CreateMapper();
            
            SandozNota sandozNota = null;
            SandozRetorno sandozRetorno = null;
            files.ForEach(file =>
            {
                if (file.Contains("nota.xlsx"))
                {
                    sandozNota = mapper.Map<SandozNota>(ExcelServices.ReadData(file, ColumnLetter));
                }

                if (file.Contains("retorno.xlsx"))
                {
                    sandozRetorno = mapper.Map<SandozRetorno>(ExcelServices.ReadData(file, ColumnLetter));
                }
            });

            return new Tuple<SandozNota, SandozRetorno>(sandozNota, sandozRetorno);
        }
    }
}