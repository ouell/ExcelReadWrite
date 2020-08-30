using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class Sanofi
    {
        private const string DateNotaColumnLetter = "F";
        private const string DateRetornoColumnLetter = "D";

        public static Tuple<SanofiNota, SanofiRetorno> ReadSanofi(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var files = Directory.GetFiles(path).ToList();

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, SanofiNota>();
                cfg.CreateMap<BaseDomain, SanofiRetorno>();
            });
            var mapper = config.CreateMapper();

            SanofiNota sanofiNota = null;
            SanofiRetorno sanofiRetorno = null;
            files.ForEach(file =>
            {
                if (file.Contains("nota.xlsx"))
                {
                    sanofiNota = mapper.Map<SanofiNota>(ExcelServices.ReadData(file, DateNotaColumnLetter));
                }

                if (file.Contains("retorno.xlsx"))
                {
                    sanofiRetorno = mapper.Map<SanofiRetorno>(ExcelServices.ReadData(file, DateRetornoColumnLetter));
                }
            });

            return new Tuple<SanofiNota, SanofiRetorno>(sanofiNota, sanofiRetorno);
        }
    }
}