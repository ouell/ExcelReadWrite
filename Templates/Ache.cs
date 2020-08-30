using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class Ache
    {
        private const string DateColumnLetter = "A";
        
        public static Tuple<AcheNota, AcheRetorno> ReadAche(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, AcheNota>();
                cfg.CreateMap<BaseDomain, AcheRetorno>();
            });
            var mapper = config.CreateMapper();      
            
            var files = Directory.GetFiles(path).ToList();

            var acheNota = new AcheNota();
            var acheRetorno = new AcheRetorno();
            files.ForEach(file =>
            {
                if (file.Contains("nota.xlsx"))
                {
                    acheNota = mapper.Map<AcheNota>(ExcelServices.ReadData(file, DateColumnLetter));
                }

                if (file.Contains("retorno.xlsx"))
                {
                    acheRetorno = mapper.Map<AcheRetorno>(ExcelServices.ReadData(file, DateColumnLetter));
                }
            });

            return new Tuple<AcheNota, AcheRetorno>(acheNota, acheRetorno);
        }
    }
}