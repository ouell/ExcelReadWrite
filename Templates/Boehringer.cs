using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class Boehringer
    {
        private const string DateNotaColumnLetter = "F";
        private const string DateRetornoColumnLetter = "D";

        public static Tuple<BoehringerNota, BoehringerRetorno> ReadBoehringer(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var files = Directory.GetFiles(path).ToList();

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, BoehringerNota>();
                cfg.CreateMap<BaseDomain, BoehringerRetorno>();
            });
            var mapper = config.CreateMapper();
            
            BoehringerNota boehringerNota = null;
            BoehringerRetorno boehringerRetorno = null;
            files.ForEach(file =>
            {
                if (file.Contains("nota.xlsx"))
                {
                    boehringerNota = mapper.Map<BoehringerNota>(ExcelServices.ReadData(file, DateNotaColumnLetter));
                }

                if (file.Contains("retorno.xlsx"))
                {
                    boehringerRetorno = mapper.Map<BoehringerRetorno>(ExcelServices.ReadData(file, DateRetornoColumnLetter));
                }
            });

            return new Tuple<BoehringerNota, BoehringerRetorno>(boehringerNota, boehringerRetorno);
        }
    }
}