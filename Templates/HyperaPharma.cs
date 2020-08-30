using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class HyperaPharma
    {
        private const string DateColumnLetter = "D";

        public static Tuple<HyperaPharmaNota, HyperaPharmaRetorno> ReadHyperaPharma(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, HyperaPharmaNota>();
                cfg.CreateMap<BaseDomain, HyperaPharmaRetorno>();
            });
            var mapper = config.CreateMapper();

            var file = Directory.GetFiles(path).FirstOrDefault();

            var hyperaPharmaNota = mapper.Map<HyperaPharmaNota>(ExcelServices.ReadData(file, DateColumnLetter, "Nota"));
            var hyperaPharmaRetorno = mapper.Map<HyperaPharmaRetorno>(ExcelServices.ReadData(file, DateColumnLetter, "Retorno"));

            return new Tuple<HyperaPharmaNota, HyperaPharmaRetorno>(hyperaPharmaNota, hyperaPharmaRetorno);
        }
    }
}