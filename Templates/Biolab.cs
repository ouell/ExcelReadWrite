using System;
using System.IO;
using System.Linq;
using AutoMapper;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class Biolab
    {
        private const string DateColumnLetter = "D";
        
        public static Tuple<BiolabNota, BiolabRetorno> ReadBiolab(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var config = new MapperConfiguration(cfg =>
            {
                cfg.CreateMap<BaseDomain, BiolabNota>();
                cfg.CreateMap<BaseDomain, BiolabRetorno>();
            });
            var mapper = config.CreateMapper();      
            
            var file = Directory.GetFiles(path).FirstOrDefault();

            var biolabNota = mapper.Map<BiolabNota>(ExcelServices.ReadData(file, DateColumnLetter, "Nota"));
            var biolabRetorno = mapper.Map<BiolabRetorno>(ExcelServices.ReadData(file, DateColumnLetter, "Retorno"));

            return new Tuple<BiolabNota, BiolabRetorno>(biolabNota, biolabRetorno);
        }
    }
}