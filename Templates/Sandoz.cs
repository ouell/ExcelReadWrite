using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ExcelReadWrite.Application;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public static class Sandoz
    {
        public static Tuple<SandozNota, SandozRetorno> ReadSandoz(string path)
        {
            if(string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var files = Directory.GetFiles(path).ToList();

            SandozNota sandozNota = null;
            SandozRetorno sandozRetorno = null;
            files.ForEach(file =>
            {
                if (file.Contains("nota.xlsx"))
                {
                    sandozRetorno = ReadRetorno(file);
                }
                if (file.Contains("retorno.xlsx"))
                {
                    sandozNota = ReadNota(file);
                }
            });

            
            return new Tuple<SandozNota, SandozRetorno>(sandozNota, sandozRetorno);
        }

        private static SandozRetorno ReadRetorno(string path)
        {
            var xlsFile = ReadFile.ReadExcelFile(path);
            
            var maxRows = xlsFile.RowsUsed().Count();
            var maxColumns = xlsFile.ColumnsUsed().Count();
            
            var listDate = new List<DateTime>();
            for (var row = 2; row <= maxRows; row++)
            {
                for (var column = 1; column <= maxColumns; column++)
                {
                    var value = xlsFile.Cell(row, column);
                    switch (value.Address?.ColumnLetter)
                    {
                        case "B":
                            listDate.Add(ReadFile.FormatDateCell(value));
                            break;
                    }
                }
            }
            
            var sandozRetorno = new SandozRetorno
            {
                QuantidadePendencias = maxRows.ToString(),
                DataPrimeiraPendencia = listDate.Min().ToString(Execute.CulturaPtBr)
            };

            return sandozRetorno;
        }
        
        private static SandozNota ReadNota(string path)
        {
            var xlsFile = ReadFile.ReadExcelFile(path);
            
            var maxRows = xlsFile.RowsUsed().Count();
            var maxColumns = xlsFile.ColumnsUsed().Count();
            
            var listDate = new List<DateTime>();
            for (var row = 2; row <= maxRows; row++)
            {
                for (var column = 1; column <= maxColumns; column++)
                {
                    var value = xlsFile.Cell(row, column);
                    switch (value.Address?.ColumnLetter)
                    {
                        case "B":
                            listDate.Add(ReadFile.FormatDateCell(value));
                            break;
                    }
                }
            }
            
            var sandozNota = new SandozNota
            {
                QuantidadePendencias = maxRows.ToString(),
                DataPrimeiraPendencia = listDate.Min().ToString(CultureInfo.InvariantCulture)
            };

            return sandozNota;
        }
    }
}