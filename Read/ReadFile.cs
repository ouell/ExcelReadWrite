using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ExcelReadWrite.Application;
using ExcelReadWrite.Domain;

namespace ExcelReadWrite.Read
{
    public class ReadFile
    {
        private static readonly string[] FormataDataValidoExcel = {
            "dd/M/yyyy", "dd/M/yyyy","d/MM/yyyy","d/M/yyyy"
        };
        private const string FormatoDataExcel = "dd/MM/yyyy";

        public static IXLWorksheet ReadExcelFile(string excelPath)
        {
            var workBook = new XLWorkbook(excelPath, XLEventTracking.Disabled);
                var workSheet = workBook.Worksheets.First();
                if(workSheet == null)
                    throw new ArgumentException("Planilha não existe");

                workSheet.Columns().CellsUsed();

                return workSheet;
        }

        private static string CreateLines(IXLWorksheet workSheet)
        {
            var text = new StringBuilderWrapper();
            var maxColumns = workSheet.ColumnsUsed().Count();
            var maxRows = workSheet.RowsUsed().Count();

            for (var row = 2; row <= maxRows; row++)
            {
                for (var column = 1; column <= maxColumns; column++)
                {
                    text.Append(workSheet.Cell(row, column).GetString().Trim());
                    text.AppendSplitter();
                }
                text.AppendBreakLine();
            }

            return text.ToString();
        }

        public static DateTime FormatDateCell(IXLCell cell)
        {
            cell.SetDataType(XLDataType.DateTime);
            var value = cell.GetDateTime().ToString(FormatoDataExcel, Execute.CulturaPtBr);
            
            return DateTime.Parse(value);
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
                DataPrimeiraPendencia = listDate.Min().ToString(CultureInfo.InvariantCulture)
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