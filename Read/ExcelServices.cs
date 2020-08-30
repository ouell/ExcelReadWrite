using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using ExcelReadWrite.Application;
using ExcelReadWrite.Domain;

namespace ExcelReadWrite.Read
{
    public class ExcelServices
    {
        private static readonly string[] FormataDataValidoExcel =
        {
            "dd/M/yyyy", "dd/M/yyyy", "d/MM/yyyy", "d/M/yyyy"
        };

        private const string FormatoDataExcel = "dd/MM/yyyy";

        private const string PathControle = @"C:\Temp\controle_pendencias.xlsx";

        public static void CreateControle(Tuple<AcheNota, AcheRetorno> ache,
                                          Tuple<BiolabNota, BiolabRetorno> biolab,
                                          Tuple<SandozNota, SandozRetorno> sandoz,
                                          Tuple<SanofiNota, SanofiRetorno> sanofi,
                                          Tuple<BoehringerNota, BoehringerRetorno> boehringer,
                                          Tuple<AstrazenecaNota, AstrazenecaRetorno> astrazeneca,
                                          Tuple<HyperaPharmaNota, HyperaPharmaRetorno> hyperaPharma,
                                          Tuple<HyperaRunningNota, HyperaRunningRetorno> hyperaRunning)
        {
            var workBook = new XLWorkbook(PathControle, XLEventTracking.Disabled);
            var sheet = workBook.Worksheets.Add($"{DateTime.Now.Day}-{DateTime.Now.Month.ToString().PadLeft(2, '0')}");

            sheet.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);


            sheet.Cell("A1").Value = $"{DateTime.Now:dd/MM/yyyy}";
            var range = sheet.Range("A1:A9");
            range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            range.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            range.Style.Alignment.SetTextRotation(90);
            range.Style.Font.Bold = true;
            range.Merge();

            sheet.Cell("B1").Value = "Laboratório";
            sheet.Cell("B1").Style.Font.Bold = true;
            sheet.Cell("C1").Value = "VAN";
            sheet.Cell("C1").Style.Font.Bold = true;
            sheet.Cell("D1").Value = "Nº Pendências Retornos Recebidos";
            sheet.Cell("D1").Style.Font.Bold = true;
            sheet.Cell("E1").Value = "Data Pendência Retorno Mais Antiga";
            sheet.Cell("E1").Style.Font.Bold = true;
            sheet.Cell("F1").Value = "Nº Pendências Notas Recebidas";
            sheet.Cell("F1").Style.Font.Bold = true;
            sheet.Cell("G1").Value = "Data Pendência Nota Mais Antiga";
            sheet.Cell("G1").Style.Font.Bold = true;
            sheet.Cell("H1").Value = "Nº Retorno Pendentes";
            sheet.Cell("H1").Style.Font.Bold = true;
            sheet.Cell("I1").Value = "Nº Notas Pendentes";
            sheet.Cell("I1").Style.Font.Bold = true;

            CreateAche(sheet, ache);
            CreateAstrazeneca(sheet, astrazeneca);
            CreateBiolab(sheet, biolab);
            CreateBoehringer(sheet, boehringer);
            CreateHyperaPharma(sheet, hyperaPharma);
            CreateSandoz(sheet, sandoz);
            CreateSanofi(sheet, sanofi);
            CreateHyperaRunning(sheet, hyperaRunning);

            workBook.Save();
        }

        private static void CreateAche(IXLWorksheet sheet,
                                       Tuple<AcheNota, AcheRetorno> ache)
        {
            sheet.Cell("B2").Value = "Aché";
            sheet.Cell("B2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C2").Value = "Visão";
            sheet.Cell("C2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D2").Value = ache.Item2.QuantidadePendencias;
            sheet.Cell("D2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E2").Value = ache.Item2.DataPrimeiraPendencia;
            sheet.Cell("E2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F2").Value = ache.Item1.QuantidadePendencias;
            sheet.Cell("F2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G2").Value = ache.Item1.DataPrimeiraPendencia;
            sheet.Cell("G2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H2").Value = "-";
            sheet.Cell("H2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I2").Value = "-";
            sheet.Cell("I2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateAstrazeneca(IXLWorksheet sheet,
                                              Tuple<AstrazenecaNota, AstrazenecaRetorno> astrazeneca)
        {
            sheet.Cell("B3").Value = "Astrazeneca";
            sheet.Cell("B3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C3").Value = "Fidelize";
            sheet.Cell("C3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D3").Value = astrazeneca.Item2.QuantidadePendencias;
            sheet.Cell("D3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E3").Value = astrazeneca.Item2.DataPrimeiraPendencia;
            sheet.Cell("E3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F3").Value = astrazeneca.Item1.QuantidadePendencias;
            sheet.Cell("F3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G3").Value = astrazeneca.Item1.DataPrimeiraPendencia;
            sheet.Cell("G3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H3").Value = "-";
            sheet.Cell("H3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I3").Value = "-";
            sheet.Cell("I3").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateBiolab(IXLWorksheet sheet,
                                         Tuple<BiolabNota, BiolabRetorno> biolab)
        {
            sheet.Cell("B4").Value = "Biolab";
            sheet.Cell("B4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C4").Value = "Pharmalink";
            sheet.Cell("C4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D4").Value = biolab.Item2.QuantidadePendencias;
            sheet.Cell("D4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E4").Value = biolab.Item2.DataPrimeiraPendencia;
            sheet.Cell("E4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F4").Value = biolab.Item1.QuantidadePendencias;
            sheet.Cell("F4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G4").Value = biolab.Item1.DataPrimeiraPendencia;
            sheet.Cell("G4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H4").Value = "-";
            sheet.Cell("H4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I4").Value = "-";
            sheet.Cell("I4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateBoehringer(IXLWorksheet sheet,
                                             Tuple<BoehringerNota, BoehringerRetorno> boehringer)
        {
            sheet.Cell("B5").Value = "Boehringe";
            sheet.Cell("B5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C5").Value = "Fidelize";
            sheet.Cell("C5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D5").Value = boehringer.Item2.QuantidadePendencias;
            sheet.Cell("D5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E5").Value = boehringer.Item2.DataPrimeiraPendencia;
            sheet.Cell("E5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F5").Value = boehringer.Item1.QuantidadePendencias;
            sheet.Cell("F5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G5").Value = boehringer.Item1.DataPrimeiraPendencia;
            sheet.Cell("G5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H5").Value = "-";
            sheet.Cell("H5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I5").Value = "-";
            sheet.Cell("I5").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateHyperaPharma(IXLWorksheet sheet,
                                               Tuple<HyperaPharmaNota, HyperaPharmaRetorno> hyperaPharma)
        {
            sheet.Cell("B6").Value = "Hypera";
            sheet.Cell("B6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C6").Value = "Pharmalink";
            sheet.Cell("C6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D6").Value = hyperaPharma.Item2.QuantidadePendencias;
            sheet.Cell("D6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E6").Value = hyperaPharma.Item2.DataPrimeiraPendencia;
            sheet.Cell("E6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F6").Value = hyperaPharma.Item1.QuantidadePendencias;
            sheet.Cell("F6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G6").Value = hyperaPharma.Item1.DataPrimeiraPendencia;
            sheet.Cell("G6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H6").Value = "-";
            sheet.Cell("H6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I6").Value = "-";
            sheet.Cell("I6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateHyperaRunning(IXLWorksheet sheet,
                                                Tuple<HyperaRunningNota, HyperaRunningRetorno> hyperaRunning)
        {
            sheet.Cell("B6").Value = "Hypera";
            sheet.Cell("B6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C6").Value = "Running";
            sheet.Cell("C6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D6").Value = hyperaRunning.Item2.QuantidadePendencias;
            sheet.Cell("D6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E6").Value = hyperaRunning.Item2.DataPrimeiraPendencia;
            sheet.Cell("E6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F6").Value = hyperaRunning.Item1.QuantidadePendencias;
            sheet.Cell("F6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G6").Value = hyperaRunning.Item1.DataPrimeiraPendencia;
            sheet.Cell("G6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H6").Value = "-";
            sheet.Cell("H6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I6").Value = "-";
            sheet.Cell("I6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateSandoz(IXLWorksheet sheet,
                                         Tuple<SandozNota, SandozRetorno> sandoz)
        {
            sheet.Cell("B7").Value = "Sandoz";
            sheet.Cell("B7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C7").Value = "Pharmalink";
            sheet.Cell("C7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D7").Value = sandoz.Item2.QuantidadePendencias;
            sheet.Cell("D7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E7").Value = sandoz.Item2.DataPrimeiraPendencia;
            sheet.Cell("E7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F7").Value = sandoz.Item1.QuantidadePendencias;
            sheet.Cell("F7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G7").Value = sandoz.Item1.DataPrimeiraPendencia;
            sheet.Cell("G7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H7").Value = "-";
            sheet.Cell("H7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I7").Value = "-";
            sheet.Cell("I7").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        private static void CreateSanofi(IXLWorksheet sheet,
                                         Tuple<SanofiNota, SanofiRetorno> sanofi)
        {
            sheet.Cell("B8").Value = "Sanofi";
            sheet.Cell("B8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("C8").Value = "Fidelize";
            sheet.Cell("C8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            sheet.Cell("D8").Value = sanofi.Item2.QuantidadePendencias;
            sheet.Cell("D8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("E8").Value = sanofi.Item2.DataPrimeiraPendencia;
            sheet.Cell("E8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("F8").Value = sanofi.Item1.QuantidadePendencias;
            sheet.Cell("F8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("G8").Value = sanofi.Item1.DataPrimeiraPendencia;
            sheet.Cell("G8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("H8").Value = "-";
            sheet.Cell("H8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            sheet.Cell("I8").Value = "-";
            sheet.Cell("I8").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
        }

        public static BaseDomain ReadData(string path,
                                          string columnLetter,
                                          string sheetName = null)
        {
            var xlsFile = ReadExcelFile(path, sheetName);

            var maxRows = xlsFile.RowsUsed().Count();
            var maxColumns = xlsFile.ColumnsUsed().Count();

            var listDate = new List<DateTime>();
            for (var row = 2;
                 row <= maxRows;
                 row++)
            {
                for (var column = 1;
                     column <= maxColumns;
                     column++)
                {
                    var value = xlsFile.Cell(row, column);
                    if (value.Address?.ColumnLetter == columnLetter)
                        listDate.Add(FormatDateCell(value));
                }
            }

            var baseDomain = new BaseDomain
            {
                QuantidadePendencias = (maxRows - 1).ToString(),
                DataPrimeiraPendencia = listDate.Min().Date.ToString(Execute.CulturaPtBr)
            };

            return baseDomain;
        }

        public static List<string> CreateLines(string path)
        {
            var xlsFile = ReadExcelFile(path);
            
            var maxColumns = xlsFile.ColumnsUsed().Count();
            var maxRows = xlsFile.RowsUsed().Count();
            var listRetorno = new List<string>();
            for (var row = 2; row <= maxRows; row++)
            {
                var text = new StringBuilderWrapper();
                for (var column = 1; column <= maxColumns; column++)
                {
                    text.Append(xlsFile.Cell(row, column).GetString().Trim());
                    text.AppendSplitter();
                }

                text.AppendBreakLine();
                listRetorno.Add(text.ToString());
            }

            return listRetorno;
        }

        private static IXLWorksheet ReadExcelFile(string excelPath,
                                                  string sheetName = null)
        {
            var workBook = new XLWorkbook(excelPath, XLEventTracking.Disabled);
            var workSheet = string.IsNullOrEmpty(sheetName) ? workBook.Worksheets.First() : workBook.Worksheet(sheetName);

            if (workSheet == null)
                throw new ArgumentException("Planilha não existe");

            workSheet.Columns().CellsUsed();

            return workSheet;
        }

        private static DateTime FormatDateCell(IXLCell cell)
        {
            cell.SetDataType(XLDataType.DateTime);
            var value = cell.GetDateTime().ToString(FormatoDataExcel, Execute.CulturaPtBr);

            return DateTime.Parse(value);
        }
    }
}