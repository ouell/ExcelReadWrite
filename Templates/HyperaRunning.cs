using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelReadWrite.Application;
using ExcelReadWrite.Domain;
using ExcelReadWrite.Read;

namespace ExcelReadWrite.Templates
{
    public class HyperaRunning
    {
        public static Tuple<HyperaRunningNota, HyperaRunningRetorno> ReadHyperaRunning(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentException("Invalid path.");

            var hyperaRunningNota = new HyperaRunningNota();
            var hyperaRunningRetorno = new HyperaRunningRetorno();

            var file = Directory.GetFiles(path).ToList().FirstOrDefault();
            var listData = new List<string>();
            listData.AddRange(ExcelServices.CreateLines(file));

            var listaDataRetorno = new List<DateTime>();
            var listaDataNota = new List<DateTime>();

            foreach (var linha in listData)
            {
                var arrayDados = linha.Split(";");
                if (arrayDados[0].ToUpper().Contains("ESPELHO"))
                    listaDataNota.Add(DateTime.Parse(arrayDados[7]).Date);

                if (arrayDados[0].ToUpper().Contains("RETORNO"))
                    listaDataRetorno.Add(DateTime.Parse(arrayDados[7]).Date);
            }

            hyperaRunningNota.QuantidadePendencias = listaDataNota.Count.ToString();
            hyperaRunningRetorno.QuantidadePendencias = listaDataRetorno.Count.ToString();
            hyperaRunningNota.DataPrimeiraPendencia = listaDataNota.Min().Date.ToString(Execute.CulturaPtBr);
            hyperaRunningRetorno.DataPrimeiraPendencia = listaDataRetorno.Min().Date.ToString(Execute.CulturaPtBr);

            return new Tuple<HyperaRunningNota, HyperaRunningRetorno>(hyperaRunningNota, hyperaRunningRetorno);
        }
    }
}