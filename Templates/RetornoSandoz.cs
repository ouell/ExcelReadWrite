using System;

namespace ExcelReadWrite.Templates
{
    public class RetornoSandoz : AbstractTemplate
    {
        public string NumeroPedido { get; set; }
        public DateTime DataPedido { get; set; }
        public string CnpjDistribuidora { get; set; }
        public string NomeDistribuidora { get; set; }
        public string CnpjCliente { get; set; }
        public string NomeCliente { get; set; }
        public string NomenclaturaArquivo { get; set; }
    }
}