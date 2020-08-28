using System.Text;

namespace ExcelReadWrite.Read
{
    public class StringBuilderWrapper
    {
        /// <summary>
        /// String builder
        /// </summary>
        private readonly StringBuilder _builder;

        /// <summary>
        /// Construtor
        /// </summary>
        public StringBuilderWrapper()
        {
            _builder = new StringBuilder();
        }

        /// <summary>
        /// Método responsável por adicionar um texto ao string builder
        /// </summary>
        /// <param name="text">Texto que ser adicionado ao string builder</param>
        public StringBuilderWrapper Append(object text)
        {
            if (text == null)
                return this;
            _builder.Append(text.ToString()?.Replace(";", ""));
            return this;
        }

        /// <summary>
        /// Adiciona um quebra de linha ao string builder
        /// </summary>
        public void AppendBreakLine()
        {
            _builder.Append("\r\n");
        }

        /// <summary>
        /// Retorna a string
        /// </summary>
        public override string ToString()
        {
            return _builder.ToString();
        }

        /// <summary>
        /// Limpa o string builder
        /// </summary>
        public void Clear()
        {
            _builder.Length = 0;
        }

        /// <summary>
        /// Adiciona um separador
        /// </summary>
        public void AppendSplitter()
        {
            _builder.Append(";");
        }
    }
}