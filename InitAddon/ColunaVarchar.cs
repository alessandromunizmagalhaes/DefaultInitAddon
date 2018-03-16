using System.Collections.Generic;

namespace InitAddon
{
    public class ColunaVarchar : Coluna
    {
        public ColunaVarchar(string nome, string descricao, int tamanho = 0, bool obrigatorio = false, string valor_padrao = "", List<ValorValido> valorValido = null) : base(nome, descricao, obrigatorio)
        {
            Tamanho = tamanho;
            ValorPadrao = ValorPadrao;
            if (valorValido != null)
            {
                ValoresValidos = valorValido;
            }
        }
    }
}
