using System.Collections.Generic;

namespace InitAddon
{
    public class ColunaVarchar : Coluna
    {
        public ColunaVarchar(string nome, string descricao, bool obrigatorio = false, string valor_padrao = "", int tamanho = 0, List<ValorValido> valorValido = null) : base(nome, descricao, obrigatorio)
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
