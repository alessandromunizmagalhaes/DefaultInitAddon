using System.Collections.Generic;

namespace InitAddon
{
    public class Coluna
    {
        public Coluna(string nome, string descricao, bool obrigatorio = false)
        {
            Nome = nome;
            Descricao = descricao;
            Obrigatorio = obrigatorio;
        }

        public string Nome { get; set; }
        public string Descricao { get; set; }
        public bool Obrigatorio { get; set; } = false;
        public string ValorPadrao { get; set; } = "";
        public List<ValorValido> ValoresValidos { get; set; } = new List<ValorValido>() { };
        public int Tamanho { get; set; } = -1;
    }
}