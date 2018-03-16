using System.Collections.Generic;

namespace InitAddon
{
    public class Coluna
    {
        protected Coluna(string nome, string descricao, bool obrigatoria = false)
        {
            Nome = nome;
            Descricao = descricao;
            Obrigatoria = obrigatoria;
        }

        public string Nome { get; set; }
        public string Descricao { get; set; }
        public bool Obrigatoria { get; set; } = false;
        public string ValorPadrao { get; set; } = "";
        public List<ValorValido> ValoresValidos { get; set; } = new List<ValorValido>() { };
        public int Tamanho { get; set; } = -1;
    }
}