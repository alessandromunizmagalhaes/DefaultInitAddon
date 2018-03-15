namespace InitAddon
{
    public class Coluna
    {
        public Coluna(string nome, string descricao, ColunaTipo tipo, int tamanho = 0)
        {
            if (CampoTamanhoObrigatorio(tipo) && tamanho <= 0)
                throw new CustomException($"Erro ao tentar adicionar Coluna. Informe o tamanho do campo {nome}");

            Nome = nome;
            Descricao = descricao;
            _tamanho = tamanho;
            Tipo = tipo;
        }

        public string Nome { get; set; }
        public string Descricao { get; set; }

        public int _tamanho;
        public int Tamanho
        {
            get
            {
                return CampoTamanhoObrigatorio(Tipo) ? _tamanho : -1;
            }
        }
        public ColunaTipo Tipo { get; set; }


        private bool CampoTamanhoObrigatorio(ColunaTipo tipo)
        {
            // somente campo varchar e campo int/numerico que se seta o tamanho
            return (tipo == ColunaTipo.Varchar || tipo == ColunaTipo.Int);
        }
    }

    public enum ColunaTipo
    {
        Varchar = 1,
        Text = 2,
        Date = 3,
        Time = 4,
        Int = 5,
        Price = 6,
        Percent = 7,
        Quantity = 8,
        Sum = 9,
        Link = 10,
        Image = 11,
    }
}