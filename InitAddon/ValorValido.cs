namespace InitAddon
{
    public class ValorValido
    {
        public string Valor { get; set; }
        public string Descricao { get; set; }

        public ValorValido(string valor, string descricao)
        {
            Valor = valor;
            Descricao = descricao;
        }
    }
}