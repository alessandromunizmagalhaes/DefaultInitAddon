using SAPbobsCOM;
using System.Collections.Generic;

namespace InitAddon
{
    public class Tabela
    {
        private string _nome;

        public Tabela(string nome, string descricao, BoUTBTableType tipo, List<Coluna> colunas)
        {
            _nome = nome;
            Descricao = descricao;
            Tipo = tipo;
            Colunas = colunas;
        }

        public string NomeComArroba { get { return _nome[0] == '@' ? _nome : "@" + _nome; } }
        public string NomeSemArroba { get { return _nome[0] == '@' ? _nome.Remove(0, 1) : _nome; } }
        public string Descricao { get; set; }
        public SAPbobsCOM.BoUTBTableType Tipo { get; set; }
        public List<Coluna> Colunas { get; set; }
    }
}