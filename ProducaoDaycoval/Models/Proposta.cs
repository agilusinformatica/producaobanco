using System;

namespace ProducaoDaycoval.Models
{
    public class Proposta
    {
        public string Filial { get; set; }
        public string Gerente { get; set; }
        public string Promotora { get; set; }
        public string Empregador { get; set; }
        public string Orgao { get; set; }
        public string Regra { get; set; }
        public string Situacao { get; set; }
        public string Agente { get; set; }

        public string NumeroProposta { get; set; }
        public string NumeroContrato { get; set; }
        public string Cliente { get; set; }
        public string Cpf { get; set; }
        public string Matricula { get; set; }
        public DateTime DataCadastro { get; set; }
        public DateTime DataBase { get; set; }
        public DateTime? DataUltimoVencimento { get; set; }
        public int Prazo { get; set; }
        public int QtdeParcelas { get; set; }
        public Decimal Taxa { get; set; }
        public Decimal ValorLiquido { get; set; }
        public Decimal Iof { get; set; }
        public Decimal ValorOperacao { get; set; }
        public Decimal ValorFinanciado { get; set; }
        public Decimal? ValorParcela { get; set; }
        public Decimal? ValorCreditado { get; set; }
    }
}