using System;
using System.Collections.Generic;
using FlexCel.XlsAdapter;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Web.Script.Serialization;
using System.Globalization;
using ProducaoDaycoval.Models;

namespace ProducaoDaycoval.Controllers
{
    public class CetelemController : Controller
    {
        //
        // GET: /Cetelem/

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(Upload upload)
        {
            string nomeArquivo = @"e:\home\agilus\Temp\"+ Path.GetRandomFileName().Replace(".", "");

            upload.arquivo.SaveAs(nomeArquivo);
            XlsFile excel = new XlsFile(nomeArquivo);
            string resposta = "";

            if (Utils.TextoCelula(excel, "C6") == "RELATÓRIO DE PROPOSTAS CADASTRADAS")
            {
                resposta = SerializaProposta(excel);
            }

            System.IO.File.Delete(nomeArquivo);
            return this.Content(resposta, "text/xml");
        }

        private string SerializaProposta(XlsFile excel)
        {
            var propostas = new List<Proposta>();
            string filial = String.Empty, 
                gerente = String.Empty, 
                promotora = String.Empty, 
                empregador = String.Empty, 
                orgao = String.Empty, 
                regra = String.Empty, 
                situacao = String.Empty, 
                agente = String.Empty;

            int QtdeLinhas = excel.RowCount;
            for (int linha = 1; linha <= QtdeLinhas; linha++)
            {
                string CelulaA = Utils.TextoCelula(excel, "A", linha);

                if (CelulaA.StartsWith("FILIAL:"))
                    filial = CelulaA.Substring(8);
                else if (CelulaA.StartsWith("EMPREGADOR:"))
                    empregador = CelulaA.Substring(12);
                else if (CelulaA.StartsWith("ORGAO:"))
                    orgao = CelulaA.Substring(7);
                else if (CelulaA.StartsWith("REGRA:"))
                    regra = CelulaA.Substring(7);
                else if (CelulaA.StartsWith("SITUAÇÃO:"))
                    situacao = CelulaA.Substring(10);
                else if (CelulaA.StartsWith("OPERADOR:"))
                    agente = CelulaA.Substring(10);
                else if ((CelulaA == Utils.SoNumeros(CelulaA)) && (CelulaA != String.Empty))
                {
                    var proposta = new Proposta();
                    proposta.Filial = filial;
                    proposta.Empregador = empregador;
                    proposta.Orgao = orgao;
                    proposta.Regra = regra;
                    proposta.Situacao = situacao;
                    proposta.Agente = agente;

                    proposta.NumeroProposta = Utils.TextoCelula(excel, "A", linha);
                    proposta.NumeroContrato = Utils.TextoCelula(excel, "D", linha);
                    proposta.Cliente = Utils.TextoCelula(excel, "E", linha);
                    proposta.Cpf = Utils.TextoCelula(excel, "F", linha);
                    proposta.Matricula = Utils.TextoCelula(excel, "G", linha);
                    proposta.DataCadastro = DateTime.ParseExact(Utils.TextoCelula(excel, "H", linha), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    proposta.DataBase = DateTime.ParseExact(Utils.TextoCelula(excel, "J", linha), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    proposta.QtdeParcelas = Convert.ToInt32(Utils.TextoCelula(excel, "O", linha));
                    proposta.Taxa = Convert.ToDecimal(Utils.TextoCelula(excel, "P", linha));
                    proposta.ValorLiquido = Convert.ToDecimal(Utils.TextoCelula(excel, "Q", linha));
                    proposta.Iof = Convert.ToDecimal(Utils.TextoCelula(excel, "R", linha));
                    proposta.ValorOperacao = Convert.ToDecimal(Utils.TextoCelula(excel, "U", linha));
                    proposta.ValorFinanciado = Convert.ToDecimal(Utils.TextoCelula(excel, "Y", linha));
                    proposta.ValorParcela = Convert.ToDecimal(Utils.TextoCelula(excel, "Z", linha));
                    proposta.ValorCreditado = Convert.ToDecimal(Utils.TextoCelula(excel, "AB", linha));
                   
                    propostas.Add(proposta);

                    linha += 2;
                }
            }
            return Utils.ToXML<List<Proposta>>(propostas);
        }
    }
}
