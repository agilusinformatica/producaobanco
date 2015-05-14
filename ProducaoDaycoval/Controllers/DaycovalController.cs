using FlexCel.XlsAdapter;
using ProducaoDaycoval;
using ProducaoDaycoval.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace ProducaoDaycoval.Controllers
{
    public class DaycovalController : Controller
    {
        //
        // GET: /Home/

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(Upload upload)
        {
            string nomeArquivo = @"e:\home\agilus\Temp\" + Path.GetRandomFileName().Replace(".", "");

            upload.arquivo.SaveAs(nomeArquivo);
            XlsFile excel = new XlsFile(nomeArquivo);
            string resposta = "";

            if (Utils.TextoCelula(excel, "C8") == "RELATÓRIO DE PROPOSTAS CADASTRADAS")
            {
                resposta = SerializaProposta(excel);
            }

            System.IO.File.Delete(nomeArquivo);
            return this.Content(resposta, "text/xml");
            //return this.Content(resposta, "application/json");
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
                else if (CelulaA.StartsWith("GERENTE:"))
                    gerente = CelulaA.Substring(9);
                else if (CelulaA.StartsWith("PROMOTORA:"))
                    promotora = CelulaA.Substring(11);
                else if (CelulaA.StartsWith("EMPREGADOR:"))
                    empregador = CelulaA.Substring(12);
                else if (CelulaA.StartsWith("ORGAO:"))
                    orgao = CelulaA.Substring(7);
                else if (CelulaA.StartsWith("REGRA:"))
                    regra = CelulaA.Substring(7);
                else if (CelulaA.StartsWith("SITUAÇÃO:"))
                    situacao = CelulaA.Substring(10);
                else if (CelulaA.StartsWith("AGENTE:"))
                    agente = CelulaA.Substring(8);
                else if ((CelulaA == Utils.SoNumeros(CelulaA)) && (CelulaA != String.Empty))
                {
                    var proposta = new Proposta();
                    proposta.Filial = filial;
                    proposta.Gerente = gerente;
                    proposta.Promotora = promotora;
                    proposta.Empregador = empregador;
                    proposta.Orgao = orgao;
                    proposta.Regra = regra;
                    proposta.Situacao = situacao;
                    proposta.Agente = agente;

                    proposta.NumeroProposta = Utils.TextoCelula(excel, "A", linha);
                    proposta.Cliente = Utils.TextoCelula(excel, "D", linha);
                    proposta.Cpf = Utils.TextoCelula(excel, "E", linha);
                    proposta.Matricula = Utils.TextoCelula(excel, "F", linha);
                    proposta.DataCadastro = DateTime.ParseExact(Utils.TextoCelula(excel, "G", linha), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    proposta.DataBase = DateTime.ParseExact(Utils.TextoCelula(excel, "H", linha), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    proposta.Prazo = Convert.ToInt32(Utils.TextoCelula(excel, "L", linha));
                    proposta.QtdeParcelas = Convert.ToInt32(Utils.TextoCelula(excel, "M", linha));
                    proposta.Taxa = Convert.ToDecimal(Utils.TextoCelula(excel, "O", linha));
                    proposta.ValorLiquido = Convert.ToDecimal(Utils.TextoCelula(excel, "P", linha));
                    proposta.Iof = Convert.ToDecimal(Utils.TextoCelula(excel, "Q", linha));
                    proposta.ValorOperacao = Convert.ToDecimal(Utils.TextoCelula(excel, "S", linha));
                    proposta.ValorFinanciado = Convert.ToDecimal(Utils.TextoCelula(excel, "U", linha));
                    if (proposta.Prazo != 0)
                    {
                        proposta.DataUltimoVencimento = DateTime.ParseExact(Utils.TextoCelula(excel, "I", linha), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        proposta.ValorParcela = Convert.ToDecimal(Utils.TextoCelula(excel, "X", linha));
                        proposta.ValorCreditado = Convert.ToDecimal(Utils.TextoCelula(excel, "Y", linha));
                    }
                    propostas.Add(proposta);

                    linha += 2;
                }
                
            }
            return Utils.ToXML<List<Proposta>>(propostas);
            //return new JavaScriptSerializer().Serialize(propostas);
        }

    }
}
