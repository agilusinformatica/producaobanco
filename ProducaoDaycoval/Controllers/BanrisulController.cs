using FlexCel.XlsAdapter;
using ProducaoDaycoval.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ProducaoDaycoval.Controllers
{
    public class BanrisulController : Controller
    {
        //
        // GET: /Banrisul/

        [HttpGet]
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

            if (Utils.TextoCelula(excel, "B2").Equals("Relatório Produção Diária"))
            {
                resposta = SerializaProposta(excel);
            }

            System.IO.File.Delete(nomeArquivo);
            return this.Content(resposta, "text/xml");
        }
        private string SerializaProposta(XlsFile excel)
        {
            var propostas = new List<Proposta>();
            string gerente = String.Empty,
            promotora = String.Empty,
            agente = String.Empty,
            empregador = String.Empty;

            int QtdeLinhas = excel.RowCount;
            for (int linha = 7; linha <= QtdeLinhas; linha++)
            {
                string CelulaA = Utils.TextoCelula(excel, "A", linha);
                string CelulaC = Utils.TextoCelula(excel, "C", linha);

                if (CelulaA.StartsWith("Gerência :"))
                    gerente = CelulaC;
                else if (CelulaA.StartsWith("Corresp :"))
                    promotora = CelulaC;
                else if (CelulaA.StartsWith("Agente :"))
                    agente = CelulaC;
                else if (CelulaA.StartsWith("Conveniada :"))
                    empregador = CelulaC;

                else if ((CelulaA.Replace("*", "") == Utils.SoNumeros(CelulaA)) && (CelulaA != String.Empty))
                {
                    var proposta = new Proposta();
                    proposta.Gerente = gerente;
                    proposta.Promotora = promotora;
                    proposta.Agente = agente;
                    proposta.Empregador = empregador;
                    proposta.NumeroProposta = CelulaA;
                    proposta.Cliente = CelulaC;
                    proposta.Cpf = Utils.TextoCelula(excel, "D", linha);
                    if (Utils.TextoCelula(excel, "R", linha) != "")
                        proposta.DataBase = DateTime.ParseExact(Utils.TextoCelula(excel, "R", linha), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    proposta.QtdeParcelas = Convert.ToInt32(Utils.TextoCelula(excel, "J", linha));
                    proposta.Situacao = Utils.TextoCelula(excel, "L", linha);
                    proposta.Regra = Utils.TextoCelula(excel, "P", linha);
                    proposta.ValorOperacao = Convert.ToDecimal(Utils.TextoCelula(excel, "E", linha));
                    proposta.ValorLiquido = Convert.ToDecimal(Utils.TextoCelula(excel, "G", linha));
                    proposta.Iof = Convert.ToDecimal(Utils.TextoCelula(excel, "I", linha));
                    proposta.ValorParcela = Convert.ToDecimal(Utils.TextoCelula(excel, "K", linha));

                    propostas.Add(proposta);
                }
            }
            return Utils.ToXML<List<Proposta>>(propostas);
        }
    }
}
