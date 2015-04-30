using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Xml.Serialization;

namespace ProducaoDaycoval
{
    public static class Utils
    {
        public static string ToXML<T>(T obj)
        {
            using (StringWriter stringWriter = new StringWriter(new StringBuilder()))
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
                xmlSerializer.Serialize(stringWriter, obj);
                return stringWriter.ToString();
            }
        }

        public static string TextoCelula(XlsFile excel, string coluna, int linha)
        {
            var numColuna = LetrasParaNumeros(coluna);
            return excel.GetStringFromCell(linha, numColuna);
        }

        public static string TextoCelula(XlsFile excel, string celula)
        {
            var linha = Convert.ToInt16(SoNumeros(celula));
            var coluna = LetrasParaNumeros(SoLetras(celula));
            return excel.GetStringFromCell(linha, coluna);
        }

        public static int IndiceLetra(string letra)
        {
            return ASCIIEncoding.ASCII.GetBytes(letra)[0] - 64;
        }

        public static int LetrasParaNumeros(string letras)
        {
            int resultado = 0;
            letras = letras.ToUpper();
            int fator = letras.Length;
            for (int i = 0; i < letras.Length; i++)
            {
                int o = IndiceLetra(letras.Substring(i));
                int p = (int)Math.Pow(26, fator - 1);
                resultado += o * p;
                fator--;
            }
            return resultado;
        }

        public static string SoNumeros(string texto)
        {
            string resultado = String.Empty;
            for (int i = 0; i < texto.Length; i++)
            {
                if ("0123456789".Contains(texto.Substring(i,1)))
                {
                    resultado += texto.Substring(i, 1);
                }
            }

            return resultado;
        }

        public static string SoLetras(string texto)
        {
            string resultado = String.Empty;
            for (int i = 0; i < texto.Length; i++)
            {
                if ("ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(texto.Substring(i,1).ToUpper()))
                {
                    resultado += texto.Substring(i, 1);
                }
            }

            return resultado;
        }
    }
}