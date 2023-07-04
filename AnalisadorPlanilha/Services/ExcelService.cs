﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnalisadorPlanilha.Services
{
    public static class ExcelService
    {
        public static void ProcessarLista(string diretorioDestinatarios,
                                          string colunasEmails,
                                          string planilhaDestinatarios,
                                          string diretorioEventos,
                                          string colunasEventos,
                                          string planilhaEventos)
        {
            Validar(diretorioDestinatarios, colunasEmails, diretorioEventos, colunasEventos);

            string[] colEmails = GetColunas(colunasEmails);
            string[] colEventos = GetColunas(colunasEventos);

            LerExcel(diretorioDestinatarios, colEmails);
        }

        private static string[] GetColunas(string coluna)
        {
            string[] colunas = coluna.Split(',');
            for (int i = 0; i < colunas.Length; i++)
            {
                colunas[i] = colunas[i].Trim();
            }

            return colunas;
        }

        private static void Validar(string diretorioDestinatarios,
                                          string colunasEmails,
                                          string planilhaDestinatarios,
                                          string diretorioEventos,
                                          string colunasEventos,
                                          string planilhaEventos)
        {
            if (string.IsNullOrEmpty(diretorioDestinatarios)) throw new Exception("Necessario informar os arquivos com os destinatarios do email");

            if (string.IsNullOrEmpty(colunasEmails)) throw new Exception("Necessario informar as colunas dos emails dos destinatarios");

            if (string.IsNullOrEmpty(diretorioEventos)) throw new Exception("Necessario informar o arquivo dos eventos");

            if (string.IsNullOrEmpty(colunasEventos)) throw new Exception("Necessario informar as colunas dos eventos a serem notificados");

            if (string.IsNullOrEmpty(planilhaDestinatarios)) throw new Exception("Necessario informar a planilha contendo os emails dos destinatarios");

            if (string.IsNullOrEmpty(planilhaEventos)) throw new Exception("Necessario informar a planilha contendo os eventos a serem notificados");
        }

        private static void LerExcel(string diretorio,
                                     string[] colunas)
        {
            if (!File.Exists(diretorio))
            {
                throw new FileNotFoundException($"O arquivo nao existe no diretorio {diretorio}");
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(diretorio))
            {

            }
        }
    }
}
