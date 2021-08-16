﻿using System;
using System.Collections.Generic;

namespace ExportExcel.Models
{
    public class Candidato
    {
        public int IdCandidato { get; set; }
        public string Nome { get; set; }
        public string NomeAbreviado { get; set; }
        public string EmailCandidato { get; set; }
        public string Partido { get; set; }        
        public string ImagemCandidato { get; set; }
        public string SituacaoCandidato { get; set; }
        public string Estado { get; set; }
        public string UF { get; set; }
        public string Cidade { get; set; }
        public string ViceCandidato { get; set; }
        public string EmailViceCandidato { get; set; }
        public string ImagemViceCandidato { get; set; }
        public string SituacaoViceCandidato { get; set; }
        public DateTime DataCadastro { get; set; }


        public static IList<Candidato> GetCandidatos()
        {
            return new List<Candidato> {
                new Candidato { IdCandidato = 1, Nome = "Angelo Andrea Matarazzo", NomeAbreviado = "Andrea Matarazzo", EmailCandidato = "aa.matarazzo@gmail.com", Partido = "PSD", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000661535_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Marta Costa", EmailViceCandidato = "depmartacosta@terra.com.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000661536_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 2, Nome = "Antônio Carlos Silva", NomeAbreviado = "Antônio Carlos", EmailCandidato = "cenpco2020@gmail.com", Partido = "PCO", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001172314_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Henrique Áreas", EmailViceCandidato = "cenpco2020@gmail.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001172315_div.jpg", SituacaoViceCandidato = "INATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 3, Nome = "Arthur Moledo Do Val", NomeAbreviado = "Arthur Do Val Mamãe Falei", EmailCandidato = "juridicopatriotasp@gmail.com", Partido = "PATRIOTA", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000641390_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Adelaide Oliveira", EmailViceCandidato = "juridicopatriotasp@gmail.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000641391_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 4, Nome = "Bruno Covas Lopes", NomeAbreviado = "Bruno Covas", EmailCandidato = "bruno@brunocovas.com.br", Partido = "PSDB", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000896546_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Ricardo Nunes", EmailViceCandidato = "15ricardo-nunes@uol.com.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000896547_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 5, Nome = "Celso Ubirajara Russomanno", NomeAbreviado = "Celso Russomanno", EmailCandidato = "jessica.cardoso@republicanos10capitalsp.org.br", Partido = "REPUBLICANOS", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001094597_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Marcos Da Costa", EmailViceCandidato = "marcos@marcosdacosta.adv.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001094596_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 6, Nome = "Guilherme Castro Boulos", NomeAbreviado = "Guilherme Boulos", EmailCandidato = "agendaboulos@gmail.com", Partido = "PSOL", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000746936_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Luiza Erundina", EmailViceCandidato = "perrucci32@gmail.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000746935_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 7, Nome = "Jilmar Augustinho Tatto", NomeAbreviado = "Jilmar Tatto", EmailCandidato = "helio@sap.adv.br", Partido = "PT", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000755896_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Carlos Zarattini", EmailViceCandidato = "helio@sap.adv.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000755897_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 8, Nome = "Joice Cristina Hasselmann", NomeAbreviado = "Joice Hasselmann", EmailCandidato = "atendimento@csc-eleitoral.com.br", Partido = "PSL", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000658458_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Ivan Sayeg", EmailViceCandidato = "jaynepc@escritoriobga.com.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000864821_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 9, Nome = "Jose Levy Fidelix Da Cruz", NomeAbreviado = "Levy Fidelix", EmailCandidato = "prtbdocumentos2020@gmail.com", Partido = "PRTB", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001013564_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Jairo Glikson", EmailViceCandidato = "prtbdocumentos2020@gmail.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001013565_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 10, Nome = "Marina Medeiros Helou", NomeAbreviado = "Marina Helou", EmailCandidato = "joaovcrezende@gmail.com", Partido = "REDE", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001152470_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Marco Dipreto", EmailViceCandidato = "marcodipretoredesp@gmai.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001152469_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 11, Nome = "Márcio Luiz França Gomes", NomeAbreviado = "Márcio França", EmailCandidato = "thiago@pomini.com.br", Partido = "PSB", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001012981_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Antonio Neto", EmailViceCandidato = "pdtspmunicipal12@gmail.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250001012982_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 12, Nome = "Orlando Silva De Jesus Junior", NomeAbreviado = "Orlando Silva", EmailCandidato = "pcdob2020@gmail.com", Partido = "PCdoB", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000766422_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Andrea Barcelos", EmailViceCandidato = "andrea.patty@uol.com.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000766421_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 13, Nome = "Filipe Tomazelli Sabara", NomeAbreviado = "Sabará", EmailCandidato = "monica@sabaranovo.com", Partido = "NOVO", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000693062_div.jpg", SituacaoCandidato = "INATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Marina Helena", EmailViceCandidato = "marina.h.santos@gmail.com", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000693061_div.jpg", SituacaoViceCandidato = "INATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") },
                new Candidato { IdCandidato = 14, Nome = "Vera Lucia Pereira Da Silva Salgado", NomeAbreviado = "Vera", EmailCandidato = "veralse16@gmail.com", Partido = "PSTU", ImagemCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000744464_div.jpg", SituacaoCandidato = "ATIVO", Estado = "São Paulo", UF = "SP", Cidade = "São Paulo", ViceCandidato = "Professor Lucas", EmailViceCandidato = "lansdp@yahoo.com.br", ImagemViceCandidato = "https://img.estadao.com.br/fotos/politica/eleicoes-2020/SP/FSP250000744465_div.jpg", SituacaoViceCandidato = "ATIVO", DataCadastro = Convert.ToDateTime("2021-08-10 23:56:44.603") }
            };
        }
    }
}
