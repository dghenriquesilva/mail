using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Web;
using Microsoft.Identity.Client;
namespace EWS_OAuth2
{
    class Cliente
    {

       //Email info
        public string EmailAssunto { get; set; }
        public string EmailOrigem { get; set; }
        public string EmailDestino { get; set; }

        // Email backup
        public string DataRecebimento { get; set; }
        public string DataBackup { get; set; }
        public string versaoProduto { get; set; }
        public string Empresa { get; set; }
        public string Computador { get; set; }
        public string Duracao { get; set; }
        public string arquivosVerificadosGB { get; set; }
        public string backupDiferencialGB { get; set; }
        public string arquivosComFalha { get; set; }

        public Cliente() { }


    }
}
