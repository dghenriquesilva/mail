using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Web;
using Microsoft.Identity.Client;
namespace EWS_OAuth2
{
    class RelatorioBackup
    {
        // Using Microsoft.Identity.Client

        private string TenantID { get; set; }
        private string ClientID { get; set; }
        private string ClientSecret { get; set; }
        private string Email { get; set; }
        private string Pasta { get; set; }
        private string URL { get; set; }

        /* var  cca = ConfidentialClientApplicationBuilder
       .Create("2c3e5fa7-3d0f-470e-a3b3-7e9945e1cfae")      //client Id
       .WithClientSecret("NOF8Q~skd3P.YpKC7TysbuQuduKiLFcSo5lvNbUx")
       .WithTenantId("19e0561e-e302-4058-a4be-ef27861568bf")
       .Build();
           */

        public RelatorioBackup(string tenantid, string clientid, string clientsecret, string email, string pasta = "Backup", string url = "https://outlook.office365.com/.default")
        {
            this.TenantID = tenantid;
            this.ClientID = clientid;
            this.ClientSecret = clientsecret;
            this.Email = email;
            this.Pasta = pasta;
            this.URL = url;
        }
        public async System.Threading.Tasks.Task GerarRelatorio(int diashistoricoemail = -7, int quantidadeemail = 250)
        {
            // Conecta ao email
            var cca = ConfidentialClientApplicationBuilder
                .Create(ClientID)
                .WithClientSecret(ClientSecret)
                .WithTenantId(TenantID)
                .Build();

            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };
            Cliente cliente = new Cliente();

            try
            {
                // Get token
                var authResult = await cca.AcquireTokenForClient(ewsScopes)
                                    .ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
                ewsClient.ImpersonatedUserId =
                new ImpersonatedUserId(ConnectingIdType.SmtpAddress, Email);

                //Include x-anchormailbox header
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", Email);

                // Ler todas as mensagens dos últimos -1 dia
                TimeSpan ts = new TimeSpan(-1, 0, 0, 0);
                DateTime date = DateTime.Now.Add(ts);


                // Make an EWS call to list folders on exhange online
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(20));
                foreach (var folder in folders.Result)
                {
                    if (folder.DisplayName.ToString() == Pasta)
                    {

                        PropertySet FindItemPropertySet = new PropertySet(BasePropertySet.IdOnly);

                        ItemView view = new ItemView(999);
                        view.PropertySet = FindItemPropertySet;
                        PropertySet GetItemsPropertySet = new PropertySet(BasePropertySet.FirstClassProperties);
                        //GetItemsPropertySet.RequestedBodyType = BodyType.Text;
                        //SearchFilter searchFilter =  new SearchFilter.SearchFilterCollection(LogicalOperator.And, new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
                        SearchFilter.IsGreaterThanOrEqualTo filtro = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);
                        FindItemsResults<Item> emailMessages = null;
                        do
                        {
                            emailMessages = await ewsClient.FindItems(folder.Id, filtro, view);
                            if (emailMessages.Items.Count > 0)
                            {

                                await ewsClient.LoadPropertiesForItems(emailMessages.Items, GetItemsPropertySet);

                               foreach ( Item Item in  emailMessages.Items)
                                {
                                    //if (Item.DisplayTo == "backup@tvconsultoria.com.br")
                                    //{
                                    
                                        cliente.DataRecebimento = Item.DateTimeReceived.ToString();
                                        cliente.EmailAssunto = Item.Subject.ToString();
                                    StringBuilder sb = new StringBuilder();
                                   int pi =  Console.WriteLine(Item.Body.ToString().IndexOf("Empresa"));
                                    int pf = Console.WriteLine(Item.)
                                    //Console.WriteLine(cliente.EmailAssunto);
                                   // Console.WriteLine(Item.Body.ToString());

                                }
                            }
                        } while (emailMessages.MoreAvailable);

                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}

