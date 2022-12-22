using System;
namespace EWS_OAuth2
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {


            if (args.Length <2)
            {
                Console.WriteLine("Modo de uso: \n" +
                    "emails.exe -2 PASTA \n" +
                    "-2 últimos dias, PASTA = pasta-de-emails");
            }
            else
            {


                RelatorioBackup rel = new RelatorioBackup("19e0561e-e302-4058-a4be-ef27861568bf"
                                                            , "2c3e5fa7-3d0f-470e-a3b3-7e9945e1cfae"
                                                            , "NOF8Q~skd3P.YpKC7TysbuQuduKiLFcSo5lvNbUx"
                                                            , "suporte@tvconsultoria.com.br", args[1], int.Parse(args[0]));
                await rel.GerarRelatorio();
            }
            
        }
    }
}