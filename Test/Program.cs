using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using Olsens.Plugins.MYOBOpportunity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestConsole
{
    public class Program
    {
        static void Main(string[] args)
        {
            // e.g. https://yourorg.crm.dynamics.com
            string url = "https://olsensdev2.crm6.dynamics.com/";
            // e.g. you@yourorg.onmicrosoft.com
            string userName = "Administrator@OlsensDev2.onmicrosoft.com";
            // e.g. y0urp455w0rd 
            string password = "pass@word1";

            string conn = $@"
    Url = {url};
    AuthType = OAuth;
    UserName = Administrator@OlsensDev2.onmicrosoft.com;
    Password = pass@word1;
    AppId = 51f81489-12ee-4a9e-aaae-a2591f45987d;
    RedirectUri = app://58145B91-0C36-4500-8554-080854F2AC97;
    LoginPrompt=Auto;
    RequireNewInstance = True";

            using (var svc = new CrmServiceClient(conn))
            {

                Entity opp = svc.Retrieve("opportunity", new Guid("E257E26A-80B1-EB11-8236-00224814BC01"), new ColumnSet(true));
                PostUpdateWrapper p = new PostUpdateWrapper("", "");
                p.ExecuteMYOB(true, new Guid("29B9C0C8-13A0-EB11-B1AC-00224815060C"), opp);




            }
        }
    }
}
