using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace Olsens.Plugins.Common
{
    public class ConfigData
    {
        public string httpServer;
        public string OrgName;
        public string Impersonate;
        public string DiscoverySvc;
        public string OrganizationSvc;
        public string pdfFolder;
        public string RptPath;
        public string RptWS;
        public string XMLPath;
        public string XMLURL;
        public string SMSAccount;
        public string SMSPassword;
        public string Message;
        public string SenderId;
        public string BDMRegAccount;
        public string BDMRegPassword;
        public string BDMRegWalterCarterAccount;
        public string BDMRegWalterCarterPassword;
        public string EndPointAddress;
        public string StatusQueryEndPointAddress;
        public string BDMRequestStatusTime;
        public string LongStayAttendee;
        public string MYOBURL;
        public string MYOBOnlieURL;
        public string MYOBAccount;
        public string MYOBPassword;
        public string MYOBCompanyName;
        public string MYOBHost;
        public string MYOBVersion;
        public string MYOBLibraryPath;

        public string MYOBClientId;
        public string MYOBClientSecret;
        public string MYOBRedirectURL;
        public string MYOBOnlineCompanyName;

        public string SMTPServer;
        public string SMTPPort;
        public string SMTPUser;
        public string SMTPPassword;
        //public decimal AddHoursTime = 0;

        public ConfigData(IOrganizationService svc)
        {
            try
            {
                if (svc != null)
                {
                    QueryExpression qe = new QueryExpression("ols_setting");
                    qe.ColumnSet = new ColumnSet("ols_name", "ols_value");
                    var settingEnts = svc.RetrieveMultiple(qe);

                    foreach (var settingEnt in settingEnts.Entities)
                    {
                        if (settingEnt.Contains("ols_name"))
                            switch (settingEnt.GetAttributeValue<string>("ols_name"))
                            {
                                case "MYOB_URL":
                                    MYOBURL = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Account":
                                    MYOBAccount = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Password":
                                    MYOBPassword = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_CompanyName":
                                    MYOBCompanyName = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Host":
                                    MYOBHost = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Version":
                                    MYOBVersion = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_LibraryPath":
                                    MYOBLibraryPath = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Online_ClientId":
                                    MYOBClientId = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Online_Secret":
                                    MYOBClientSecret = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Online_RedirectUrl":
                                    MYOBRedirectURL = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOB_Online_CompanyName":
                                    MYOBOnlineCompanyName = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SMTPServer":
                                    SMTPServer = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SMTPPort":
                                    SMTPPort = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SMTPUser":
                                    SMTPUser = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SMTPPassword":
                                    SMTPPassword = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "httpServer":
                                    httpServer = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "OrgName":
                                    OrgName = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "Impersonate":
                                    Impersonate = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "DiscoverySvc":
                                    DiscoverySvc = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "OrganizationSvc":
                                    OrganizationSvc = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "pdfFolder":
                                    pdfFolder = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "RptPath":
                                    RptPath = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "RptWS":
                                    RptWS = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "XMLURL":
                                    XMLURL = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "XMLPath":
                                    XMLPath = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SMSAccount":
                                    SMSAccount = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SMSPassword":
                                    SMSPassword = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "Message":
                                    Message = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "SenderId":
                                    SenderId = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "BDMRegAccount":
                                    BDMRegAccount = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "BDMRegPassword":
                                    BDMRegPassword = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "BDMRegWalterCarterAccount":
                                    BDMRegWalterCarterAccount = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "BDMRegWalterCarterPassword":
                                    BDMRegWalterCarterPassword = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "EndPointAddress":
                                    EndPointAddress = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "StatusQueryEndPointAddress":
                                    StatusQueryEndPointAddress = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "BDMRequestStatusTime":
                                    BDMRequestStatusTime = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "LongStayAttendee":
                                    LongStayAttendee = settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                case "MYOBOnline_URL":
                                    MYOBOnlieURL= settingEnt.Contains("ols_value") ? settingEnt.GetAttributeValue<string>("ols_value") : string.Empty; break;
                                    //case "AddHoursTime":
                                    //    if (settingEnt.Contains("ols_value"))
                                    //        decimal.TryParse(settingEnt.GetAttributeValue<string>("ols_value"), out AddHoursTime);
                                    //    break;
                            }
                    }
                }
            }
            catch (Exception e) { throw e; }
        }
    }
}
