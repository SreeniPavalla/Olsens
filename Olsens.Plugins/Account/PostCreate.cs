using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Account
{
    public class PostCreateWrapper : PluginHelper
    {
        /// <summary>
        /// This triggers on PostCreate of Account entity. 
        /// If Need Type is Prepaid, Creates Quote & Quote products.
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>

        public PostCreateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "create" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Account.PostCreate - Started");

                Entity target = (Entity)Context.InputParameters["Target"];

                if ((target == null || target.LogicalName != "account" || target.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }

                int accountType = target.Contains("ols_needtype") ? target.GetAttributeValue<OptionSetValue>("ols_needtype").Value : 0;
                if (accountType == Convert.ToInt32(Constants.AccountType.PrePaid))
                {
                    IOrganizationService service = GetService(UserType.User);
                    AppendLog("CreateQuote method Started.");
                    Util.CreateQuote(target.Id, target.Contains("ols_prepaidnumber") ? target.GetAttributeValue<string>("ols_prepaidnumber") : string.Empty, service);
                    AppendLog("CreateQuote method completed.");
                }
                AppendLog("Account.PostCreate - Completed");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }

    public class PostCreate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PostCreate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostCreateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
