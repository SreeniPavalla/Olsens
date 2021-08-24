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
    public class PostUpdateWrapper : PluginHelper
    {
        /// <summary>
        /// This triggers on PostUpdate of Account entity. 
        /// If Need Type is Prepaid, Creates Quote & Quote products.
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>

        public PostUpdateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Account.PostUpdate - Started");

                Entity target = (Entity)Context.InputParameters["Target"];
                Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;
                Entity postImage = Context.PostEntityImages.Contains("PostImage") ? (Entity)Context.PostEntityImages["PostImage"] : null;

                if ((target == null || target.LogicalName != "account" || target.Id == Guid.Empty) || (preImage == null || preImage.Id == Guid.Empty) || (postImage == null || postImage.Id == Guid.Empty))
                {
                    AppendLog("Target/PreImage/PostImage is null");
                    return;
                }
                int preAccountType = preImage.Contains("ols_needtype") ? preImage.GetAttributeValue<OptionSetValue>("ols_needtype").Value : 0;
                int postAccountType = postImage.Contains("ols_needtype") ? postImage.GetAttributeValue<OptionSetValue>("ols_needtype").Value : 0;

                if (preAccountType != postAccountType && postAccountType == Convert.ToInt32(Constants.AccountType.PrePaid))
                {
                    IOrganizationService service = GetService(UserType.User);
                    AppendLog("CreateQuote method Started.");
                    Util.CreateQuote(target.Id, postImage.Contains("ols_prepaidnumber") ? postImage.GetAttributeValue<string>("ols_prepaidnumber") : string.Empty, service);
                    AppendLog("CreateQuote method completed.");
                }
                AppendLog("Account.PostUpdate - Completed");

            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }

    public class PostUpdate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PostUpdate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostUpdateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
