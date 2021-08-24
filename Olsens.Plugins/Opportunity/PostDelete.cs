using Microsoft.Xrm.Sdk;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Opportunity
{
    public class PostDeleteWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostDelete of Funeral. 
        /// Deletes Operation & Mortuary Register
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PostDeleteWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            if (Context.MessageName.ToLower() != "delete" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is EntityReference)) return;

            AppendLog("Opportunity PostDelete - Plugin Excecution is Started.");

            Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;

            if (preImage == null || preImage.Id == Guid.Empty)
            {
                AppendLog("PreImage is null");
                return;
            }

            EntityReference mortuaryRegRef = preImage.Contains("ols_mortuaryregisterid") ? preImage.GetAttributeValue<EntityReference>("ols_mortuaryregisterid") : null;
            EntityReference operationsRef = preImage.Contains("ols_operationsid") ? preImage.GetAttributeValue<EntityReference>("ols_operationsid") : null;
            if (mortuaryRegRef != null)
            {
                Delete(UserType.User, mortuaryRegRef.LogicalName, mortuaryRegRef.Id);
                AppendLog("Mortuary Register is deleted");
            }
            if (operationsRef != null)
            {
                Delete(UserType.User, operationsRef.LogicalName, operationsRef.Id);
                AppendLog("Operation is deleted");

            }
            AppendLog("Opportunity PostDelete - Plugin Excecution is Completed.");
        }
    }
    public class PostDelete : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PostDelete(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostDeleteWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
