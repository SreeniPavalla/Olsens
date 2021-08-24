using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.OpportunityProduct
{
    public class PreDeleteWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PreDelete of Funeral Product
        /// Throws error in case of prepaid status active funeral
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PreDeleteWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        #region [Variables]
        public int funeralStatus = -1;
        public int prePaidStatus = -1;
        #endregion

        protected override void Execute()
        {
            if (Context.MessageName.ToLower() != "delete") return;
            EntityReference target = (EntityReference)Context.InputParameters["Target"];
            Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;

            if (target == null || target.LogicalName != "opportunityproduct" || target.Id == Guid.Empty || preImage == null || preImage.Id == Guid.Empty)
            {
                AppendLog("Target is null");
                return;
            }
            AppendLog("opportunityproduct PreDelete - Plugin Excecution is Started.");

            #region Translate Fields 
            Guid oppId = preImage.Contains("opportunityid") ? preImage.GetAttributeValue<EntityReference>("opportunityid").Id : Guid.Empty;
            if (oppId != Guid.Empty)
            {
                Entity opportunity = Retrieve(UserType.User, "opportunity", oppId, new ColumnSet("ols_status", "ols_prepaidstatus"));
                if (opportunity != null)
                {
                    if (opportunity.Contains("ols_prepaidstatus"))
                        prePaidStatus = opportunity.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value;
                    if (opportunity.Contains("ols_status"))
                        funeralStatus = opportunity.GetAttributeValue<OptionSetValue>("ols_status").Value;
                }
            }
            #endregion

            if (funeralStatus == (int)Constants.FuneralStatus.PrePaid && prePaidStatus == (int)Constants.PrePaidStatus.Active)
            {
                throw new Exception("No Data can be deleted when Pre Paid Status is Active.");
            }

            if (funeralStatus == (int)Constants.FuneralStatus.PrePaid && prePaidStatus == (int)Constants.PrePaidStatus.AtNeed && preImage.Contains("ols_prepaid") && preImage.GetAttributeValue<bool>("ols_prepaid"))
            {
                throw new Exception("Pre Paid Charges cannot be deleted.");
            }
            AppendLog("opportunityproduct PreDelete - Plugin Excecution is Completed.");
        }
    }
    public class PreDelete : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PreDelete(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PreDeleteWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
