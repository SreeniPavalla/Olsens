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
    public class PreUpdateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PreUpdate of Funeral Product
        /// Updates Copies value in BDM by the quantity of funeral product
        /// Throws error in case of prepaid status active funeral
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PreUpdateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        #region [Variables]
        public int funeralStatus = -1;
        public int prePaidStatus = -1;
        #endregion

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("OpportunityProduct PreUpdate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];
                Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;

                if ((target == null || target.LogicalName != "opportunityproduct" || target.Id == Guid.Empty || preImage == null || preImage.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }
                if (Context.Depth > 1) { return; }

                #region Translate Fields 
                Guid oppId = preImage.Contains("opportunityid") ? preImage.GetAttributeValue<EntityReference>("opportunityid").Id : Guid.Empty;
                string productdescription = target.Contains("productdescription") ? target.GetAttributeValue<string>("productdescription") : string.Empty;
                if (string.IsNullOrEmpty(productdescription))
                    productdescription = preImage.Contains("productdescription") ? preImage.GetAttributeValue<string>("productdescription") : string.Empty;
                Guid productId = target.Contains("productid") ? target.GetAttributeValue<EntityReference>("productid").Id : Guid.Empty;
                if (productId == Guid.Empty)
                    productId = preImage.Contains("productid") ? preImage.GetAttributeValue<EntityReference>("productid").Id : Guid.Empty;

                decimal quantity = target.Contains("quantity") ? target.GetAttributeValue<decimal>("quantity") : 0;
                Entity prod = null;
                if (oppId != Guid.Empty)
                {
                    Entity opportunity = Retrieve(UserType.User, "opportunity", oppId, new ColumnSet("ols_status", "ols_prepaidstatus", "ols_bdmid"));
                    if (opportunity != null)
                    {
                        if (opportunity.Contains("ols_prepaidstatus"))
                            prePaidStatus = opportunity.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value;
                        if (opportunity.Contains("ols_status"))
                            funeralStatus = opportunity.GetAttributeValue<OptionSetValue>("ols_status").Value;

                        var productname = string.Empty;
                        if (!string.IsNullOrEmpty(productdescription))
                            productname = productdescription;
                        else
                        if (productId != Guid.Empty)
                        {
                            prod = Retrieve(UserType.User, "product", productId, new ColumnSet("productnumber", "ols_sequencenumber"));
                            if (prod != null) productname = prod.Contains("productnumber") ? prod.GetAttributeValue<string>("productnumber") : string.Empty;
                        }

                        if (productname.Contains("Death Certificate") || productname.Contains("CH009"))
                        {
                            Entity updateBDM = new Entity("ols_bdm", opportunity.GetAttributeValue<EntityReference>("ols_bdmid").Id);
                            updateBDM["ols_copies"] = (int)quantity;
                            Update(UserType.User, updateBDM);
                        }
                    }
                }
                #endregion

                if (funeralStatus == (int)Constants.FuneralStatus.PrePaid && prePaidStatus == (int)Constants.PrePaidStatus.Active)
                {
                    throw new Exception("No Data can be updated when Pre Paid Status is Active.");
                }
                AppendLog("OpportunityProduct PreUpdate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
    public class PreUpdate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PreUpdate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PreUpdateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
