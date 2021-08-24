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
    public class PreCreateWrapper : PluginHelper
    {
        /// <summary>
        ///  Triggers on PreCreate of Funeral Product
        ///  Copies value in BDM will be updated with the quantity of Funeral Product
        ///  Throws error in case of prepaid status active funeral
        ///  Sequnce number will map from product to funeral Product
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PreCreateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }
        public int funeralStatus = -1;
        public int prePaidStatus = -1;
        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "create" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("OpportunityProduct PreCreate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];

                if ((target == null || target.LogicalName != "opportunityproduct" || target.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }

                #region Translate Fields 
                Guid oppId = target.Contains("opportunityid") ? target.GetAttributeValue<EntityReference>("opportunityid").Id : Guid.Empty;
                string productdescription = target.Contains("productdescription") ? target.GetAttributeValue<string>("productdescription") : string.Empty;
                Guid productId = target.Contains("productid") ? target.GetAttributeValue<EntityReference>("productid").Id : Guid.Empty;
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

                        if ((productname.Contains("Death Certificate") || productname.Contains("CH009"))&& opportunity.Contains("ols_bdmid"))
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
                    throw new InvalidPluginExecutionException("No Data can be entered when Pre Paid Status is Active.");
                }

                #region PopSequenceNumber
                if (prod != null && prod.Contains("ols_sequencenumber"))
                {
                    target["ols_sequencenumber"] = prod.GetAttributeValue<int>("ols_sequencenumber");
                }
                #endregion

                AppendLog("OpportunityProduct PreCreate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
    }
    public class PreCreate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PreCreate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PreCreateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
