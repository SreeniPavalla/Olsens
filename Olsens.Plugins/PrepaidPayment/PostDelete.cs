using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.PrepaidPayment
{
    public class PostDeleteWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostDelete of PrepaidPayment
        /// Calculates the total Paid to date, Prepaid GST amount, Last Payment date, Outstanding amount
        /// Updates these values to Funeral
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PostDeleteWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "delete" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is EntityReference)) return;

                AppendLog("PrePaidPayment PostDelete - Plugin Excecution is Started.");

                EntityReference target = (EntityReference)Context.InputParameters["Target"];
                Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;

                if ((target == null || target.LogicalName != "ols_prepaidfuneralpayments" || target.Id == Guid.Empty) || (preImage == null || preImage.Id == Guid.Empty))
                {
                    AppendLog("Target/PreImage is null");
                    return;
                }

                DateTime LastPaymentDate = new DateTime(1900, 1, 1);
                decimal PaidToDate = 0;
                decimal PrePaidGST = 0;
                Guid opportunityId = preImage.Contains("ols_funeralid") ? preImage.GetAttributeValue<EntityReference>("ols_funeralid").Id : Guid.Empty;
                if (opportunityId != Guid.Empty)
                {
                    EntityCollection prepaidPaymentsEnts = GetAllPrePaidPayments(opportunityId);
                    if (prepaidPaymentsEnts != null && prepaidPaymentsEnts.Entities.Count > 0)
                    {
                        AppendLog("PrepaidPayments count: "+ prepaidPaymentsEnts.Entities.Count);
                        foreach (var prepaidPayment in prepaidPaymentsEnts.Entities)
                        {
                            if (prepaidPayment.Contains("ols_dateofinstalment") && LastPaymentDate < prepaidPayment.GetAttributeValue<DateTime>("ols_dateofinstalment"))
                                LastPaymentDate = prepaidPayment.GetAttributeValue<DateTime>("ols_dateofinstalment").AddHours(5.5);
                            PaidToDate += prepaidPayment.Contains("ols_amountpaid") ? prepaidPayment.GetAttributeValue<Money>("ols_amountpaid").Value : 0;
                            PrePaidGST += prepaidPayment.Contains("ols_gst") ? prepaidPayment.GetAttributeValue<Money>("ols_gst").Value : 0;
                        }
                    }
                    UpdateOpportunity(opportunityId, LastPaymentDate, PaidToDate, PrePaidGST);
                    AppendLog("PrePaidPayment PostDelete - Plugin Excecution is Completed.");
                }
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public EntityCollection GetAllPrePaidPayments(Guid opportunityId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_prepaidfuneralpayments");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, opportunityId);
                return RetrieveMultiple(UserType.User, qe);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateOpportunity(Guid opportunityId, DateTime lastPaymentDate, decimal paidToDate, decimal gst)
        {
            try
            {
                Entity opp = Retrieve(UserType.User, "opportunity", opportunityId, new ColumnSet("ols_funeralamount"));
                Entity updateOpp = new Entity("opportunity", opportunityId);
                if (lastPaymentDate != new DateTime(1900, 1, 1))
                    updateOpp["ols_lastpaymentdate"] = lastPaymentDate;
                else
                    updateOpp["ols_lastpaymentdate"] = null;
                updateOpp["ols_paidtodate"] = paidToDate;

                decimal outstandingAmount = 0;
                if (opp != null && opp.Contains("ols_funeralamount"))
                    outstandingAmount = opp.GetAttributeValue<decimal>("ols_funeralamount") - paidToDate;
                updateOpp["ols_amountoutstanding"] = outstandingAmount;
                updateOpp["ols_prepaidgstamount"] = gst;
                Update(UserType.User, updateOpp);
            }
            catch (Exception ex) { throw ex; }
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
