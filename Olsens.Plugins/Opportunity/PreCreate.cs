using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Opportunity
{
    public class PreCreateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PreCreate of Funeral
        /// If Bpay Number not found for related brand with empty Funeral number, throws error
        /// Generates Funeral Number & sets in Funeral
        /// Updates Account with funeral & Bpay number
        /// Updates Bpay number record with funeral number
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>

        public PreCreateWrapper(string unsecConfig, string securestring) : base(unsecConfig, securestring) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "create" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Opportunity PreCreate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];

                if ((target == null || target.LogicalName != "opportunity" || target.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }
                Entity bPayNumberEnt = GetBPayNumberEntity(target);
                if (bPayNumberEnt == null) throw new InvalidPluginExecutionException("You don't have free BPayNumber. You cannot create the Funeral.");

                string funeralNumber = "FN-" + FormatNum(GetCurrentNumber(), 6);
                target["ols_funeralnumber"] = funeralNumber;

                #region Update Informant
                if (target.Contains("customerid"))
                    UpdateInformant(target.GetAttributeValue<EntityReference>("customerid").Id, bPayNumberEnt.Id, funeralNumber);
                #endregion

                #region Update BPaynumber 
                bPayNumberEnt["ols_funeralnumber"] = funeralNumber;
                Update(UserType.User, bPayNumberEnt);
                #endregion

                target["ols_bpaynumberid"] = new EntityReference(bPayNumberEnt.LogicalName, bPayNumberEnt.Id);
                //string name = target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty;
                //target["ols_mortuaryregisterid"] = new EntityReference("ols_mortuaryregister", CreateMortuaryReg(name, funeralNumber));
                //target["ols_operationsid"] = new EntityReference("ols_funeraloperations", CreateOperations(name, funeralNumber));

                AppendLog("Opportunity PreCreate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #region [Public Methods]
        public Entity GetBPayNumberEntity(Entity target)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_bpaynumber");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralnumber", ConditionOperator.Null);
                if (target.Contains("pricelevelid"))
                    qe.Criteria.AddCondition("ols_brand", ConditionOperator.Equal, Util.GetBrandValue(GetService(UserType.User), target.GetAttributeValue<EntityReference>("pricelevelid").Id));
                else
                    qe.Criteria.AddCondition("ols_brand", ConditionOperator.Equal, 0); // HN Olsen Funerals Pty Ltd
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public int GetCurrentNumber()
        {
            try
            {
                int currentNo = 0;
                QueryExpression qe = new QueryExpression("ols_counter");
                qe.ColumnSet = new ColumnSet("ols_currentnumber");
                qe.Criteria.AddCondition("ols_name", ConditionOperator.Equal, "opportunity");
                Entity counter = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                if (counter != null)
                {
                    currentNo = counter.Contains("ols_currentnumber") ? counter.GetAttributeValue<int>("ols_currentnumber") : 0;

                    #region IncrementNumber
                    Entity updateCounter = new Entity("ols_counter", counter.Id);
                    updateCounter["ols_currentnumber"] = currentNo + 1;
                    Update(UserType.User, updateCounter);
                    #endregion
                }
                return currentNo;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public string FormatNum(int Num, int Digits)
        {

            try
            {
                string numStr = Num.ToString();

                for (int i = numStr.Length; i < Digits; i++)
                {
                    numStr = "0" + numStr;
                }

                return numStr;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }
        public void UpdateInformant(Guid informantId, Guid bPayNumberId, string funeralNumber)
        {
            try
            {
                Entity updateAccount = new Entity("account", informantId);
                updateAccount["ols_bpaynumberid"] = new EntityReference("ols_bpaynumber", bPayNumberId);
                updateAccount["ols_funeralnumber"] = funeralNumber;
                Update(UserType.User, updateAccount);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        #endregion
    }

    public class PreCreate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string Securestring = string.Empty;
        public PreCreate(string unsecConfig, string securestring)
        {
            UnsecConfig = unsecConfig;
            Securestring = securestring;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PreCreateWrapper(UnsecConfig, Securestring);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
