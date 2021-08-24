using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.BDMEmail
{
    public class PreCreateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PreCreate of Email Message. 
        /// If email creates with regarding entity as contact and contact name as "ELO ELO"
        /// Retrieves BDM record by comparing email TempId with BDM E Deaths ID
        /// Updates BDM Reg No with Death Certificate no
        /// Updated BDM dateofregistration with actualend date of email
        /// Updates Email regarding from contact to BDM
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>
        
        public PreCreateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        #region [Variables]
        public Guid BDM_Id = Guid.Empty;
        public string BDM_Name = string.Empty;
        public DateTime? actualEndDate = null;
        #endregion

        protected override void Execute()
        {
            if (Context.MessageName.ToLower() != "create" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

            AppendLog("EmailMessage PreCreate - Plugin Excecution is Started.");

            Entity target = (Entity)Context.InputParameters["Target"];

            if ((target == null || target.LogicalName != "email" || target.Id == Guid.Empty))
            {
                AppendLog("Target is null");
                return;
            }
            if (target.Contains("regardingobjectid") && target.GetAttributeValue<EntityReference>("regardingobjectid").LogicalName == "contact" && target.GetAttributeValue<EntityReference>("regardingobjectid").Name == "ELO ELO")
            {
                AppendLog("BDM Email identified");

                #region Translate Fields
                string description = target.Contains("description") ? target.GetAttributeValue<string>("description") : string.Empty;
                actualEndDate = target.Contains("actualend") ? target.GetAttributeValue<DateTime>("actualend") : (DateTime?)null;

                var tempId = GetTempId(description);
                if (!string.IsNullOrEmpty(tempId))
                {
                    Entity bdm = GetBDM(tempId);
                    if (bdm == null) return;

                    AppendLog("BDM found with E Deaths ID: " + tempId);
                    var reg_number = GetData(description, "Death Certificate No:");
                    if (!string.IsNullOrEmpty(reg_number) && !string.IsNullOrEmpty(tempId))
                    {
                        if (UpdateBDM(bdm, reg_number, tempId))
                        {
                            BDM_Id = bdm.Id;
                            BDM_Name = bdm.Contains("ols_name") ? bdm.GetAttributeValue<string>("ols_name") : string.Empty;
                        }
                    }
                }
                #endregion

                if (BDM_Id != Guid.Empty)
                {
                    EntityReference regardingRef = new EntityReference("ols_bdm", BDM_Id);
                    regardingRef.Name = BDM_Name;
                    target["regardingobjectid"] = regardingRef;
                }
            }
            AppendLog("EmailMessage PreCreate - Plugin Excecution is competed.");
        }

        #region [Public Methods]
        public string GetTempId(string description)
        {

            int startIndex = description.IndexOf("Temporary ID#:");

            if (startIndex < 0)
                return string.Empty;

            startIndex += 14;

            if (startIndex >= description.Length)
                return string.Empty;

            int endIndex = description.IndexOf(')', startIndex);

            if (endIndex < 0)
                return string.Empty;
            return description.Substring(startIndex, endIndex - startIndex).Trim();

        }
        public Entity GetBDM(string tempId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_bdm");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_edeathsid", ConditionOperator.Equal, tempId);
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public string GetData(string description, string startword)
        {
            // var lower = description.ToLower();
            int startIndex = description.IndexOf(startword);

            if (startIndex < 0)
                return string.Empty;

            startIndex += startword.Length;

            if (startIndex >= description.Length)
                return string.Empty;

            int endIndex = description.IndexOf('<', startIndex);

            if (endIndex < 0)
                return string.Empty;

            return description.Substring(startIndex, endIndex - startIndex).Trim();

        }
        public bool UpdateBDM(Entity BDM, string reg_number, string tempId)
        {
            try
            {
                AppendLog("Updating BDM with registrationnumber: " + reg_number + " & dateofregistration: " + actualEndDate);
                Entity updateEntity = new Entity(BDM.LogicalName, BDM.Id);
                updateEntity["ols_registrationnumber"] = reg_number + "/" + DateTime.Now.Year.ToString();
                if (actualEndDate != null)
                {
                    updateEntity["ols_dateofregistration"] = actualEndDate;
                }
                Update(UserType.User, updateEntity);
                AppendLog("BDM is updated");
                return true;
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion
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
