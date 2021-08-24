using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Music
{
    public class PreValidationWrapper : PluginHelper
    {
        public PreValidationWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (!Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Opportunity PostCreate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];

                if ((target == null || target.LogicalName != "ols_music" || target.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }
                if (Context.MessageName.ToLower() == "create")
                {
                    if (target.Contains("ols_sequencenumber"))
                    {
                        Guid oppId = target.Contains("ols_funeralid") ? target.GetAttributeValue<EntityReference>("ols_funeralid").Id : Guid.Empty;
                        int seqNumber = target.Contains("ols_sequencenumber") ? target.GetAttributeValue<int>("ols_sequencenumber") : 0;

                        bool isValidSequence = ValidateSequenceNumber(oppId, seqNumber);
                        if (!isValidSequence)
                            throw new InvalidPluginExecutionException("Duplicate Sequence Number found!");
                    }
                }
                else if (Context.MessageName.ToLower() == "update")
                {
                    Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;

                    if (target.Contains("ols_sequencenumber"))
                    {
                        Guid oppId = preImage.Contains("ols_funeralid") ? preImage.GetAttributeValue<EntityReference>("ols_funeralid").Id : Guid.Empty;
                        int seqNumber = target.Contains("ols_sequencenumber") ? target.GetAttributeValue<int>("ols_sequencenumber") : 0;

                        bool isValidSequence = ValidateSequenceNumber(oppId, seqNumber);
                        if (!isValidSequence)
                            throw new InvalidPluginExecutionException("Duplicate Sequence Number found!");
                    }
                }

            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public bool ValidateSequenceNumber(Guid oppId, int seqNumber)
        {
            bool isValid = true;
            string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='ols_music'>
                        <attribute name='ols_musicid' />
                        <attribute name='ols_name' />
                        <attribute name='createdon' />
                        <attribute name='ols_source' />
                        <attribute name='ols_songdetails' />
                        <attribute name='ols_serviceschedule' />
                        <attribute name='ols_songinstructions' />
                        <attribute name='ols_sequencenumber' />
                        <order attribute='createdon' descending='false' />
                        <filter type='and'>
                          <condition attribute='ols_funeralid' operator='eq' uitype='opportunity' value='{0}' />
                        </filter>
                      </entity>
                    </fetch>";
            EntityCollection musicColl = RetrieveMultiple(UserType.User, new FetchExpression(string.Format(fetch, oppId)));
            if (musicColl != null)
            {
                foreach (var item in musicColl.Entities)
                {
                    int musicSeqNumber = item.Contains("ols_sequencenumber") ? item.GetAttributeValue<int>("ols_sequencenumber") : 0;
                    if (musicSeqNumber == seqNumber && seqNumber != 0)
                        return false;
                }
            }
            return isValid;
        }
    }

    public class PreValidation : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PreValidation(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PreValidationWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
