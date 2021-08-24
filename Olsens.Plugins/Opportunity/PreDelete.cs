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
    public class PreDeleteWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PreDelete of Funeral
        /// Allows only Admin to delete funeral else throws error
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PreDeleteWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            if (Context.MessageName.ToLower() != "delete") return;

            AppendLog("Opportunity PreDelete - Plugin Excecution is Started.");

            if (!IsAdmin(Context.InitiatingUserId))
                throw new InvalidPluginExecutionException("Only System Administrator can delete the Funeral record.");

            AppendLog("Opportunity PreDelete - Plugin Excecution is Completed.");
        }

        public bool IsAdmin(Guid userId)
        {
            string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                      <entity name='systemuser'>
                        <attribute name='fullname' />
                        <attribute name='businessunitid' />
                        <attribute name='title' />
                        <attribute name='address1_telephone1' />
                        <attribute name='positionid' />
                        <attribute name='systemuserid' />
                        <order attribute='fullname' descending='false' />
                        <filter type='and'>
                          <condition attribute='systemuserid' operator='eq' uitype='systemuser' value='{0}' />
                        </filter>
                        <link-entity name='systemuserroles' from='systemuserid' to='systemuserid' visible='false' intersect='true'>
                          <link-entity name='role' from='roleid' to='roleid' alias='ae'>
                            <filter type='and'>
                              <condition attribute='name' operator='eq' value='System Administrator' />
                            </filter>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
            Entity user = RetrieveMultiple(UserType.User, new FetchExpression(string.Format(fetch, userId))).Entities.FirstOrDefault();
            if (user != null)
                return true;
            else
                return false;
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
            var pluginCode = new PreCreateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
