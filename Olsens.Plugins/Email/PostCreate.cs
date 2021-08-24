using Microsoft.Xrm.Sdk;
using Olsens.Plugins.Common;
using Olsens.Plugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Email
{
    public class PostCreateWrapper : PluginHelper
    {
        public PostCreateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "create" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Email PostCreate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];

                if (target == null || target.LogicalName != "email" || target.Id == Guid.Empty)
                {
                    AppendLog("Target is null");
                    return;
                }

                if (target.Contains("subject") && (target.GetAttributeValue<string>("subject") == "Olsens Lead" || target.GetAttributeValue<string>("subject") == "Website Contact Lead"))
                {
                    AppendLog("Create lead for Email started.");
                    string description = target.Contains("description") ? target.GetAttributeValue<string>("description") : string.Empty;
                    LeadDetails objLeadDetails = SetLeadDetails(description);
                    Guid leadId = CreateLead(objLeadDetails);
                    if (leadId != Guid.Empty)
                        UpdateEmailRegarding(leadId, target);
                    AppendLog("Create lead for Email completed.");
                }

            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }
        public LeadDetails SetLeadDetails(string body)
        {
            var leaddetails = new LeadDetails();
            try
            {
                body = body.Replace("<BR>", "<br>").Replace("<br>", "<br>\n\r");

                using (var reader = new System.IO.StringReader(body))
                {
                    for (string line = reader.ReadLine(); line != null; line = reader.ReadLine())
                    {
                        if (line.Contains("first_name"))
                        {
                            leaddetails.firstname = line.Substring(line.IndexOf("first_name") + 11, line.IndexOf("</span>") - (line.IndexOf("first_name") + 11));
                        }
                        else
                         if (line.Contains("last_name"))
                        {
                            leaddetails.lastname = line.Substring(line.IndexOf("last_name") + 10, line.IndexOf("</span>") - (line.IndexOf("last_name") + 10));
                        }
                        else
                         if (line.Contains("user_email"))
                        {
                            leaddetails.email = line.Substring(line.IndexOf("user_email") + 11, line.IndexOf("</span>") - (line.IndexOf("user_email") + 11));
                        }
                        else
                         if (line.Contains("email"))
                        {
                            leaddetails.email = line.Substring(line.IndexOf("email") + 6, line.IndexOf("</span>") - (line.IndexOf("email") + 6));
                        }
                        else
                         if (line.Contains("telephone"))
                        {
                            leaddetails.telephone = line.Substring(line.IndexOf("telephone") + 10, line.IndexOf("</span>") - (line.IndexOf("telephone") + 10));
                        }
                        else
                         if (line.Contains("postcode"))
                        {
                            leaddetails.postcode = line.Substring(line.IndexOf("postcode") + 9, line.IndexOf("</span>") - (line.IndexOf("postcode") + 9));
                        }
                        else
                            if (line.Contains("<br>"))
                            leaddetails.description += line.Replace("<br>", "\n\r");
                    }
                }
                leaddetails.subject = leaddetails.firstname + " " + leaddetails.lastname;

            }
            catch (Exception ex) { throw new Exception("Error in SetLeadDetails: firstname=" + leaddetails.firstname + " lastname=" + leaddetails.lastname + "Error: " + ex.Message); }

            return leaddetails;
        }

        public Guid CreateLead(LeadDetails objLeadDetails)
        {
            try
            {
                Guid leadId = Guid.Empty;

                Entity createLead = new Entity("lead");
                createLead["firstname"] = objLeadDetails.firstname;
                createLead["lastname"] = objLeadDetails.lastname;
                createLead["emailaddress1"] = objLeadDetails.email;
                createLead["ols_description"] = objLeadDetails.description;
                createLead["address1_postalcode"] = objLeadDetails.postcode;
                createLead["telephone1"] = objLeadDetails.telephone;
                createLead["subject"] = objLeadDetails.subject;
                leadId = Create(UserType.User, createLead);
                return leadId;
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in CreateLead method");
                throw ex;
            }

        }

        public void UpdateEmailRegarding(Guid leadId, Entity target)
        {
            try
            {
                Entity updateEmail = new Entity(target.LogicalName, target.Id);
                updateEmail["regardingobjectid"] = new EntityReference("lead", leadId);
                Update(UserType.User, updateEmail);
            }
            catch (Exception ex) { throw ex; };
        }
    }
    public class PostCreate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PostCreate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostCreateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
