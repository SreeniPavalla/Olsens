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
    public class PostUpdateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostUpdate of Funeral Product
        /// If send invitation is yes, Funeral having service place session from date and if product vendor is having contact
        /// Service place session Email will be sent to vendor
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PostUpdateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("OpportunityProduct PostUpdate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];
                Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;
                Entity postImage = Context.PostEntityImages.Contains("PostImage") ? (Entity)Context.PostEntityImages["PostImage"] : null;

                if ((target == null || target.LogicalName != "opportunityproduct" || target.Id == Guid.Empty) || (preImage == null || preImage.Id == Guid.Empty) || (postImage == null || postImage.Id == Guid.Empty))
                {
                    AppendLog("Target/PreImage/PostImage is null");
                    return;
                }
                if (Context.Depth > 1) { return; }
                bool preData_SendInvitation = preImage.Contains("ols_sendinvitation") ? preImage.GetAttributeValue<bool>("ols_sendinvitation") : false;
                bool postData_SendInvitation = postImage.Contains("ols_sendinvitation") ? postImage.GetAttributeValue<bool>("ols_sendinvitation") : false;
                if (postData_SendInvitation && postData_SendInvitation != preData_SendInvitation)
                {
                    Guid oppId = postImage.Contains("opportunityid") ? postImage.GetAttributeValue<EntityReference>("opportunityid").Id : Guid.Empty;
                    Create_SendEmail(Context.UserId, oppId, postImage);
                }
                AppendLog("OpportunityProduct PostUpdate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #region [Public Methods]
        public void Create_SendEmail(Guid userId, Guid oppId, Entity postImage)
        {
            var emailid = CreateEmail(userId, oppId, postImage);
            if (emailid != null)
            {
                SendEmail(new Guid(emailid));
            }
        }
        public void SendEmail(Guid emailid)
        {
            try
            {
                var sendemailreq = new Microsoft.Crm.Sdk.Messages.SendEmailRequest();
                sendemailreq.EmailId = emailid;
                sendemailreq.IssueSend = true;
                sendemailreq.TrackingToken = "";
                Execute(UserType.User, sendemailreq);
            }
            catch (Exception ex) { throw new Exception("Could not send email: " + ex.Message); }
        }
        public string CreateEmail(Guid userId, Guid oppId, Entity postImage)
        {
            try
            {
                var emailFrom = string.Empty;
                var body = string.Empty;
                var subject = string.Empty;

                Entity opp = Retrieve(UserType.User, "opportunity", oppId, new ColumnSet("ols_funeralnumber", "ols_deceasedfamilyname", "ols_serviceplacesessionfrom", "ownerid"));

                if (opp == null) return null;
                // emailFrom = getOwnerEmail(svc, ((EntityReference)opp.Attributes["ownerid"]).Id);

                Entity email = new Entity();
                email.LogicalName = "email";

                if (opp.Contains("ols_serviceplacesessionfrom"))
                {
                    var servicefrom = Util.LocalFromUTCUserDateTime(GetService(UserType.User), opp.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom"));
                    subject = opp["ols_funeralnumber"].ToString() + " " + opp["ols_deceasedfamilyname"].ToString() + " " + servicefrom.ToString("dd/MM/yyyy hh:mm tt");

                    //email["description"] = ModifyConfirmationBody(svc, email["description"].ToString()); //getConfirmationBody(svc);
                    // Entity activityTo = new Entity("activityparty");
                    Entity activityFrom = new Entity("activityparty");
                    //activityTo["partyid"] = new EntityReference("contact", achs_contactid.Id);
                    activityFrom["partyid"] = new EntityReference("systemuser", userId);



                    System.Text.StringBuilder descr = new System.Text.StringBuilder();
                    if (postImage.Contains("ols_vendorserviceid"))
                        descr.AppendLine("Vendor service: " + postImage.GetAttributeValue<EntityReference>("ols_vendorserviceid").Name);
                    if (postImage.Contains("ols_additionalvendorserviceinformation"))
                        descr.AppendLine("Additional details: " + postImage.GetAttributeValue<string>("ols_additionalvendorserviceinformation"));
                    if (postImage.Contains("ols_floraltributemessage"))
                        descr.AppendLine("Message: " + postImage.GetAttributeValue<string>("ols_floraltributemessage"));
                    if (postImage.Contains("ols_flowercolor"))
                        descr.AppendLine("Flower Color: " + postImage.GetAttributeValue<string>("ols_flowercolor"));
                    if (postImage.Contains("ols_ribboncolor"))
                        descr.AppendLine("Ribbon Color: " + postImage.GetAttributeValue<string>("ols_ribboncolor"));



                    if (postImage.Contains("ols_flowerdeliverytime") && postImage.Contains("ols_flowerdeliveredto"))
                    {
                        var flowerdeliverytime = Util.LocalFromUTCUserDateTime(GetService(UserType.User), postImage.GetAttributeValue<DateTime>("ols_flowerdeliverytime"));
                        descr.AppendLine("Flowers delivered To: " + postImage.GetAttributeValue<string>("ols_flowerdeliveredto") + " at " + flowerdeliverytime.ToString("dd/MM/yyyy hh:mm tt"));
                    }
                    else if (postImage.Contains("ols_flowerdeliveredto")) descr.AppendLine("Flowers delivered To: " + postImage.GetAttributeValue<string>("ols_flowerdeliveredto"));

                    body = descr.ToString();



                    email["from"] = new Entity[] { activityFrom };
                    email["description"] = body;
                    email["subject"] = subject;
                    email["scheduledstart"] = (DateTime)opp["ols_serviceplacesessionfrom"];

                    email["scheduledend"] = ((DateTime)opp["ols_serviceplacesessionfrom"]).AddMinutes(10);
                    email["regardingobjectid"] = new EntityReference("opportunity", oppId);
                    //Ent["location"] = location;
                    var party = new Entity("activityparty");
                    var contact = GetVendor(postImage);
                    if (contact != null && contact.Contains("ols_contactid"))
                    {
                        party["partyid"] = contact.GetAttributeValue<EntityReference>("ols_contactid");
                        email["to"] = new Entity[] { party };
                    }
                    else throw new Exception("Could not create an Email: Vendor contact is blank.");
                    // EntityHelper.CreateEntity(svc, Ent);


                }
                else throw new Exception("Could not create an Email: Service Date is blank.");


                return Create(UserType.User, email).ToString();

            }
            catch (Exception ex) { throw new Exception("The Email was not created. Error: " + ex.Message); }
        }
        public Entity GetVendor(Entity PostImage)
        {
            try
            {
                if (!PostImage.Contains("ols_productvendorid")) return null;
                return Retrieve(UserType.User, "ols_productvendor", PostImage.GetAttributeValue<EntityReference>("ols_productvendorid").Id, new ColumnSet("ols_contactid"));
            }
            catch { return null; }
        }
        #endregion
    }
    public class PostUpdate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PostUpdate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostUpdateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
