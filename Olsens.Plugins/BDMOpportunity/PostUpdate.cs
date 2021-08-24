using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using Olsens.Plugins.LodgeDRSService;
using Olsens.Plugins.Model;
using Olsens.Plugins.StatusQueryService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace Olsens.Plugins.BDMOpportunity
{
    public class PostUpdateWrapper : PluginHelper
    {
        #region [Variables]
        private ConfigData _config;
        Entity brand = null;
        DRSRequestType DRSRequest = null;
        XmlDocument XMLData = null;
        EntityReference xmlContainerId = null;
        string funeralNumber = string.Empty;
        BDMFuneralData objBDMFuneralData = null;
        private string Message = string.Empty;
        private string To = string.Empty;
        private string EmailTo = string.Empty;
        private int ContactMethod = 0;
        SMSData smsData;
        #endregion

        public PostUpdateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

            AppendLog("Opportunity PostUpdate - Plugin Excecution is Started.");

            Entity target = (Entity)Context.InputParameters["Target"];
            //Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;
            Entity postImage = Context.PostEntityImages.Contains("PostImage") ? (Entity)Context.PostEntityImages["PostImage"] : null;

            if ((target == null || target.LogicalName != "ols_bdm" || target.Id == Guid.Empty) || (postImage == null || postImage.Id == Guid.Empty))
            {
                AppendLog("Target/PostImage is null");
                return;
            }
            if (Context.Depth > 1)
                return;

            EntityReference funeralRef = postImage.Contains("ols_funeralid") ? postImage.GetAttributeValue<EntityReference>("ols_funeralid") : null;
            Entity funeral = null;
            if (funeralRef != null)
            {
                funeral = Retrieve(UserType.User, funeralRef.LogicalName, funeralRef.Id, new ColumnSet(true));
                if (funeral != null)
                {
                    funeralNumber = funeral.Contains("ols_funeralnumber") ? funeral.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
                }
            }
            xmlContainerId = postImage.Contains("ols_xmlcontainerid") ? postImage.GetAttributeValue<EntityReference>("ols_xmlcontainerid") : null;


            #region BDM Registration
            if (target.Contains("ols_sendtobdm") && target.GetAttributeValue<bool>("ols_sendtobdm") && funeral != null)
            {
                objBDMFuneralData = new BDMFuneralData();
                if (xmlContainerId != null)
                    objBDMFuneralData.XMLId = xmlContainerId.Id;
                AppendLog("BDM: Send To BDM started");
                BDMExecute(funeral, postImage, Context.UserId);
                AppendLog("BDM: Send To BDM completed");

            }
            if (target.Contains("ols_sendstatusrequest") && target.GetAttributeValue<bool>("ols_sendstatusrequest"))
            {
                objBDMFuneralData = new BDMFuneralData();
                AppendLog("BDM: Send To BDM Status Request started");
                RequestStatus(postImage);
                AppendLog("BDM: Send To BDM Status Request completed");

            }
            #endregion

            #region DefaultDoctor
            if (funeral != null && funeral.Contains("ols_funeraltype") && funeral.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == 1) //Cremation
            {
                EntityReference mortuaryRef = funeral.Contains("ols_mortuaryregisterid") ? funeral.GetAttributeValue<EntityReference>("ols_mortuaryregisterid") : null;
                if (mortuaryRef != null)
                {
                    string doctorname = "Downing, Dr Margaret";
                    Guid medicalrefereeId = GetDoctorIdByName(doctorname);
                    if (medicalrefereeId != Guid.Empty)
                    {
                        Entity updateMortuary = new Entity(mortuaryRef.LogicalName, mortuaryRef.Id);
                        updateMortuary["ols_medicalrefereeid"] = new EntityReference("ols_medicalreferee", medicalrefereeId);
                        Update(UserType.User, updateMortuary);
                    }
                }
            }
            #endregion

            #region SMS/Email
            if (target.Contains("ols_sendsms") && target.GetAttributeValue<bool>("ols_sendsms") && funeral != null)
            {
                SendNotification(funeral);
            }
            #endregion
        }

        public void SendNotification(Entity funeral)
        {
            var smsres = "";
            var emailres = "";

            SetSMSDetails(funeral);

            switch (ContactMethod)
            {
                case 3: emailres = SendEmail(funeral); break;
                case 0: smsres = SendSMS(); emailres = SendEmail(funeral); break;
                case 1: smsres = SendSMS(); break;
                case 2: break;
                default: smsres = SendSMS(); emailres = SendEmail(funeral); break;
            }

            //UpdateSMSStatus(smsres, emailres);
        }

        public string Send()
        {
            return smsData.Send();
        }

        public Guid GetDoctorIdByName(string name)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_medicalreferee");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_name", ConditionOperator.Equal, name);
                Entity medicalReferee = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                if (medicalReferee != null)
                    return medicalReferee.Id;
                else
                    return Guid.Empty;
            }
            catch { return Guid.Empty; }
        }

        public void RequestStatus(Entity postImage)
        {
            _config = new ConfigData(GetService(UserType.User));

            objBDMFuneralData.Applicationid = postImage.Contains("ols_applicationid") ? postImage.GetAttributeValue<string>("ols_applicationid") : string.Empty;
            objBDMFuneralData.Notificationid = postImage.Contains("ols_notificationid") ? postImage.GetAttributeValue<string>("ols_notificationid") : string.Empty;

            Status(objBDMFuneralData);
            UpdateRequestStatus(postImage);
        }
        public void UpdateRequestStatus(Entity postImage)
        {
            Entity ent = new Entity(postImage.LogicalName, postImage.Id);
            ent["ols_bdmerror"] = objBDMFuneralData.Error;
            ent["ols_requestdate"] = DateTime.UtcNow;
            ent["ols_sendstatusrequest"] = false;
            if (objBDMFuneralData.Notificationstatus != 0)
                ent["ols_notificationstatus"] = new OptionSetValue(objBDMFuneralData.Notificationstatus);
            ent["ols_applicationstatus"] = objBDMFuneralData.Applicationstatus;

            try
            {
                Update(UserType.User, ent);
            }
            catch (Exception ex) { AppendLog(ex.Message); throw ex; }
        }
        public void Status(BDMFuneralData fd)
        {
            if (fd.Notificationid == string.Empty)
            {
                fd.Error = "Notification Id is blank.";
                return;
            }
            try
            {

                var proxy = GetStatusQueryProxy();

                if (proxy == null) return;

                InspectedSOAPMessages soapMessages = new InspectedSOAPMessages();

                proxy.Endpoint.Behaviors.Add(new Common.MBSGuru.CapturingEndpointBehavior(soapMessages));

                fd = BDMApplicationRequest(proxy, soapMessages, fd);

                fd = BDMNotificationRequest(proxy, soapMessages, fd);

                proxy.Close();


            }
            catch (Exception global_ex)
            {

                fd.Error = global_ex.Message;


            }
        }
        private BDMFuneralData BDMNotificationRequest(StatusQueryPortTypeClient proxy, InspectedSOAPMessages soapMessages, BDMFuneralData fd)
        {

            try
            {

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (se, cert, chain, sslerror) => { return true; };
                StatusQueryRequestType request = new StatusQueryRequestType();

                request.NotificationID = fd.Notificationid;
                var response = proxy.StatusQuery(request);


                //response will be null if the contract was violated . . . 
                if (response == null)
                {
                    var raw_result = soapMessages.Response;
                    //fd.new_notificationid = fd.NotificationId;
                    fd.Notificationstatus = GetNotificationStatusValue(GetNotificationStatusFromRequest(raw_result));

                }
                else
                {

                    fd.Notificationid = response.NotificationID;
                    fd.Notificationstatus = GetNotificationStatusValue(response.Status);
                }
                fd.Error = "";

                return fd;

            }
            catch (Exception ex)
            {

                fd.Error = "Notification Status: " + ex.Message;
                return fd;
            }

        }
        private BDMFuneralData BDMApplicationRequest(StatusQueryPortTypeClient proxy, InspectedSOAPMessages soapMessages, BDMFuneralData fd)
        {

            try
            {
                System.Net.ServicePointManager.ServerCertificateValidationCallback += (se, cert, chain, sslerror) => { return true; };
                StatusQueryRequestType request = new StatusQueryRequestType();

                if (string.IsNullOrEmpty(fd.Applicationid)) return fd;
                request.NotificationID = fd.Applicationid;

                var response = proxy.StatusQuery(request);

                if (response == null)
                {
                    var raw_result = soapMessages.Response;

                    fd.Applicationstatus = GetNotificationStatusFromRequest(raw_result);

                    //fd.new_applicationid = fd.ApplicationId;
                }
                else
                {

                    fd.Applicationid = response.NotificationID;
                    fd.Applicationstatus = response.Status;
                }

                fd.Error = "";

                return fd;

            }
            catch (Exception ex)
            {
                fd.Error = "Application Status: " + ex.Message;
                return fd;
            }

        }
        private StatusQueryPortTypeClient GetStatusQueryProxy()
        {
            try
            {
                // ConfigData cdata = new ConfigData(_svc);
                System.ServiceModel.EndpointAddress endpointAddress = new System.ServiceModel.EndpointAddress(_config.StatusQueryEndPointAddress);

                System.ServiceModel.BasicHttpBinding binding = new System.ServiceModel.BasicHttpBinding();
                binding.Security.Mode = System.ServiceModel.BasicHttpSecurityMode.TransportWithMessageCredential;

                binding.Security.Message.ClientCredentialType = System.ServiceModel.BasicHttpMessageCredentialType.UserName;
                binding.Name = "StatusQueryBinding1";


                // LodgeDRSPortTypeClient proxy = new LodgeDRSPortTypeClient("LodgeDRSPort", endpointAddress);
                StatusQueryPortTypeClient proxy = new StatusQueryPortTypeClient(binding, endpointAddress);
                proxy.Endpoint.Binding.SendTimeout = new TimeSpan(0, 10, 00);
                proxy.Endpoint.Binding.ReceiveTimeout = new TimeSpan(0, 10, 00);


                // var factory = new System.ServiceModel.ChannelFactory<LodgeDRSPortTypeClient>(binding, endpointAddress);
                //LodgeDRSPortTypeClient proxy = factory.CreateChannel();

                proxy.ClientCredentials.UserName.UserName = _config.BDMRegAccount;
                proxy.ClientCredentials.UserName.Password = _config.BDMRegPassword;

                System.ServiceModel.Channels.BindingElementCollection elements = proxy.Endpoint.Binding.CreateBindingElements();
                elements.Find<System.ServiceModel.Channels.SecurityBindingElement>().EnableUnsecuredResponse = true;
                proxy.Endpoint.Binding = new System.ServiceModel.Channels.CustomBinding(elements);

                return proxy;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public void BDMExecute(Entity funeral, Entity postImage, Guid CurrentUserId)
        {
            _config = new ConfigData(GetService(UserType.User));

            try
            {
                try
                {
                    AppendLog("Build XML Started..");
                    BuildXML(funeral, postImage);
                    AppendLog("XML was build successfully");
                    //using (System.IO.StreamWriter sw = new System.IO.StreamWriter("c:\\temp\\log.txt", true))
                    //{
                    //    sw.WriteLine("XML was build successfully");
                    //}

                }
                catch (Exception ex)
                {
                    AppendLog("BuildXML Error: " + ex.Message);
                    objBDMFuneralData.Error = "BuildXML Error: " + ex.Message;
                    //BDMRegData.Error += "BuildXML Error: " + ex.Message;
                    //using (System.IO.StreamWriter sw = new System.IO.StreamWriter("c:\\temp\\log.txt", true))
                    //{
                    //    sw.WriteLine(BDMRegData.Error);
                    //}
                }

                try
                {
                    Submit(CurrentUserId);
                }
                catch (Exception ex) { objBDMFuneralData.Error = "Submit Error: " + ex.Message; /*BDMRegData.Error += "Submit Error: " + ex.Message;*/ }
            }
            catch { }
            finally { UpdateStatus(postImage.Id); }
        }

        public void UpdateStatus(Guid BDMId)
        {
            Entity ent = new Entity("ols_bdm", BDMId);
            ent["ols_sendtobdm"] = false;

            if (DRSRequest != null)
            {
                ent["ols_bdmstatus"] = new OptionSetValue(objBDMFuneralData.BDM_Status);
                ent["ols_bdmerror"] = objBDMFuneralData.Error;
                ent["ols_submissiondate"] = DateTime.UtcNow;
                ent["ols_xmlcontainerid"] = new EntityReference("ols_xmlcontainer", objBDMFuneralData.XMLId);
                ent["ols_validationerror"] = objBDMFuneralData.ValidationError;

                if (objBDMFuneralData.BDM_Status == (int)BDMStatus.Submitted)
                {
                    if (!string.IsNullOrEmpty(objBDMFuneralData.Notificationid))
                        ent["ols_notificationid"] = objBDMFuneralData.Notificationid;
                    //ent["new_notificationmessage"] = objBDMFuneralData.Notificationmessage;
                    ent["ols_notificationstatus"] = new OptionSetValue(objBDMFuneralData.Notificationstatus);
                    if (!string.IsNullOrEmpty(objBDMFuneralData.Applicationid))
                        ent["ols_applicationid"] = objBDMFuneralData.Applicationid;
                    //ent.Attributes.Add("new_applicationmessage", this.new_applicationmessage);
                    ent["ols_applicationstatus"] = objBDMFuneralData.Applicationstatus;
                }
            }
            try
            {
                Update(UserType.User, ent);
            }
            catch (Exception ex) { throw ex; }
        }
        public void Submit(Guid currentUserId)
        {
            try
            {
                AppendLog("Submit Started.");

                System.Net.ServicePointManager.ServerCertificateValidationCallback += (se, cert, chain, sslerror) => { return true; };
                AppendLog("Creating proxy.");
                var proxy = GetLodgeProxy(currentUserId);
                if (proxy == null || DRSRequest == null) return;
                AppendLog("Proxy created.");

                var soapMessages = new InspectedSOAPMessages();
                proxy.Endpoint.Behaviors.Add(new Common.MBSGuru.CapturingEndpointBehavior(soapMessages));

                try
                {
                    AppendLog("Calling LodgeDRS.");
                    var response = proxy.LodgeDRS(DRSRequest);
                    proxy.Close();
                    //response will be null if the contract was violated . . . 
                    if (response == null)
                    {
                        AppendLog("LodgeDRS response is null.");

                        var raw_result = soapMessages.Response;

                        objBDMFuneralData.Notificationid = GetNotificationId(raw_result);
                        // fd.new_notificationmessage = response1.Notification.Message;
                        objBDMFuneralData.Notificationstatus = GetNotificationStatusValue(GetNotificationStatus(raw_result));
                        objBDMFuneralData.Applicationid = GetApplicationId(raw_result);
                        // fd.new_applicationmessage = response.Application.Message;
                        objBDMFuneralData.Applicationstatus = GetApplicationStatus(raw_result);
                    }
                    else
                    {
                        AppendLog("LodgeDRS response is received.");

                        objBDMFuneralData.Notificationid = response.Notification.Id;
                        objBDMFuneralData.Notificationmessage = response.Notification.Message;
                        objBDMFuneralData.Notificationstatus = GetNotificationStatusValue(response.Notification.Status);
                        objBDMFuneralData.Applicationid = response.Application.Id;
                        objBDMFuneralData.Applicationmessage = response.Application.Message;
                        objBDMFuneralData.Applicationstatus = response.Application.Status;
                        AppendLog("LodgeDRS response Application msg: " + response.Application.Message);
                        AppendLog("LodgeDRS response Notification msg: " + response.Notification.Message);


                    }


                    objBDMFuneralData.BDM_Status = Convert.ToInt32(BDMStatus.Submitted);
                    objBDMFuneralData.Error = "";
                    objBDMFuneralData.ValidationError = "";
                }
                catch (System.ServiceModel.FaultException ex)
                {
                    proxy.Close();
                    var raw_result = soapMessages.Response;

                    objBDMFuneralData.BDM_Status = Convert.ToInt32(BDMStatus.ValidationError);
                    objBDMFuneralData.Error = ex.Message;
                    var errorsArray = GetErrorDetails(raw_result);

                    if (errorsArray != null && errorsArray.Count > 0)
                    {
                        System.Text.StringBuilder errors = new System.Text.StringBuilder();
                        for (var i = 0; i < errorsArray.Count; i++)
                        {
                            errors.AppendLine(errorsArray[i]);
                        }
                        objBDMFuneralData.ValidationError = errors.ToString();
                        AppendLog("proxy.LodgeDRS Catch Error: " + errors.ToString());
                    }
                    else
                    {
                        objBDMFuneralData.ValidationError = raw_result;
                        AppendLog("proxy.LodgeDRS Catch Error: " + raw_result);

                    }

                }
            }
            catch (Exception global_ex)
            {
                AppendLog(global_ex.Message);
                objBDMFuneralData.BDM_Status = Convert.ToInt32(BDMStatus.RegistrationError);
                objBDMFuneralData.Error = global_ex.Message;
                objBDMFuneralData.ValidationError = "";

            }
        }
        private int GetNotificationStatusValue(string status)
        {
            if (!string.IsNullOrEmpty(status))
            {
                switch (status)
                {
                    case "Non Compliant": return 1;
                    case "Compliant": return 2;
                    case "Cancelled": return 3;
                    case "Suspended": return 4;
                    case "Registered": return 5;
                    case "Complete": return 6;
                }
            }
            return -1;
        }
        private LodgeDRSPortTypeClient GetLodgeProxy(Guid currentUserId)
        {
            try
            {
                //  ConfigData cdata = new ConfigData(svc);
                System.ServiceModel.EndpointAddress endpointAddress = new System.ServiceModel.EndpointAddress(_config.EndPointAddress);
                // System.ServiceModel.Channels.Binding binding = new System.ServiceModel.BasicHttpBinding(System.ServiceModel.BasicHttpSecurityMode.TransportWithMessageCredential);
                System.ServiceModel.BasicHttpBinding binding = new System.ServiceModel.BasicHttpBinding();
                binding.Security.Mode = System.ServiceModel.BasicHttpSecurityMode.TransportWithMessageCredential;
                // binding.Security.Transport.ClientCredentialType = System.ServiceModel.HttpClientCredentialType.Basic;
                binding.Security.Message.ClientCredentialType = System.ServiceModel.BasicHttpMessageCredentialType.UserName;
                binding.Name = "LodgeDRSBinding";



                // LodgeDRSPortTypeClient proxy = new LodgeDRSPortTypeClient("LodgeDRSPort", endpointAddress);
                LodgeDRSPortTypeClient proxy = new LodgeDRSPortTypeClient(binding, endpointAddress);
                proxy.Endpoint.Binding.SendTimeout = new TimeSpan(0, 10, 00);
                proxy.Endpoint.Binding.ReceiveTimeout = new TimeSpan(0, 10, 00);


                // var factory = new System.ServiceModel.ChannelFactory<LodgeDRSPortTypeClient>(binding, endpointAddress);
                //LodgeDRSPortTypeClient proxy = factory.CreateChannel();

                if (CheckBusinessUnit(currentUserId) == "Walter Carter Funerals Pty Ltd")
                {
                    proxy.ClientCredentials.UserName.UserName = _config.BDMRegWalterCarterAccount;
                    proxy.ClientCredentials.UserName.Password = _config.BDMRegWalterCarterPassword;
                }
                else
                {
                    proxy.ClientCredentials.UserName.UserName = _config.BDMRegAccount;
                    proxy.ClientCredentials.UserName.Password = _config.BDMRegPassword;
                }

                System.ServiceModel.Channels.BindingElementCollection elements = proxy.Endpoint.Binding.CreateBindingElements();
                elements.Find<System.ServiceModel.Channels.SecurityBindingElement>().EnableUnsecuredResponse = true;
                proxy.Endpoint.Binding = new System.ServiceModel.Channels.CustomBinding(elements);

                return proxy;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private string CheckBusinessUnit(Guid currentUserid)
        {
            Entity user = Retrieve(UserType.User, "systemuser", currentUserid, new ColumnSet("businessunitid"));
            if (user != null)
            {
                return user.Contains("businessunitid") ? user.GetAttributeValue<EntityReference>("businessunitid").Name : string.Empty;
            }
            else
                return string.Empty;
        }
        public void BuildXML(Entity funeral, Entity postImage)
        {
            try
            {
                FillDRSReq(funeral, postImage);
                Serialize();
                SaveXML();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void Serialize()//getting xml from DRSReq object
        {
            try
            {
                AppendLog("Serialize() Method started..");
                XmlSerializer xsSubmit = new XmlSerializer(typeof(DRSRequestType));
                XmlDocument doc = new XmlDocument();
                System.IO.StringWriter sww = new System.IO.StringWriter();
                XmlWriter writer = XmlWriter.Create(sww);
                xsSubmit.Serialize(writer, DRSRequest);
                var xml = sww.ToString();
                doc.LoadXml(xml);
                XMLData = doc;
                AppendLog("Serialize() Method completed..");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void SaveXML()
        {
            AppendLog("SaveXML() started..");
            if (xmlContainerId == null)
                CreateNewXMLContainer();
            else EditXmlContainer();
            AppendLog("SaveXML() completed.");
        }
        private void CreateNewXMLContainer()
        {
            AppendLog("CreateNewXMLContainer() started..");
            Guid XMLId = Guid.Empty;
            try
            {
                Entity xmlEntity = new Entity("ols_xmlcontainer");
                //xmlEntity.Attributes.Add("new_xml", XMLData.InnerXml);
                xmlEntity["ols_funeralnumber"] = funeralNumber;
                xmlEntity["ols_name"] = funeralNumber;

                XMLId = Create(UserType.User, xmlEntity);
                objBDMFuneralData.XMLId = XMLId;
            }
            catch (Exception ex) { }

            if (XMLId != null)
            {
                try
                {
                    CreateAnnotation(XMLId);
                }
                catch (Exception ex) { throw ex; }
            }
            AppendLog("CreateNewXMLContainer() completed.");
        }
        private void CreateAnnotation(Guid XMLId)
        {
            try
            {
                AppendLog("CreateAnnotation method started..");
                Entity attachment = new Entity("annotation");
                attachment.Attributes.Add("subject", "File Attachment");
                attachment.Attributes.Add("filename", funeralNumber + ".xml");

                byte[] bytes = new byte[XMLData.InnerXml.Length * sizeof(char)];
                System.Buffer.BlockCopy(XMLData.InnerXml.ToCharArray(), 0, bytes, 0, bytes.Length);
                attachment.Attributes.Add("documentbody", Convert.ToBase64String(bytes));
                attachment.Attributes.Add("objectid", new EntityReference("ols_xmlcontainer", XMLId));
                attachment.Attributes.Add("objecttypecode", "ols_xmlcontainer");
                attachment.Attributes.Add("mimetype", @"text/plain");
                Create(UserType.User, attachment);
                AppendLog("CreateAnnotation method completed.");
            }
            catch (Exception ex) { throw ex; }
        }

        private void EditXmlContainer()
        {
            AppendLog("EditXmlContainer() started..");
            if (xmlContainerId != null)
            {
                DeleteAnnotation(xmlContainerId.Id);
                CreateAnnotation(xmlContainerId.Id);
            }
            AppendLog("EditXmlContainer() completed.");
        }
        private void DeleteAnnotation(Guid xmlId)
        {
            try
            {
                AppendLog("DeleteAnnotation method started..");
                QueryExpression qe = new QueryExpression("annotation");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("objectid", ConditionOperator.Equal, xmlId);
                Entity annotation = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                if (annotation != null)
                {
                    Delete(UserType.User, "annotation", annotation.Id);
                }
                AppendLog("DeleteAnnotation method completed.");
            }
            catch { }
        }
        public void FillDRSReq(Entity funeral, Entity postImage)
        {
            AppendLog("FillDRSReq Started..");
            var DRSReq = new DRSRequestType();

            AppendLog("Getting CertificateRequest");
            DRSReq.CertificateRequest = GetCertificateRequest(funeral, postImage);

            var DRSDetails = new DRSDetailsType();

            AppendLog("Getting DeceasedDetails");
            //DRSDetails.AdditionalDetails = new DBM.AdditionalDetails().;
            DRSDetails.DeceasedDetails = GetDeceasedDetails(funeral);

            AppendLog("Getting ChildrenOfDeceased");
            DRSDetails.ChildrenOfDeceased = GetChildren(funeral.Id);

            AppendLog("Getting ParentsOfDeceased");
            DRSDetails.ParentsOfDeceased = GetParentsDetails(funeral);

            AppendLog("Getting DisposalDetails");
            DRSDetails.DisposalDetails = GetDisposalDetails(funeral);

            AppendLog("Getting FuneralDirectorDetails");
            DRSDetails.FuneralDirectorDetails = GetFuneralDirectorDetails(funeral);

            AppendLog("Getting InformantDetails");
            DRSDetails.InformantDetails = GetInformantDetails(funeral);

            AppendLog("Getting MarriageDetails");
            EntityCollection MarriageHistoryColl = GetMarriageHistoryEntity(funeral.Id);
            if (MarriageHistoryColl != null && MarriageHistoryColl.Entities.Count > 0)
            {
                DRSDetails.MarriageDetails = GetMarriageDetails(MarriageHistoryColl.Entities.FirstOrDefault());
                DRSDetails.PreviousMarriageDetails = GetMarriageHistoryDetails(MarriageHistoryColl);
            }

            AppendLog("Getting DeathCertificate");
            Entity deathCertificate = null;
            EntityReference deathCertificateRef = funeral.Contains("ols_deathcertificateid") ? funeral.GetAttributeValue<EntityReference>("ols_deathcertificateid") : null;
            if (deathCertificateRef != null)
                deathCertificate = Retrieve(UserType.User, deathCertificateRef.LogicalName, deathCertificateRef.Id, new ColumnSet(true));

            if (deathCertificate != null && deathCertificate.Contains("ols_typeofdeathcertificate"))
            {
                DRSDetails.TypeOfDeathCertificate = GetTypeOfDeathCertificateDetails(deathCertificate.GetAttributeValue<OptionSetValue>("ols_typeofdeathcertificate").Value);
                DRSDetails.TypeOfDeathCertificateSpecified = true;
            }

            DRSReq.DRSDetails = DRSDetails;

            DRSRequest = DRSReq;
            AppendLog("FillDRSReq completed..");
        }

        private ApplicationDetails GetCertificateRequest(Entity funeral, Entity postImage)
        {
            var appDetails = new ApplicationDetails();


            //  appDetail. = getDeceasedDetails4App(fd);

            appDetails.Item = GetDeceasedDetails4App(funeral);

            appDetails.Product = new ProductDetailsType();
            appDetails.Product.ProductCode = "eNDO";
            appDetails.Product.Quantity = postImage.Contains("ols_copies") ? postImage.GetAttributeValue<int>("ols_copies") : 0;
            appDetails.Product.Specification = "DR Certificate";

            appDetails.Delivery = new DeliveryDetailsType();

            #region Get Brand Details
            EntityReference brandRef = funeral.Contains("pricelevelid") ? funeral.GetAttributeValue<EntityReference>("pricelevelid") : null;
            if (brandRef != null)
                brand = Retrieve(UserType.User, brandRef.LogicalName, brandRef.Id, new ColumnSet(true));
            #endregion

            if (brand != null)
            {
                appDetails.Delivery.CompanyName = brand.Contains("ols_companyname") ? brand.GetAttributeValue<string>("ols_companyname") : string.Empty;

                #region Address
                appDetails.Delivery.DeliveryAddress = new AddressType();
                if (!brand.Contains("ols_countryid") || (brand.Contains("ols_countryid") && brand.GetAttributeValue<EntityReference>("ols_countryid").Name == "Australia"))
                {
                    appDetails.Delivery.DeliveryAddress.Line1 = brand.Contains("ols_street1") ? brand.GetAttributeValue<string>("ols_street1") : string.Empty;

                    appDetails.Delivery.DeliveryAddress.Line2 = brand.Contains("ols_street2") ? brand.GetAttributeValue<string>("ols_street2") : string.Empty;
                    if (brand.Contains("ols_postcode"))
                    {
                        appDetails.Delivery.DeliveryAddress.Postcode = Convert.ToInt16(brand.GetAttributeValue<string>("ols_postcode"));
                        appDetails.Delivery.DeliveryAddress.PostcodeSpecified = true;
                    }
                    appDetails.Delivery.DeliveryAddress.Country = "Australia";
                    appDetails.Delivery.DeliveryAddress.SuburbOrTown = brand.Contains("ols_suburbid") ? brand.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                    if (brand.Contains("ols_stateid"))
                    {
                        appDetails.Delivery.DeliveryAddress.State = GetState(brand.GetAttributeValue<EntityReference>("ols_stateid").Name);
                        appDetails.Delivery.DeliveryAddress.StateSpecified = true;
                    }
                }
                else
                {
                    appDetails.Delivery.DeliveryAddress.Country = brand.Contains("ols_countryid") ? brand.GetAttributeValue<EntityReference>("ols_countryid").Name : string.Empty;
                    appDetails.Delivery.DeliveryAddress.InternationalAddress = brand.Contains("ols_internationaladdress") ? brand.GetAttributeValue<string>("ols_internationaladdress") : string.Empty;
                }

                #endregion

                appDetails.Delivery.DeliveryName = new NameType();
                appDetails.Delivery.DeliveryName.FamilyName = brand.Contains("ols_familyname") ? brand.GetAttributeValue<string>("ols_familyname") : string.Empty;
                appDetails.Delivery.DeliveryName.FirstGivenName = brand.Contains("ols_firstgivenname") ? brand.GetAttributeValue<string>("ols_firstgivenname") : string.Empty;
                appDetails.Delivery.DeliveryName.OtherGivenNames = brand.Contains("ols_othergivennames") ? brand.GetAttributeValue<string>("ols_othergivennames") : string.Empty;

                // appDetails.Delivery.FaxToPassportOffice = true;
                // appDetails.Delivery.FaxToPassportOfficeSpecified = true;
                //  appDetails.Extras = getProductDetails();
                // appDetails.Item = new object();
                appDetails.PaymentBy = ApplicationDetailsPaymentBy.Account;
                appDetails.PaymentBySpecified = true;
                appDetails.Stakeholder = brand.Contains("ols_stakeholdername") ? brand.GetAttributeValue<string>("ols_stakeholdername") : string.Empty;
                // appDetails.Subproducts = getSubProductsDetails();
            }
            appDetails.Delivery.ExternalRefNumber = funeral.Contains("ols_funeralnumber") ? funeral.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
            if (postImage.Contains("ols_tobepostedcollected"))
                appDetails.Delivery.DeliveryMethod = GetDeliveryMethod(postImage.GetAttributeValue<OptionSetValue>("ols_tobepostedcollected").Value);

            return appDetails;
        }

        private DeceasedDetailsType GetDeceasedDetails4App(Entity funeral)
        {
            var deceasedDetail = new DeceasedDetailsType();

            #region Name
            var nameType = new ExtendedNameType();
            nameType.FamilyNameAtBirth = funeral.Contains("ols_deceasedfamilynameatbirth") ? funeral.GetAttributeValue<string>("ols_deceasedfamilynameatbirth") : string.Empty;
            nameType.Name = new NameType();
            nameType.Name.FamilyName = funeral.Contains("ols_deceasedfamilyname") ? funeral.GetAttributeValue<string>("ols_deceasedfamilyname") : string.Empty;
            nameType.Name.OtherGivenNames = funeral.Contains("ols_deceasedothergivennames") ? funeral.GetAttributeValue<string>("ols_deceasedothergivennames") : string.Empty;
            nameType.Name.FirstGivenName = funeral.Contains("ols_deceasedgivennames") ? funeral.GetAttributeValue<string>("ols_deceasedgivennames") : string.Empty;
            deceasedDetail.DeceasedName = nameType;
            // deceasedDetail.FamilyNameAtBirth = fd.new_deceasedfamilynameatbirth;

            #endregion

            #region DateOfDeath
            if (!funeral.Contains("ols_dateofdeathtype") || (funeral.Contains("ols_dateofdeathtype") && funeral.GetAttributeValue<OptionSetValue>("ols_dateofdeathtype").Value != 2))
            {
                if (funeral.Contains("ols_dateofdeathtype"))
                    deceasedDetail.DateOfDeathType = GetDateOfDeathType(funeral.GetAttributeValue<OptionSetValue>("ols_dateofdeathtype").Value);
                if (funeral.Contains("ols_dateofdeath"))
                    deceasedDetail.DateOfDeathExact = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdeath"));
                deceasedDetail.DateOfDeathExactSpecified = true;
            }
            else
            {
                deceasedDetail.DateOfDeathType = DateRangeBasis.Between;
                if (funeral.Contains("ols_dateofdeathfrom"))
                    deceasedDetail.DateOfDeathFrom = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdeathfrom"));
                deceasedDetail.DateOfDeathFromSpecified = true;
                if (funeral.Contains("ols_dateofdeathto"))
                    deceasedDetail.DateOfDeathTo = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdeathto"));
                deceasedDetail.DateOfDeathToSpecified = true;
            }
            #endregion

            return deceasedDetail;
        }

        private DateRangeBasis GetDateOfDeathType(int type)
        {
            switch (type)
            {
                case 1: return DateRangeBasis.Approx;
                case 2: return DateRangeBasis.Between;
                case 3: return DateRangeBasis.On;
                case 4: return DateRangeBasis.OnorAbout;
                case 5: return DateRangeBasis.OnorAfter;
                case 6: return DateRangeBasis.SometimeOnorAfter;
                case 7: return DateRangeBasis.Unknown;
                default: return DateRangeBasis.On;
            }
        }
        private string GetDeliveryMethod(int type)
        {
            switch (type)
            {
                case -1: return "";
                case 1: return "Registered Mail";
                case 2: return "Collect";

            }
            return "";
        }
        private AddressTypeState GetState(string state)
        {
            if (string.IsNullOrEmpty(state)) return AddressTypeState.Unknown;
            state = state.ToUpper();
            switch (state)
            {
                case "NSW": return AddressTypeState.NSW;
                case "ACT": return AddressTypeState.ACT;
                case "NT": return AddressTypeState.NT;
                case "QLD": return AddressTypeState.QLD;
                case "SA": return AddressTypeState.SA;
                case "TAS": return AddressTypeState.TAS;
                case "VIC": return AddressTypeState.VIC;
                case "WA": return AddressTypeState.WA;
                default: return AddressTypeState.Unknown;
            }
        }

        private DeceasedDetails GetDeceasedDetails(Entity funeral)
        {
            var deceasedDetail = new DeceasedDetails();

            #region Name
            var nameType = new NameType();
            nameType.FamilyName = funeral.Contains("ols_deceasedfamilyname") ? funeral.GetAttributeValue<string>("ols_deceasedfamilyname").ToUpper() : string.Empty;
            nameType.OtherGivenNames = funeral.Contains("ols_deceasedothergivennames") ? funeral.GetAttributeValue<string>("ols_deceasedothergivennames") : string.Empty;
            nameType.FirstGivenName = funeral.Contains("ols_deceasedgivennames") ? funeral.GetAttributeValue<string>("ols_deceasedgivennames") : string.Empty;
            deceasedDetail.DeceasedName = nameType;
            deceasedDetail.FamilyNameAtBirth = funeral.Contains("ols_deceasedfamilynameatbirth") ? funeral.GetAttributeValue<string>("ols_deceasedfamilynameatbirth").ToUpper() : string.Empty;
            #endregion

            #region Age
            if (funeral.Contains("ols_ageatdod"))
            {
                deceasedDetail.Age = int.Parse(Regex.Match(funeral.GetAttributeValue<string>("ols_ageatdod"), @"^\d+").Value);
                deceasedDetail.AgeSpecified = true;
                if (funeral.Contains("ols_ageunit"))
                {
                    deceasedDetail.AgeUnit = GetAgeUnit(funeral.GetAttributeValue<OptionSetValue>("ols_ageunit").Value);
                    deceasedDetail.AgeUnitSpecified = true;
                }
            }
            #endregion

            #region Aboriginality

            if (funeral.Contains("ols_yearofarrivalinaustralia") && ((funeral.Contains("ols_placeofbirthcountryid") && funeral.GetAttributeValue<EntityReference>("ols_placeofbirthcountryid").Name != "Australia") || !funeral.Contains("ols_placeofbirthcountryid")))
            {
                try
                {
                    deceasedDetail.DateOfArrivalToAus = new DateTime(Convert.ToInt32(funeral.GetAttributeValue<string>("ols_yearofarrivalinaustralia")), 1, 1);
                    deceasedDetail.DateOfArrivalToAusSpecified = true;
                }
                catch { }
            }
            if (funeral.Contains("ols_aboriginalortorresstraitislanderorigin"))
            {
                deceasedDetail.Aboriginality = GetAboriginalityStatus(funeral.GetAttributeValue<OptionSetValue>("ols_aboriginalortorresstraitislanderorigin").Value);
                deceasedDetail.AboriginalitySpecified = true;
            }


            #endregion

            #region DateOfDeath
            if (funeral.Contains("ols_deceaseddob"))
            {
                deceasedDetail.DateOfBirth = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_deceaseddob"));
                deceasedDetail.DateOfBirthSpecified = true;
            }
            if (!funeral.Contains("ols_dateofdeathtype") || (funeral.Contains("ols_dateofdeathtype") && funeral.GetAttributeValue<OptionSetValue>("ols_dateofdeathtype").Value != 2))
            {
                if (funeral.Contains("ols_dateofdeathtype"))
                    deceasedDetail.DateOfDeathType = GetDateOfDeathType(funeral.GetAttributeValue<OptionSetValue>("ols_dateofdeathtype").Value);
                if (funeral.Contains("ols_dateofdeath"))
                {
                    deceasedDetail.DateOfDeathExact = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdeath"));
                    deceasedDetail.DateOfDeathExactSpecified = true;
                }
            }
            else
            {
                AppendLog("DeceasedDetails DateOfDeathType is Between");
                deceasedDetail.DateOfDeathType = DateRangeBasis.Between;
                if (funeral.Contains("ols_dateofdeathfrom"))
                {
                    deceasedDetail.DateOfDeathFrom = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdeathfrom"));
                    deceasedDetail.DateOfDeathFromSpecified = true;
                }
                if (funeral.Contains("ols_dateofdeathto"))
                {
                    deceasedDetail.DateOfDeathTo = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdeathto"));
                    deceasedDetail.DateOfDeathToSpecified = true;
                }
            }

            #endregion

            #region DeathAddress

            if (funeral.Contains("ols_transferlocationid"))
                deceasedDetail.DeathLocation = funeral.GetAttributeValue<EntityReference>("ols_transferlocationid").Name;
            else deceasedDetail.DeathLocation = funeral.Contains("ols_placeofdeath") ? funeral.GetAttributeValue<string>("ols_placeofdeath") : string.Empty;

            #region Get Transfer Location
            Entity transferLocation = null;
            if (funeral.Contains("ols_transferlocationid"))
                transferLocation = Retrieve(UserType.User, "ols_transferlocations", funeral.GetAttributeValue<EntityReference>("ols_transferlocationid").Id, new ColumnSet(true));

            #endregion

            if (transferLocation != null && transferLocation.Contains("ols_hospital") && transferLocation.GetAttributeValue<bool>("ols_hospital"))// Death in Hospital
            {
                deceasedDetail.DeathOccuredInHospital = true;
                deceasedDetail.DeathHospital = funeral.Contains("ols_transferlocationid") ? funeral.GetAttributeValue<EntityReference>("ols_transferlocationid").Name : string.Empty;
                deceasedDetail.DeathHospitalSuburb = funeral.Contains("ols_placeofdeathcityid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofdeathcityid").Name : string.Empty;
            }
            else
            {
                var deathAddress = new AddressType();
                if (!funeral.Contains("ols_placeofdeathcountryid") || (funeral.Contains("ols_placeofdeathcountryid") && funeral.GetAttributeValue<EntityReference>("ols_placeofdeathcountryid").Name == "Australia"))
                {
                    deathAddress.Line1 = funeral.Contains("ols_placeofdeathstreet") ? funeral.GetAttributeValue<string>("ols_placeofdeathstreet") : string.Empty;
                    deathAddress.Line2 = funeral.Contains("ols_placeofdeathstreet2") ? funeral.GetAttributeValue<string>("ols_placeofdeathstreet2") : string.Empty;
                    if (funeral.Contains("ols_placeofdeathpostcodeid"))
                    {
                        deathAddress.Postcode = Convert.ToInt16(funeral.GetAttributeValue<EntityReference>("ols_placeofdeathpostcodeid").Name);
                        deathAddress.PostcodeSpecified = true;
                    }
                    deathAddress.Country = "Australia";
                    deathAddress.SuburbOrTown = funeral.Contains("ols_placeofdeathcityid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofdeathcityid").Name : string.Empty;
                    if (funeral.Contains("ols_placeofdeathstateid"))
                    {
                        deathAddress.State = GetState(funeral.GetAttributeValue<EntityReference>("ols_placeofdeathstateid").Name);
                        deathAddress.StateSpecified = true;
                    }
                }
                else
                {
                    deathAddress.Country = funeral.Contains("ols_placeofdeathcountryid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofdeathcountryid").Name : string.Empty;
                    deathAddress.InternationalAddress = funeral.Contains("ols_internationaladdress") ? funeral.GetAttributeValue<string>("ols_internationaladdress") : string.Empty;
                }
                deceasedDetail.DeathAddress = deathAddress;
            }

            #endregion

            #region Gender
            if (funeral.Contains("ols_deceasedgender"))
            {
                deceasedDetail.Gender = GetGender(funeral.GetAttributeValue<OptionSetValue>("ols_deceasedgender").Value);
                deceasedDetail.GenderSpecified = true;
            }
            #endregion

            #region Occupation
            var occupation = new OccupationType();
            occupation.MainOccupationActivity = funeral.Contains("ols_maintasksperformed") ? funeral.GetAttributeValue<string>("ols_maintasksperformed") : string.Empty;
            occupation.Occupation = funeral.Contains("ols_usualoccupationduringworkinglife") ? funeral.GetAttributeValue<string>("ols_usualoccupationduringworkinglife") : string.Empty;
            deceasedDetail.Occupation = occupation;
            #endregion

            #region Usual Residential Address
            var residentialAddress = new AddressType();
            if (!funeral.Contains("ols_usualresidenceofdeceasedcountryid") || (funeral.Contains("ols_usualresidenceofdeceasedcountryid") && funeral.GetAttributeValue<EntityReference>("ols_usualresidenceofdeceasedcountryid").Name == "Australia"))
            {
                residentialAddress.Line1 = funeral.Contains("ols_usualresidenceofdeceasedstreet1") ? funeral.GetAttributeValue<string>("ols_usualresidenceofdeceasedstreet1") : string.Empty;
                residentialAddress.Line2 = funeral.Contains("ols_usualresidenceofdeceasedstreet2") ? funeral.GetAttributeValue<string>("ols_usualresidenceofdeceasedstreet2") : string.Empty;
                if (funeral.Contains("ols_postcodeid"))
                {
                    residentialAddress.Postcode = Convert.ToInt16(funeral.GetAttributeValue<EntityReference>("ols_postcodeid").Name);
                    residentialAddress.PostcodeSpecified = true;
                }
                residentialAddress.Country = "Australia";
                residentialAddress.SuburbOrTown = funeral.Contains("ols_usualresidenceofcityid") ? funeral.GetAttributeValue<EntityReference>("ols_usualresidenceofcityid").Name : string.Empty;
                if (funeral.Contains("ols_usualresidenceofdeceasedstateid"))
                {
                    residentialAddress.State = GetState(funeral.GetAttributeValue<EntityReference>("ols_usualresidenceofdeceasedstateid").Name);
                    residentialAddress.StateSpecified = true;
                }

            }
            else
            {
                residentialAddress.Country = funeral.Contains("ols_usualresidenceofdeceasedcountryid") ? funeral.GetAttributeValue<EntityReference>("ols_usualresidenceofdeceasedcountryid").Name : string.Empty;
                residentialAddress.InternationalAddress = funeral.Contains("ols_usualresidenceofdeceasedinternationaladdr") ? funeral.GetAttributeValue<string>("ols_usualresidenceofdeceasedinternationaladdr") : string.Empty;
            }
            deceasedDetail.ResidentialAddress = residentialAddress;

            #endregion

            #region Pension

            deceasedDetail.PensionerAtTimeOfDeath = funeral.Contains("ols_pensioneratdateofdeath") ? funeral.GetAttributeValue<bool>("ols_pensioneratdateofdeath") : false;
            deceasedDetail.PensionerAtTimeOfDeathSpecified = true;
            if (funeral.Contains("ols_pensiontype"))
                deceasedDetail.PensionType = funeral.FormattedValues["ols_pensiontype"].ToString().ToUpper(); //EntityHelper.RetrieveAttributeMetadata(svc,"opportunity","new_pensiontype", fd.new_pensiontype);

            deceasedDetail.RetiredAtTimeOfDeath = funeral.Contains("ols_retiredatdateofdeath") ? funeral.GetAttributeValue<bool>("ols_retiredatdateofdeath") : false;
            deceasedDetail.RetiredAtTimeOfDeathSpecified = true;
            #endregion


            #region Place of Birth
            deceasedDetail.PlaceOfBirth = new AddressSuburbStateCountryType();

            if (!funeral.Contains("ols_placeofbirthcountryid") || (funeral.Contains("ols_placeofbirthcountryid") && funeral.GetAttributeValue<EntityReference>("ols_placeofbirthcountryid").Name == "Australia"))
            {
                deceasedDetail.PlaceOfBirth.Country = "Australia";
                deceasedDetail.PlaceOfBirth.SuburbOrTownOrCity = funeral.Contains("ols_placeofbirthsuburbid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofbirthsuburbid").Name : string.Empty;
                if (funeral.Contains("ols_placeofbirthstateid"))
                {
                    deceasedDetail.PlaceOfBirth.StateorTerritory = GetStateofBirth(funeral.GetAttributeValue<EntityReference>("ols_placeofbirthstateid").Name);
                    deceasedDetail.PlaceOfBirth.StateorTerritorySpecified = true;
                }
            }
            else
            {
                deceasedDetail.PlaceOfBirth.Country = funeral.Contains("ols_placeofbirthcountryid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofbirthcountryid").Name : string.Empty;
                deceasedDetail.PlaceOfBirth.SuburbOrTownOrCity = funeral.Contains("ols_placeofbirthsuburbid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofbirthsuburbid").Name : string.Empty;
            }
            #endregion



            return deceasedDetail;
        }

        private TimeUnits GetAgeUnit(int ageunit)
        {
            switch (ageunit)
            {
                case 1: return TimeUnits.days;
                case 2: return TimeUnits.hours;
                case 3: return TimeUnits.minutes;
                case 4: return TimeUnits.months;
                case 5: return TimeUnits.weeks;
                case 6: return TimeUnits.years;
                case 7: return TimeUnits.unknown;
                default: return TimeUnits.unknown;
            }
        }
        private AboriginalityType GetAboriginalityStatus(int status)
        {
            switch (status)
            {
                case 1: return AboriginalityType.BothAboriginalandTorresStraitIslander;
                case 2: return AboriginalityType.Aboriginal;
                case 3: return AboriginalityType.TorresStraitIslander;
                case 4: return AboriginalityType.Neither;
                default: return AboriginalityType.Neither;
            }
        }
        private GenderType GetGender(int gendertype)
        {
            switch (gendertype)
            {
                case 1: return GenderType.Male;
                case 2: return GenderType.Female;
                case 3: return GenderType.Indeterminate;
                case 4: return GenderType.Intersex;
                //  case 5: return GenderType.NotStated;
                case 6: return GenderType.Unknown;
                default: return GenderType.Unknown;
            }
        }
        private string GetPensionType(int type)
        {
            switch (type)
            {
                case -1: return "";
                case 2:
                case 22:
                    return "VETERAN";
                case 4: return "INVALID";
                case 5: return "AGED";
                case 6:
                case 7: return "WIDOW";
                case 19: return "UNKNOWN";
                case 20: return "NONE";
            }
            return "";
        }
        private AddressSuburbStateCountryTypeStateorTerritory GetStateofBirth(string state)
        {
            if (string.IsNullOrEmpty(state)) return AddressSuburbStateCountryTypeStateorTerritory.Unknown;
            state = state.ToUpper();
            switch (state)
            {
                case "NSW": return AddressSuburbStateCountryTypeStateorTerritory.NSW;
                case "ACT": return AddressSuburbStateCountryTypeStateorTerritory.ACT;
                case "NT": return AddressSuburbStateCountryTypeStateorTerritory.NT;
                case "QLD": return AddressSuburbStateCountryTypeStateorTerritory.QLD;
                case "SA": return AddressSuburbStateCountryTypeStateorTerritory.SA;
                case "TAS": return AddressSuburbStateCountryTypeStateorTerritory.TAS;
                case "VIC": return AddressSuburbStateCountryTypeStateorTerritory.VIC;
                case "WA": return AddressSuburbStateCountryTypeStateorTerritory.WA;
                default: return AddressSuburbStateCountryTypeStateorTerritory.Unknown;
            }
        }

        private ChildrenOfDeceased[] GetChildren(Guid OpportunityId)
        {
            try
            {
                var children = new List<ChildrenOfDeceased>();

                QueryExpression qe = new QueryExpression("ols_children");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, OpportunityId);
                EntityCollection ChildCollection = RetrieveMultiple(UserType.User, qe);

                if (ChildCollection != null && ChildCollection.Entities.Count > 0)
                {
                    foreach (var child in ChildCollection.Entities)
                    {
                        ChildrenOfDeceased child1 = new ChildrenOfDeceased();
                        if (child.Contains("ols_dateofbirth"))
                        {

                            child1.ChildDateOfBirth = Util.LocalFromUTCUserDateTime(GetService(UserType.User), child.GetAttributeValue<DateTime>("ols_dateofbirth"));



                            child1.ChildDateOfBirthSpecified = true;

                        }
                        var name = new NameType();

                        name.FirstGivenName = child.Contains("ols_firstgivenname") ? child.GetAttributeValue<string>("ols_firstgivenname") : string.Empty;

                        name.FamilyName = child.Contains("ols_familyname") ? child.GetAttributeValue<string>("ols_familyname").ToUpper() : string.Empty;

                        name.OtherGivenNames = child.Contains("ols_othergivennames") ? child.GetAttributeValue<string>("ols_othergivennames") : string.Empty;

                        child1.ChildName = name;

                        if (child.Contains("ols_sex"))
                        {
                            child1.ChildGender = GetGender(child.GetAttributeValue<OptionSetValue>("ols_sex").Value);
                            child1.ChildGenderSpecified = true;
                        }

                        if (child.Contains("ols_lifestatus"))
                        {
                            child1.LifeStatus = GetLifeStatus(child.GetAttributeValue<OptionSetValue>("ols_lifestatus").Value);
                            child1.LifeStatusSpecified = true;
                        }

                        if (child1.LifeStatus == ChildrenOfDeceasedLifeStatus.Alive)
                        {
                            if (child.Contains("ols_age"))
                            {
                                var age1 = child.GetAttributeValue<string>("ols_age");
                                child1.ChildAge = Convert.ToInt16(new String(age1.Where(x => x >= '0' && x <= '9').ToArray()));
                                child1.ChildAgeSpecified = true;
                            }
                            if (child.Contains("ols_ageunit"))
                            {
                                child1.ChildAgeUnit = GetAgeUnit(child.GetAttributeValue<OptionSetValue>("ols_ageunit").Value);
                                child1.ChildAgeUnitSpecified = true;
                            }
                        }

                        children.Add(child1);
                    }
                }

                return children.ToArray();
            }
            catch (Exception ex) { return null; }
        }
        private ChildrenOfDeceasedLifeStatus GetLifeStatus(int status)
        {
            switch (status)
            {
                case 1: return ChildrenOfDeceasedLifeStatus.Alive;

                case 2: return ChildrenOfDeceasedLifeStatus.Deceased;

                case 3: return ChildrenOfDeceasedLifeStatus.Stillborn;

                case 4: return ChildrenOfDeceasedLifeStatus.Unknown;

                default: return ChildrenOfDeceasedLifeStatus.Unknown;
            }
        }

        private ParentsOfDeceased GetParentsDetails(Entity funeral)
        {
            var parentsOfDeceased = new ParentsOfDeceased();

            var parentOneType = GetParentType(funeral.Contains("ols_parentonetype") ? funeral.GetAttributeValue<OptionSetValue>("ols_parentonetype").Value : 0);
            var parentTwoType = GetParentType(funeral.Contains("ols_parenttwotype") ? funeral.GetAttributeValue<OptionSetValue>("ols_parenttwotype").Value : 0);

            if (parentOneType != ParentRelationshipType.MOTHER && (parentTwoType == ParentRelationshipType.MOTHER || parentTwoType == ParentRelationshipType.PARENT))
            {
                SetParentOneFromTwo(ref parentsOfDeceased, funeral);
                parentsOfDeceased.ParentOneType = ParentRelationshipType.MOTHER;
                parentsOfDeceased.ParentOneTypeSpecified = true;
                parentsOfDeceased.ParentTwoType = parentOneType;
                parentsOfDeceased.ParentTwoTypeSpecified = true;

            }
            else
                if (parentOneType == ParentRelationshipType.MOTHER || parentOneType == ParentRelationshipType.PARENT)
            {
                parentsOfDeceased.ParentOneType = ParentRelationshipType.MOTHER;
                parentsOfDeceased.ParentOneTypeSpecified = true;
                SetParentOne(ref parentsOfDeceased, funeral);
                parentsOfDeceased.ParentTwoType = parentTwoType;
                parentsOfDeceased.ParentTwoTypeSpecified = true;
                SetParentTwo(ref parentsOfDeceased, funeral);
            }
            return parentsOfDeceased;
        }
        private ParentRelationshipType GetParentType(int parenttype)
        {
            switch (parenttype)
            {
                case 1: return ParentRelationshipType.FATHER;
                case 2: return ParentRelationshipType.MOTHER;
                case 3: return ParentRelationshipType.PARENT;
                default: return ParentRelationshipType.PARENT;
            }
        }

        private void SetParentOneFromTwo(ref ParentsOfDeceased parentsOfDeceased, Entity funeral)
        {

            if (funeral.Contains("ols_parentonegender"))
            {
                parentsOfDeceased.ParentTwoGender = GetGender(funeral.GetAttributeValue<OptionSetValue>("ols_parentonegender").Value);
                parentsOfDeceased.ParentTwoGenderSpecified = true;
            }
            var exname = new ExtendedNameType();
            exname.FamilyNameAtBirth = funeral.Contains("ols_parentonefamilynameatbirth") ? funeral.GetAttributeValue<string>("ols_parentonefamilynameatbirth").ToUpper() : string.Empty;
            var name = new NameType();
            name.FamilyName = funeral.Contains("ols_parentonefamilyname") ? funeral.GetAttributeValue<string>("ols_parentonefamilyname").ToUpper() : string.Empty;
            name.FirstGivenName = funeral.Contains("ols_parentonefirstgivenname") ? funeral.GetAttributeValue<string>("ols_parentonefirstgivenname") : string.Empty;
            name.OtherGivenNames = funeral.Contains("ols_parentoneothergivennames") ? funeral.GetAttributeValue<string>("ols_parentoneothergivennames") : string.Empty;
            exname.Name = name;
            parentsOfDeceased.ParentTwoName = exname;
            var occupation = new OccupationType();
            occupation.Occupation = funeral.Contains("ols_parentoneusualoccupation") ? funeral.GetAttributeValue<string>("ols_parentoneusualoccupation") : string.Empty;
            occupation.MainOccupationActivity = funeral.Contains("ols_parentonemaintasksperformed") ? funeral.GetAttributeValue<string>("ols_parentonemaintasksperformed") : string.Empty;
            parentsOfDeceased.ParentTwoOccupation = occupation;


            if (funeral.Contains("ols_parenttwogender"))
            {
                parentsOfDeceased.ParentOneGender = GetGender(funeral.GetAttributeValue<OptionSetValue>("ols_parenttwogender").Value);
                parentsOfDeceased.ParentOneGenderSpecified = true;
            }
            var exname2_2 = new ExtendedNameType();
            exname2_2.FamilyNameAtBirth = funeral.Contains("ols_parenttwofamilynameatbirth") ? funeral.GetAttributeValue<string>("ols_parenttwofamilynameatbirth").ToUpper() : string.Empty;
            var name2_2 = new NameType();
            name2_2.FamilyName = funeral.Contains("ols_parenttwofamilyname") ? funeral.GetAttributeValue<string>("ols_parenttwofamilyname").ToUpper() : string.Empty;
            name2_2.FirstGivenName = funeral.Contains("ols_parenttwofirstgivenname") ? funeral.GetAttributeValue<string>("ols_parenttwofirstgivenname") : string.Empty;
            name2_2.OtherGivenNames = funeral.Contains("ols_parenttwoothergivennames") ? funeral.GetAttributeValue<string>("ols_parenttwoothergivennames") : string.Empty;
            exname2_2.Name = name2_2;
            parentsOfDeceased.ParentOneName = exname2_2;
            var occupation2_2 = new OccupationType();
            occupation2_2.Occupation = funeral.Contains("ols_parenttwousualoccupation") ? funeral.GetAttributeValue<string>("ols_parenttwousualoccupation") : string.Empty;
            occupation2_2.MainOccupationActivity = funeral.Contains("ols_parenttwomaintasksperformed") ? funeral.GetAttributeValue<string>("ols_parenttwomaintasksperformed") : string.Empty;
            parentsOfDeceased.ParentOneOccupation = occupation2_2;

        }

        private void SetParentOne(ref ParentsOfDeceased parentsOfDeceased, Entity funeral)
        {
            if (funeral.Contains("ols_parentonegender"))
            {
                parentsOfDeceased.ParentOneGender = GetGender(funeral.GetAttributeValue<OptionSetValue>("ols_parentonegender").Value);
                parentsOfDeceased.ParentOneGenderSpecified = true;
            }
            var exname = new ExtendedNameType();
            exname.FamilyNameAtBirth = funeral.Contains("ols_parentonefamilynameatbirth") ? funeral.GetAttributeValue<string>("ols_parentonefamilynameatbirth").ToUpper() : string.Empty;
            var name = new NameType();
            name.FamilyName = funeral.Contains("ols_parentonefamilyname") ? funeral.GetAttributeValue<string>("ols_parentonefamilyname").ToUpper() : string.Empty;
            name.FirstGivenName = funeral.Contains("ols_parentonefirstgivenname") ? funeral.GetAttributeValue<string>("ols_parentonefirstgivenname") : string.Empty;
            name.OtherGivenNames = funeral.Contains("ols_parentoneothergivennames") ? funeral.GetAttributeValue<string>("ols_parentoneothergivennames") : string.Empty;
            exname.Name = name;
            parentsOfDeceased.ParentOneName = exname;
            var occupation = new OccupationType();
            occupation.Occupation = funeral.Contains("ols_parentoneusualoccupation") ? funeral.GetAttributeValue<string>("ols_parentoneusualoccupation") : string.Empty;
            occupation.MainOccupationActivity = funeral.Contains("ols_parentonemaintasksperformed") ? funeral.GetAttributeValue<string>("ols_parentonemaintasksperformed") : string.Empty;
            parentsOfDeceased.ParentOneOccupation = occupation;
        }

        private void SetParentTwo(ref ParentsOfDeceased parentsOfDeceased, Entity funeral)
        {

            if (funeral.Contains("ols_parenttwogender"))
            {
                parentsOfDeceased.ParentTwoGender = GetGender(funeral.GetAttributeValue<OptionSetValue>("ols_parenttwogender").Value);
                parentsOfDeceased.ParentTwoGenderSpecified = true;
            }
            var exname2_2 = new ExtendedNameType();
            exname2_2.FamilyNameAtBirth = funeral.Contains("ols_parenttwofamilynameatbirth") ? funeral.GetAttributeValue<string>("ols_parenttwofamilynameatbirth").ToUpper() : string.Empty;
            var name2_2 = new NameType();
            name2_2.FamilyName = funeral.Contains("ols_parenttwofamilyname") ? funeral.GetAttributeValue<string>("ols_parenttwofamilyname").ToUpper() : string.Empty;
            name2_2.FirstGivenName = funeral.Contains("ols_parenttwofirstgivenname") ? funeral.GetAttributeValue<string>("ols_parenttwofirstgivenname") : string.Empty;
            name2_2.OtherGivenNames = funeral.Contains("ols_parenttwoothergivennames") ? funeral.GetAttributeValue<string>("ols_parenttwoothergivennames") : string.Empty;
            exname2_2.Name = name2_2;
            parentsOfDeceased.ParentTwoName = exname2_2;
            var occupation2_2 = new OccupationType();
            occupation2_2.Occupation = funeral.Contains("ols_parenttwousualoccupation") ? funeral.GetAttributeValue<string>("ols_parenttwousualoccupation") : string.Empty;
            occupation2_2.MainOccupationActivity = funeral.Contains("ols_parenttwomaintasksperformed") ? funeral.GetAttributeValue<string>("ols_parenttwomaintasksperformed") : string.Empty;
            parentsOfDeceased.ParentTwoOccupation = occupation2_2;
        }

        private DisposalDetails GetDisposalDetails(Entity funeral)
        {
            var disposalDetails = new DisposalDetails();
            disposalDetails.DisposalDetails1 = new BodyDisposalDataType();

            if (funeral.Contains("ols_methodofdisposal"))
            {
                disposalDetails.MethodOfDisposal = GetMethodOfDisposal(funeral.GetAttributeValue<OptionSetValue>("ols_methodofdisposal").Value);
                disposalDetails.MethodOfDisposalSpecified = true;

                if (disposalDetails.MethodOfDisposal == DisposalDetailsMethodOfDisposal.Cremated ||
                    disposalDetails.MethodOfDisposal == DisposalDetailsMethodOfDisposal.BodyDonation ||
                    disposalDetails.MethodOfDisposal == DisposalDetailsMethodOfDisposal.Buried)
                {


                    #region Address
                    disposalDetails.DisposalDetails1.DisposalAddress = new AddressType();
                    if (!funeral.Contains("ols_disposalcountryid") || (funeral.Contains("ols_disposalcountryid") && funeral.GetAttributeValue<EntityReference>("ols_disposalcountryid").Name == "Australia"))
                    {
                        disposalDetails.DisposalDetails1.DisposalAddress.Line1 = funeral.Contains("ols_disposaladdressline1") ? funeral.GetAttributeValue<string>("ols_disposaladdressline1") : string.Empty;
                        disposalDetails.DisposalDetails1.DisposalAddress.Line2 = funeral.Contains("ols_disposaladdressline2") ? funeral.GetAttributeValue<string>("ols_disposaladdressline2") : string.Empty;
                        if (funeral.Contains("ols_disposaladdresspostcodeid"))
                        {
                            disposalDetails.DisposalDetails1.DisposalAddress.Postcode = Convert.ToInt16(funeral.GetAttributeValue<EntityReference>("ols_disposaladdresspostcodeid").Name);
                            disposalDetails.DisposalDetails1.DisposalAddress.PostcodeSpecified = true;
                        }
                        disposalDetails.DisposalDetails1.DisposalAddress.Country = "Australia";
                        disposalDetails.DisposalDetails1.DisposalAddress.SuburbOrTown = funeral.Contains("ols_disposaladdresssuburbid") ? funeral.GetAttributeValue<EntityReference>("ols_disposaladdresssuburbid").Name : string.Empty;
                        if (funeral.Contains("ols_disposaladdressstateid"))
                        {
                            disposalDetails.DisposalDetails1.DisposalAddress.State = GetState(funeral.GetAttributeValue<EntityReference>("ols_disposaladdressstateid").Name);
                            disposalDetails.DisposalDetails1.DisposalAddress.StateSpecified = true;
                        }
                    }
                    else
                    {
                        disposalDetails.DisposalDetails1.DisposalAddress.Country = funeral.Contains("ols_disposalcountryid") ? funeral.GetAttributeValue<EntityReference>("ols_disposalcountryid").Name : string.Empty;
                        disposalDetails.DisposalDetails1.DisposalAddress.InternationalAddress = funeral.Contains("ols_disposalinternationaladdress") ? funeral.GetAttributeValue<string>("ols_disposalinternationaladdress") : string.Empty;
                    }
                    #endregion

                    if (funeral.Contains("ols_dateofdisposal"))
                    {
                        disposalDetails.DateOfDisposal = Util.LocalFromUTCUserDateTime(GetService(UserType.User), funeral.GetAttributeValue<DateTime>("ols_dateofdisposal"));
                        disposalDetails.DateOfDisposalSpecified = true;
                    }
                    switch (disposalDetails.MethodOfDisposal)
                    {
                        case DisposalDetailsMethodOfDisposal.Cremated:
                            disposalDetails.DisposalDetails1.DisposalOrganisationName = funeral.Contains("ols_crematoriumid") ? funeral.GetAttributeValue<EntityReference>("ols_crematoriumid").Name : string.Empty;
                            break;
                        case DisposalDetailsMethodOfDisposal.BodyDonation:
                            disposalDetails.DisposalDetails1.DisposalOrganisationName = funeral.Contains("ols_disposalorganizationname") ? funeral.GetAttributeValue<string>("ols_disposalorganizationname") : string.Empty;
                            break;
                        case DisposalDetailsMethodOfDisposal.Buried:
                            if (funeral.Contains("ols_cemeteryid"))
                                disposalDetails.DisposalDetails1.DisposalOrganisationName = funeral.GetAttributeValue<EntityReference>("ols_cemeteryid").Name;
                            break;
                    }
                }
                else if (disposalDetails.MethodOfDisposal == DisposalDetailsMethodOfDisposal.BodyNotRecovered)
                {
                    disposalDetails.DisposalDetails1.NoBodyFound = funeral.Contains("ols_nobodyfound") ? funeral.GetAttributeValue<bool>("ols_nobodyfound") : false;
                }
                else if (disposalDetails.MethodOfDisposal == DisposalDetailsMethodOfDisposal.Repatriated)
                {
                    disposalDetails.DisposalDetails1.DestinationCountry = funeral.Contains("ols_destinationcountryid") ? funeral.GetAttributeValue<EntityReference>("ols_destinationcountryid").Name : string.Empty;
                    disposalDetails.DisposalDetails1.DisposalOrganisationName = funeral.Contains("ols_airportofdelivery") ? funeral.GetAttributeValue<string>("ols_airportofdelivery") : string.Empty;
                }
            }
            return disposalDetails;
        }

        private DisposalDetailsMethodOfDisposal GetMethodOfDisposal(int method)
        {
            switch (method)
            {
                case 1: return DisposalDetailsMethodOfDisposal.BodyDonation;
                case 2: return DisposalDetailsMethodOfDisposal.BodyNotRecovered;
                case 3: return DisposalDetailsMethodOfDisposal.Buried;
                case 4: return DisposalDetailsMethodOfDisposal.Cremated;
                case 5: return DisposalDetailsMethodOfDisposal.Repatriated;
                default: return DisposalDetailsMethodOfDisposal.Cremated;
            }
        }

        private FuneralDirectorDetails GetFuneralDirectorDetails(Entity funeral)
        {
            var funeralDirectorDetails = new FuneralDirectorDetails();

            #region Get Brand Details
            if (brand == null)
            {
                EntityReference brandRef = funeral.Contains("pricelevelid") ? funeral.GetAttributeValue<EntityReference>("pricelevelid") : null;
                if (brandRef != null)
                    brand = Retrieve(UserType.User, brandRef.LogicalName, brandRef.Id, new ColumnSet(true));
            }
            #endregion
            if (brand != null)
            {

                funeralDirectorDetails.CompanyName = brand.Contains("ols_companyname") ? brand.GetAttributeValue<string>("ols_companyname") : string.Empty;
                funeralDirectorDetails.ContactName = brand.Contains("ols_contactname") ? brand.GetAttributeValue<string>("ols_contactname") : string.Empty;
                funeralDirectorDetails.FuneralDirectorID = brand.Contains("ols_funeraldirectorid") ? brand.GetAttributeValue<string>("ols_funeraldirectorid") : string.Empty;

                #region Address
                funeralDirectorDetails.FuneralDirectorAddress = new AddressType();
                if (!brand.Contains("ols_countryid") || (brand.Contains("ols_countryid") && brand.GetAttributeValue<EntityReference>("ols_countryid").Name == "Australia"))
                {
                    funeralDirectorDetails.FuneralDirectorAddress.Line1 = brand.Contains("ols_street1") ? brand.GetAttributeValue<string>("ols_street1") : string.Empty;
                    funeralDirectorDetails.FuneralDirectorAddress.Line2 = brand.Contains("ols_street2") ? brand.GetAttributeValue<string>("ols_street2") : string.Empty;
                    if (brand.Contains("ols_postcode"))
                    {
                        funeralDirectorDetails.FuneralDirectorAddress.Postcode = Convert.ToInt16(brand.GetAttributeValue<string>("ols_postcode"));
                        funeralDirectorDetails.FuneralDirectorAddress.PostcodeSpecified = true;
                    }
                    funeralDirectorDetails.FuneralDirectorAddress.Country = "Australia";
                    funeralDirectorDetails.FuneralDirectorAddress.SuburbOrTown = brand.Contains("ols_suburbid") ? brand.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                    if (brand.Contains("ols_stateid"))
                    {
                        funeralDirectorDetails.FuneralDirectorAddress.State = GetState(brand.GetAttributeValue<EntityReference>("ols_stateid").Name);
                        funeralDirectorDetails.FuneralDirectorAddress.StateSpecified = true;
                    }

                }
                else
                {
                    funeralDirectorDetails.FuneralDirectorAddress.Country = brand.Contains("ols_countryid") ? brand.GetAttributeValue<EntityReference>("ols_countryid").Name : string.Empty;
                    if (brand.Contains("ols_stateid"))
                        funeralDirectorDetails.FuneralDirectorAddress.InternationalAddress = brand.Contains("ols_internationaladdress") ? brand.GetAttributeValue<string>("ols_internationaladdress") : string.Empty;
                }

                #endregion

                funeralDirectorDetails.FuneralDirectorContacts = new ContactDetailsType();
                funeralDirectorDetails.FuneralDirectorContacts.ContactEmail = brand.Contains("ols_contactemail") ? brand.GetAttributeValue<string>("ols_contactemail") : string.Empty;
                funeralDirectorDetails.FuneralDirectorContacts.ContactPhone = brand.Contains("ols_phone") ? brand.GetAttributeValue<string>("ols_phone") : string.Empty;

                funeralDirectorDetails.NameDetails = new MandatoryNameType();
                funeralDirectorDetails.NameDetails.FamilyName = brand.Contains("ols_familyname") ? brand.GetAttributeValue<string>("ols_familyname") : string.Empty;
                funeralDirectorDetails.NameDetails.FirstGivenName = brand.Contains("ols_firstgivenname") ? brand.GetAttributeValue<string>("ols_firstgivenname") : string.Empty;
                funeralDirectorDetails.NameDetails.OtherGivenNames = brand.Contains("ols_othergivennames") ? brand.GetAttributeValue<string>("ols_othergivennames") : string.Empty;
            }
            return funeralDirectorDetails;

        }

        private InformantDetailsType GetInformantDetails(Entity funeral)
        {
            AppendLog("GetInformantDetails method started..");
            #region Get Informant Details
            Entity informant = null;
            AppendLog("Retrieving informantRef");
            EntityReference informantRef = funeral.Contains("customerid") ? funeral.GetAttributeValue<EntityReference>("customerid") : null;
            if (informantRef != null)
                informant = Retrieve(UserType.User, informantRef.LogicalName, informantRef.Id, new ColumnSet(true));
            #endregion

            var informantDetails = new InformantDetailsType();
            if (informant != null)
            {
                AppendLog("Informant retrieved.");
                #region Name
                var nameType = new NameType();
                nameType.FamilyName = informant.Contains("ols_familyname") ? informant.GetAttributeValue<string>("ols_familyname").ToUpper() : string.Empty;
                nameType.OtherGivenNames = informant.Contains("ols_othergivennames") ? informant.GetAttributeValue<string>("ols_othergivennames") : string.Empty;
                nameType.FirstGivenName = informant.Contains("ols_givennames") ? informant.GetAttributeValue<string>("ols_givennames") : string.Empty;
                informantDetails.InformantName = nameType;
                #endregion

                #region Contact
                informantDetails.InformantContact = new ContactDetailsType();
                informantDetails.InformantContact.ContactEmail = informant.Contains("emailaddress1") ? informant.GetAttributeValue<string>("emailaddress1") : string.Empty;
                //informantDetails.InformantContact.ContactPhone = informant.Contains("address2_telephone1") ? informant.GetAttributeValue<string>("address2_telephone1") : string.Empty;
                informantDetails.InformantContact.ContactPhone = informant.Contains("telephone1") ? informant.GetAttributeValue<string>("telephone1") : string.Empty;
                #endregion

                #region Address
                informantDetails.InformantPostal = new AddressType();
                if (!informant.Contains("ols_countryid") || (informant.Contains("ols_countryid") && informant.GetAttributeValue<EntityReference>("ols_countryid").Name == "Australia"))
                {
                    informantDetails.InformantPostal.Line1 = informant.Contains("address1_line1") ? informant.GetAttributeValue<string>("address1_line1") : string.Empty;
                    informantDetails.InformantPostal.Line2 = informant.Contains("address1_line2") ? informant.GetAttributeValue<string>("address1_line2") : string.Empty;
                    if (informant.Contains("ols_postcodeid"))
                    {
                        AppendLog("Converting PostCode from string to number.");
                        informantDetails.InformantPostal.Postcode = Convert.ToInt16(informant.GetAttributeValue<EntityReference>("ols_postcodeid").Name);
                        AppendLog("PostCode converted from string to number.");

                        informantDetails.InformantPostal.PostcodeSpecified = true;
                    }
                    informantDetails.InformantPostal.Country = "Australia";
                    informantDetails.InformantPostal.SuburbOrTown = informant.Contains("ols_suburbid") ? informant.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                    if (informant.Contains("ols_stateid"))
                    {
                        informantDetails.InformantPostal.State = GetState(informant.GetAttributeValue<EntityReference>("ols_stateid").Name);
                        informantDetails.InformantPostal.StateSpecified = true;
                    }

                }
                else
                {
                    informantDetails.InformantPostal.Country = informant.Contains("ols_countryid") ? informant.GetAttributeValue<EntityReference>("ols_countryid").Name : string.Empty;
                    informantDetails.InformantPostal.InternationalAddress = informant.Contains("ols_address1internationaladdress") ? informant.GetAttributeValue<string>("ols_address1internationaladdress") : string.Empty;
                }

                #endregion

                #region Relationship
                if (informant.Contains("ols_personalrelationshiptodeceased"))
                {
                    informantDetails.InformantRelationship = GetRelationship(informant.GetAttributeValue<OptionSetValue>("ols_personalrelationshiptodeceased").Value); //InformantDetailsTypeInformantRelationship.
                    informantDetails.InformantRelationshipSpecified = true;
                }
                #endregion

                #region Residential Address

                informantDetails.InformantResidential = informantDetails.InformantPostal;
                #endregion
            }
            else
                AppendLog("Informant not found");
            AppendLog("GetInformantDetails method completed..");

            return informantDetails;
        }

        private InformantDetailsTypeInformantRelationship GetRelationship(int relationship)
        {
            switch (relationship)
            {
                //   case : return InformantDetailsTypeInformantRelationship.AGENT;
                case 18: return InformantDetailsTypeInformantRelationship.Aunt;
                case 19: return InformantDetailsTypeInformantRelationship.Brother;
                case 20: return InformantDetailsTypeInformantRelationship.BrotherinLaw;
                //  case: return InformantDetailsTypeInformantRelationship.CASEWORKER;
                case 21: return InformantDetailsTypeInformantRelationship.Cousin;
                case 1: return InformantDetailsTypeInformantRelationship.Daughter;
                case 22: return InformantDetailsTypeInformantRelationship.DaughterinLaw;
                case 23: return InformantDetailsTypeInformantRelationship.Defacto;
                case 24: return InformantDetailsTypeInformantRelationship.DirectorofNursing;
                case 2: return InformantDetailsTypeInformantRelationship.Executor;
                case 3: return InformantDetailsTypeInformantRelationship.Executrix;
                case 25: return InformantDetailsTypeInformantRelationship.ExSpouse;
                case 4: return InformantDetailsTypeInformantRelationship.Father;
                case 26: return InformantDetailsTypeInformantRelationship.FatherinLaw;
                case 27: return InformantDetailsTypeInformantRelationship.FosterFather;
                case 28: return InformantDetailsTypeInformantRelationship.FosterMother;
                case 29: return InformantDetailsTypeInformantRelationship.Friend;
                case 30: return InformantDetailsTypeInformantRelationship.FuneralDirector;
                case 31: return InformantDetailsTypeInformantRelationship.Goddaughter;
                case 32: return InformantDetailsTypeInformantRelationship.Godfather;
                case 33: return InformantDetailsTypeInformantRelationship.Godmother;
                case 34: return InformantDetailsTypeInformantRelationship.Godson;
                case 5: return InformantDetailsTypeInformantRelationship.Granddaughter;
                case 6: return InformantDetailsTypeInformantRelationship.Grandfather;
                case 7: return InformantDetailsTypeInformantRelationship.Grandmother;
                case 35: return InformantDetailsTypeInformantRelationship.Grandnephew;
                case 36: return InformantDetailsTypeInformantRelationship.Grandniece;
                case 8: return InformantDetailsTypeInformantRelationship.Grandson;
                case 39: return InformantDetailsTypeInformantRelationship.GreatAunt;
                case 37: return InformantDetailsTypeInformantRelationship.GreatGranddaughter;
                case 38: return InformantDetailsTypeInformantRelationship.GreatGrandson;
                case 40: return InformantDetailsTypeInformantRelationship.GreatUncle;
                case 9: return InformantDetailsTypeInformantRelationship.Guardian;
                case 41: return InformantDetailsTypeInformantRelationship.HalfBrother;
                case 42: return InformantDetailsTypeInformantRelationship.HalfSister;
                //  case: return InformantDetailsTypeInformantRelationship.INFORMANT;
                case 10: return InformantDetailsTypeInformantRelationship.Mother;
                case 43: return InformantDetailsTypeInformantRelationship.MotherinLaw;
                case 11: return InformantDetailsTypeInformantRelationship.Nephew;
                case 12: return InformantDetailsTypeInformantRelationship.Niece;
                case 44: return InformantDetailsTypeInformantRelationship.NoRelation;
                case 13: return InformantDetailsTypeInformantRelationship.Parent;
                //  case: return InformantDetailsTypeInformantRelationship.PARENTS;
                case 45: return InformantDetailsTypeInformantRelationship.SecondCousin;
                // case: return InformantDetailsTypeInformantRelationship.SELF;
                case 46: return InformantDetailsTypeInformantRelationship.Sister;
                case 47: return InformantDetailsTypeInformantRelationship.SisterinLaw;
                case 55: return InformantDetailsTypeInformantRelationship.Son;
                case 48: return InformantDetailsTypeInformantRelationship.SoninLaw;
                case 49: return InformantDetailsTypeInformantRelationship.StepBrother;
                case 14: return InformantDetailsTypeInformantRelationship.StepDaughter;
                case 50: return InformantDetailsTypeInformantRelationship.StepFather;
                case 51: return InformantDetailsTypeInformantRelationship.StepMother;
                case 52: return InformantDetailsTypeInformantRelationship.StepSister;
                case 15: return InformantDetailsTypeInformantRelationship.StepSon;
                case 53: return InformantDetailsTypeInformantRelationship.ThirdCousin;
                case 54: return InformantDetailsTypeInformantRelationship.Uncle;
                case 16: return InformantDetailsTypeInformantRelationship.Widow;
                case 17: return InformantDetailsTypeInformantRelationship.Widower;
                default: return InformantDetailsTypeInformantRelationship.NoRelation;
            }
        }

        private EntityCollection GetMarriageHistoryEntity(Guid OpportunityId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_marriagedetails");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, OpportunityId);
                qe.Orders.Add(new OrderExpression("ols_sequencenumber", OrderType.Ascending));
                return RetrieveMultiple(UserType.User, qe);
            }
            catch { return null; }
        }

        private MarriageDetailsType GetMarriageDetails(Entity marriagedetails)
        {

            var marriageDetails = new MarriageDetailsType();
            if (marriagedetails.Contains("ols_marriagestatus"))
                marriageDetails.MarriageStatus = GetMarriageStatus(marriagedetails.GetAttributeValue<OptionSetValue>("ols_marriagestatus").Value);
            if (marriageDetails.MarriageStatus == MaritalStatusType.Defacto ||
                marriageDetails.MarriageStatus == MaritalStatusType.Divorced ||
                marriageDetails.MarriageStatus == MaritalStatusType.Married ||
                marriageDetails.MarriageStatus == MaritalStatusType.Widowed)
            {
                if (marriagedetails.Contains("ols_ageatdateofmarriage"))
                {
                    marriageDetails.AgeAtCommencement = marriagedetails.GetAttributeValue<int>("ols_ageatdateofmarriage");
                    marriageDetails.AgeAtCommencementSpecified = true;
                }

                var name = new NameType();
                name.FirstGivenName = marriagedetails.Contains("ols_firstgivenname") ? marriagedetails.GetAttributeValue<string>("ols_firstgivenname") : string.Empty;
                name.FamilyName = marriagedetails.Contains("ols_spousefamilyname") ? marriagedetails.GetAttributeValue<string>("ols_spousefamilyname").ToUpper() : string.Empty;
                name.OtherGivenNames = marriagedetails.Attributes.Contains("ols_spouseothergivennames") ? marriagedetails.Attributes["ols_spouseothergivennames"].ToString() : string.Empty;
                marriageDetails.SpouseName = name;

                if (marriagedetails.Contains("ols_spousegender"))
                {
                    marriageDetails.SpouseGender = GetGender(marriagedetails.GetAttributeValue<OptionSetValue>("ols_spousegender").Value);
                    marriageDetails.SpouseGenderSpecified = true;
                }

            }
            if (marriageDetails.MarriageStatus == MaritalStatusType.Divorced ||
                marriageDetails.MarriageStatus == MaritalStatusType.Married ||
                marriageDetails.MarriageStatus == MaritalStatusType.Widowed)
            {
                var address = new AddressSuburbStateCountryType();
                if (!marriagedetails.Contains("ols_countryid") || (marriagedetails.Contains("ols_countryid") && marriagedetails.GetAttributeValue<EntityReference>("ols_countryid").Name == "Australia"))
                {
                    address.Country = "Australia";
                    if (marriagedetails.Contains("ols_stateid"))
                    {
                        address.StateorTerritory = GetStateofBirth(marriagedetails.GetAttributeValue<EntityReference>("ols_stateid").Name);
                        address.StateorTerritorySpecified = true;
                    }
                    if (marriagedetails.Contains("ols_suburbid"))
                        address.SuburbOrTownOrCity = marriagedetails.GetAttributeValue<EntityReference>("ols_suburbid").Name;
                }
                else
                {
                    address.Country = marriagedetails.Contains("ols_countryid") ? marriagedetails.GetAttributeValue<EntityReference>("ols_countryid").Name : string.Empty;
                    address.SuburbOrTownOrCity = marriagedetails.Contains("ols_suburbid") ? marriagedetails.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                }
                marriageDetails.MarriageAddress = address;

                marriageDetails.MarriageLocation = marriagedetails.Contains("ols_placeofmarriage") ? marriagedetails.GetAttributeValue<string>("ols_placeofmarriage") : string.Empty;
            }

            return marriageDetails;
        }
        private MaritalStatusType GetMarriageStatus(int marriagestatus)
        {
            switch (marriagestatus)
            {
                case 1: return MaritalStatusType.Defacto;
                case 2: return MaritalStatusType.Divorced;
                case 3: return MaritalStatusType.Married;
                case 4: return MaritalStatusType.NeverMarried;
                //case : return MaritalStatusType.Separated;
                case 5: return MaritalStatusType.Widowed;
                case 6: return MaritalStatusType.Unknown;
                default: return MaritalStatusType.Unknown;
            }
        }

        private PreviousMarriageDetailsType[] GetMarriageHistoryDetails(EntityCollection marriagedetailsColl)
        {
            var previousMarriageDetails = new List<PreviousMarriageDetailsType>();
            Entity firstMarrToRemove = marriagedetailsColl.Entities.FirstOrDefault();
            if (firstMarrToRemove != null)
                marriagedetailsColl.Entities.Remove(firstMarrToRemove);

            foreach (var marriagedetails in marriagedetailsColl.Entities)
            {
                if (marriagedetails.Contains("ols_firstgivenname"))
                {
                    PreviousMarriageDetailsType marriage2 = new PreviousMarriageDetailsType();
                    var name = new NameType();
                    name.FirstGivenName = marriagedetails.Contains("ols_firstgivenname") ? marriagedetails.GetAttributeValue<string>("ols_firstgivenname") : string.Empty;
                    name.FamilyName = marriagedetails.Contains("ols_spousefamilyname") ? marriagedetails.GetAttributeValue<string>("ols_spousefamilyname").ToUpper() : string.Empty; ;
                    name.OtherGivenNames = marriagedetails.Contains("ols_spouseothergivennames") ? marriagedetails.GetAttributeValue<string>("ols_spouseothergivennames") : string.Empty;
                    marriage2.SpouseName = name;

                    if (marriagedetails.Contains("ols_marriagestatus"))
                        marriage2.MarriageType = GetMarriageType(marriagedetails.GetAttributeValue<OptionSetValue>("ols_marriagestatus").Value);
                    if (marriagedetails.Contains("ols_spousegender"))
                    {
                        marriage2.SpouseGender = GetGender(marriagedetails.GetAttributeValue<OptionSetValue>("ols_spousegender").Value);
                        marriage2.SpouseGenderSpecified = true;

                    }
                    if (marriage2.MarriageType != MarriageType.DEFACTO)
                    {
                        var address = new AddressSuburbStateCountryType();
                        if (!marriagedetails.Contains("ols_countryid") || (marriagedetails.Contains("ols_countryid") && marriagedetails.GetAttributeValue<EntityReference>("ols_countryid").Name == "Australia"))
                        {
                            address.Country = "Australia";
                            if (marriagedetails.Contains("ols_stateid"))
                            {
                                address.StateorTerritory = GetStateofBirth(marriagedetails.GetAttributeValue<EntityReference>("ols_stateid").Name);
                                address.StateorTerritorySpecified = true;
                            }
                            address.SuburbOrTownOrCity = marriagedetails.Contains("ols_suburbid") ? marriagedetails.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                        }
                        else
                        {
                            address.Country = marriagedetails.Contains("ols_countryid") ? marriagedetails.GetAttributeValue<EntityReference>("ols_countryid").Name : string.Empty;
                            address.SuburbOrTownOrCity = marriagedetails.Contains("ols_suburbid") ? marriagedetails.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                        }
                        marriage2.MarriageAddress = address;
                    }
                    marriage2.MarriageLocation = marriagedetails.Contains("ols_placeofmarriage") ? marriagedetails.GetAttributeValue<string>("ols_placeofmarriage") : string.Empty;
                    if (marriagedetails.Contains("ols_ageatdateofmarriage"))
                    {
                        marriage2.AgeAtCommencement = marriagedetails.GetAttributeValue<int>("ols_ageatdateofmarriage");
                        marriage2.AgeAtCommencementSpecified = true;
                    }
                    previousMarriageDetails.Add(marriage2);
                }
            }

            return previousMarriageDetails.ToArray();
        }

        private MarriageType GetMarriageType(int marriagetype)
        {
            switch (marriagetype)
            {
                case 1: return MarriageType.DEFACTO;
                case 2: return MarriageType.DIVORCED;
                // case 9: return MarriageType.SEPARATED;
                case 3: return MarriageType.MARRIED;
                case 6: return MarriageType.UNKNOWN;
                case 5: return MarriageType.WIDOWED;
                default: return MarriageType.UNKNOWN;
            }
        }

        private TypeOfDeathCertificate GetTypeOfDeathCertificateDetails(int type)
        {
            switch (type)
            {
                case 1: return TypeOfDeathCertificate.Coroner;
                case 2: return TypeOfDeathCertificate.MCCD;
                case 3: return TypeOfDeathCertificate.MCPD;
                default: return TypeOfDeathCertificate.Coroner;
            }
        }

        private string SendSMS()
        {
            var smsres = "";
            if (string.IsNullOrEmpty(To)) smsres = "Error: Notification Mobile Number is blank";
            else
            {
                var config = new ConfigData(GetService(UserType.User));

                smsData = new SMSData(config.SMSAccount, config.SMSPassword);//"liudmila", "f67de470");
                smsData.senderId = config.SenderId;
                smsData.MobileNumber = To;

                smsData.Message = Message;

                smsres = Send();
            }
            return smsres;
        }

        private string SendEmail(Entity funeral)
        {
            var emailres = "";

            if (string.IsNullOrEmpty(EmailTo)) emailres = "Error: Notification Email Address is blank";
            else
            {
                try
                {
                    Entity Fromparty = new Entity("activityparty");
                    Entity Toparty = new Entity("activityparty");
                    var contact = GetContactByEmail(EmailTo);

                    if (contact == null) return string.Empty;

                    Toparty["partyid"] = contact;

                    if (funeral.Contains("ownerid"))
                        Fromparty["partyid"] = new EntityReference("systemuser", funeral.GetAttributeValue<EntityReference>("ownerid").Id);


                    Entity _Email = new Entity("email");
                    _Email["from"] = new Entity[] { Fromparty };
                    _Email["to"] = new Entity[] { Toparty };
                    _Email["directioncode"] = true;
                    _Email["description"] = Message;

                    _Email["regardingobjectid"] = new EntityReference("opportunity", funeral.Id);

                    Guid EmailID = Create(UserType.User, _Email);

                    var sendEmailreq = new SendEmailRequest()
                    {
                        EmailId = EmailID,
                        TrackingToken = "",
                        IssueSend = true
                    };
                    var sendEmailresp = Execute(UserType.User, sendEmailreq);


                    emailres = "OK";

                }
                catch (Exception ex) { return "Error:" + ex.Message; }
            }
            return emailres;
        }

        private EntityReference GetContactByEmail(string emailaddress)
        {
            try
            {
                EntityReference contactRef = null;

                QueryExpression qeContact = new QueryExpression("contact");
                qeContact.ColumnSet = new ColumnSet("contactid");
                qeContact.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, emailaddress);
                Entity contact = RetrieveMultiple(UserType.User, qeContact).Entities.FirstOrDefault();
                if (contact != null)
                    contactRef = new EntityReference("contact", contact.Id);

                QueryExpression qeAccount = new QueryExpression("account");
                qeAccount.ColumnSet = new ColumnSet("accountid");
                qeAccount.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, emailaddress);
                Entity account = RetrieveMultiple(UserType.User, qeAccount).Entities.FirstOrDefault();
                if (account != null)
                    contactRef = new EntityReference("account", account.Id);


                QueryExpression qeUser = new QueryExpression("systemuser");
                qeUser.ColumnSet = new ColumnSet("systemuserid");
                qeUser.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, emailaddress);
                Entity user = RetrieveMultiple(UserType.User, qeUser).Entities.FirstOrDefault();
                if (user != null)
                    contactRef = new EntityReference("systemuser", user.Id);

                return contactRef;
            }
            catch (Exception ex) { throw ex; }
        }

        private void SetSMSDetails(Entity funeral)
        {
            #region Retrieve Contractor
            Entity contractor = null;
            if (funeral.Contains("ols_contractorid"))
            {
                EntityReference contractorRef = funeral.GetAttributeValue<EntityReference>("ols_contractorid");
                if (contractorRef != null)
                    contractor = Retrieve(UserType.User, contractorRef.LogicalName, contractorRef.Id, new ColumnSet(true));
            }
            #endregion

            if (contractor != null)
            {
                To = contractor.Contains("ols_notificationmobilenumber") ? contractor.GetAttributeValue<string>("ols_notificationmobilenumber") : string.Empty;
                EmailTo = contractor.Contains("ols_notificationemailaddress") ? contractor.GetAttributeValue<string>("ols_notificationemailaddress") : string.Empty;
                ContactMethod = contractor.Contains("ols_contactmethod") ? contractor.GetAttributeValue<OptionSetValue>("ols_contactmethod").Value : 0;
            }
            var mess = new StringBuilder();
            var brand = "";
            var transfer = "";
            var address = new StringBuilder();

            EntityReference brandRef = funeral.Contains("pricelevelid") ? funeral.GetAttributeValue<EntityReference>("pricelevelid") : null;
            if (brandRef != null)
                brand = brandRef.Name + ",";
            if (funeral.Contains("ols_funeralnumber"))
                mess.AppendLine(brand + " " + funeral.GetAttributeValue<string>("ols_funeralnumber"));
            if (funeral.Contains("name"))
                mess.AppendLine(funeral.GetAttributeValue<string>("name"));

            EntityReference coronerRef = funeral.Contains("ols_coronerid") ? funeral.GetAttributeValue<EntityReference>("ols_coronerid") : null;
            int transferFrom = funeral.Contains("ols_transferfrom") ? funeral.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value : 0;
            if (transferFrom == 4 && coronerRef != null)
            {
                transfer = coronerRef.Name;
                address.Append(funeral.Contains("ols_transferstreetaddress") ? funeral.GetAttributeValue<string>("ols_transferstreetaddress") : string.Empty);
                address.Append(" ");
                address.Append(funeral.Contains("ols_transfersuburbid") ? funeral.GetAttributeValue<EntityReference>("ols_transfersuburbid").Name : string.Empty);
                address.Append(" ");
                address.Append(funeral.Contains("ols_transferstateid") ? funeral.GetAttributeValue<EntityReference>("ols_transferstateid").Name : string.Empty);
                address.Append(" ");
                address.Append(funeral.Contains("ols_transferpostcodeid") ? funeral.GetAttributeValue<EntityReference>("ols_transferpostcodeid").Name : string.Empty);
            }
            else
            {
                if (transferFrom == 1)
                {
                    transfer = funeral.Contains("ols_transferlocationid") ? funeral.GetAttributeValue<EntityReference>("ols_transferlocationid").Name : funeral.Contains("ols_placeofdeath") ? funeral.GetAttributeValue<string>("ols_placeofdeath") : string.Empty;
                    address.Append(funeral.Contains("ols_placeofdeathstreet") ? funeral.GetAttributeValue<string>("ols_placeofdeathstreet") : string.Empty);
                    address.Append(" ");
                    address.Append(funeral.Contains("ols_placeofdeathcityid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofdeathcityid").Name : string.Empty);
                    address.Append(" ");
                    address.Append(funeral.Contains("ols_placeofdeathstateid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofdeathstateid").Name : string.Empty);
                    address.Append(" ");
                    address.Append(funeral.Contains("ols_placeofdeathpostcodeid") ? funeral.GetAttributeValue<EntityReference>("ols_placeofdeathpostcodeid").Name : string.Empty);
                }
                else
                {
                    address.Append(funeral.Contains("ols_usualresidenceofdeceasedstreet1") ? funeral.GetAttributeValue<string>("ols_usualresidenceofdeceasedstreet1") : string.Empty);
                    address.Append(" ");
                    address.Append(funeral.Contains("ols_usualresidenceofcityid") ? funeral.GetAttributeValue<EntityReference>("ols_usualresidenceofcityid").Name : string.Empty);
                    address.Append(" ");
                    address.Append(funeral.Contains("ols_usualresidenceofdeceasedstateid") ? funeral.GetAttributeValue<EntityReference>("ols_usualresidenceofdeceasedstateid").Name : string.Empty);
                    address.Append(" ");
                    address.Append(funeral.Contains("ols_postcodeid") ? funeral.GetAttributeValue<EntityReference>("ols_postcodeid").Name : string.Empty);
                }
            }
            mess.AppendLine("Transfer from: " + transfer);
            mess.AppendLine("Address:" + address.ToString());
            Message = mess.ToString();
        }

        private void UpdateSMSStatus(string res, string emailres, Entity target)
        {
            try
            {
                //Entity update
                //if (Ent.Contains("new_smsstatus")) Ent.Attributes["new_smsstatus"] = res;
                //else Ent.Attributes.Add("new_smsstatus", res);

                //if (Ent.Contains("new_emailstatus")) Ent.Attributes["new_emailstatus"] = emailres;
                //else Ent.Attributes.Add("new_emailstatus", emailres);

                //if (Ent.Contains("new_smswassent")) Ent.Attributes["new_smswassent"] = res == "OK" || emailres == "OK";
                //else Ent.Attributes.Add("new_smswassent", res == "OK" || emailres == "OK");
                //if (Ent.Contains("new_sendsms")) Ent.Attributes["new_sendsms"] = res == "OK" || emailres == "OK";
                //else Ent.Attributes.Add("new_sendsms", res == "OK" || emailres == "OK");

                //EntityHelper.UpdateEntity(svc, Ent);
            }
            catch (Exception ex) { throw ex; }
        }

        #region Response Methods
        public static string GetNotificationId(string xml)
        {
            var id = string.Empty;
            var doc = GetXmlDocument(xml);
            var tagNotificationList = doc.GetElementsByTagName("Notification");
            foreach (XmlNode tagNotification in tagNotificationList)
            {
                for (int i = 0; i < tagNotification.ChildNodes.Count; i++)
                {
                    if (tagNotification.ChildNodes[i].Name == "Id")
                        return tagNotification.ChildNodes[i].InnerText;
                }
            }
            return id;
        }
        public static string GetApplicationId(string xml)
        {
            var id = string.Empty;
            var doc = GetXmlDocument(xml);
            var tagNotificationList = doc.GetElementsByTagName("Application");
            foreach (XmlNode tagNotification in tagNotificationList)
            {
                for (int i = 0; i < tagNotification.ChildNodes.Count; i++)
                {
                    if (tagNotification.ChildNodes[i].Name == "Id")
                        return tagNotification.ChildNodes[i].InnerText;
                }
            }
            return id;
        }
        public static string GetNotificationStatus(string xml)
        {
            var status = string.Empty;
            var doc = GetXmlDocument(xml);
            var tagNotificationList = doc.GetElementsByTagName("Notification");
            foreach (XmlNode tagNotification in tagNotificationList)
            {
                for (int i = 0; i < tagNotification.ChildNodes.Count; i++)
                {
                    if (tagNotification.ChildNodes[i].Name == "Status")
                        return tagNotification.ChildNodes[i].InnerText;
                }
            }
            return status;
        }
        public static string GetApplicationStatus(string xml)
        {
            var status = string.Empty;
            var doc = GetXmlDocument(xml);
            var tagNotificationList = doc.GetElementsByTagName("Application");
            foreach (XmlNode tagNotification in tagNotificationList)
            {
                for (int i = 0; i < tagNotification.ChildNodes.Count; i++)
                {
                    if (tagNotification.ChildNodes[i].Name == "Status")
                        return tagNotification.ChildNodes[i].InnerText;
                }
            }
            return status;
        }
        public static List<string> GetErrorDetails(string xml)
        {
            var errors = new List<string>();
            var doc = GetXmlDocument(xml);
            var tagErrorDetail = doc.GetElementsByTagName("ErrorDetail");
            foreach (XmlNode tagdetail in tagErrorDetail)
            {
                errors.Add(tagdetail.InnerText);
            }
            return errors;
        }
        private static XmlDocument GetXmlDocument(string xml)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            return doc;
        }

        public static string GetNotificationStatusFromRequest(string xml)
        {
            var status = string.Empty;
            var doc = GetXmlDocument(xml);
            var tagStatusList = doc.GetElementsByTagName("Status");
            foreach (XmlNode tagStatus in tagStatusList)
            {
                return tagStatus.InnerText;
            }
            return status;
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
