using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using Olsens.Plugins.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Xml.Serialization;

namespace Olsens.Plugins.MYOBOpportunity
{
    public class PostUpdateWrapper : PluginHelper
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PostUpdateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        #region [Variables]
        public int FuneralStatus = -1;
        public string InvoiceNumber = string.Empty;
        public Money Paid = new Money(0);
        private string AccessToken = string.Empty;
        private bool isCloud = false;
        private ConfigData _config;
        string arrangerFirstName = string.Empty;
        string arrangerLastName = string.Empty;
        string arrangerTitle = string.Empty;
        string arrangerMobile = string.Empty;
        string arrangerEmail = string.Empty;
        Response objResponse = null;
        string MYOBAPIBaseURL = string.Empty;

        #endregion

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Opportunity PostUpdate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];
                Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;
                Entity postImage = Context.PostEntityImages.Contains("PostImage") ? (Entity)Context.PostEntityImages["PostImage"] : null;

                if ((target == null || target.LogicalName != "opportunity" || target.Id == Guid.Empty) || (preImage == null || preImage.Id == Guid.Empty) || (postImage == null || postImage.Id == Guid.Empty))
                {
                    AppendLog("Target/PreImage/PostImage is null");
                    return;
                }
                if (Context.Depth > 1)
                    return;

                FuneralStatus = postImage.Contains("ols_prepaidstatus") ? postImage.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value : -1;
                InvoiceNumber = postImage.Contains("ols_invoicenumber") ? postImage.GetAttributeValue<string>("ols_invoicenumber").Trim() : string.Empty;
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber").Trim() : string.Empty;
                objResponse = new Response();
                objResponse.Success = true;

                #region MYOB Registration
                bool preData_SendToMyOB = preImage.Contains("ols_sendtomyob") ? preImage.GetAttributeValue<bool>("ols_sendtomyob") : false;
                bool postData_SendToMyOB = postImage.Contains("ols_sendtomyob") ? postImage.GetAttributeValue<bool>("ols_sendtomyob") : false;
                if (preData_SendToMyOB != postData_SendToMyOB && postData_SendToMyOB && postImage.Contains("ols_readytoinvoice") && postImage.GetAttributeValue<bool>("ols_readytoinvoice"))
                {
                    AppendLog(string.Format("MYOB - SendToMYOB {0}", postImage.Contains("ols_invoicenumber") ? postImage.GetAttributeValue<string>("ols_invoicenumber") : string.Empty));

                    var res = "";
                    try
                    {
                        AppendLog("------------------------MYOB-Start-------------------------------------");
                        res = ExecuteMYOB(false, Context.InitiatingUserId, postImage);

                        if (!string.IsNullOrEmpty(InvoiceNumber) && !res.Contains("Error"))
                        {
                            AppendLog("MYOB - SetPaidAmount started");
                            res = SetPaidAmount(funeralNumber, Context.InitiatingUserId);
                            AppendLog("MYOB - SetPaidAmount completed");
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog(ex.Message);
                        res = ex.Message;
                        objResponse.Success = false;
                        objResponse.Message = ex.Message;
                    }
                    finally
                    {
                        UpdateStatus(postImage.Id, res, Paid, objResponse);
                        AppendLog("------------------------MYOB-End-------------------------------------");
                    }
                }
                #endregion

                #region MYOB OnLine Registration
                bool preData_SendToMyOBOnline = preImage.Contains("ols_sendtomyobonline") ? preImage.GetAttributeValue<bool>("ols_sendtomyobonline") : false;
                bool postData_SendToMyOBOnline = postImage.Contains("ols_sendtomyobonline") ? postImage.GetAttributeValue<bool>("ols_sendtomyobonline") : false;
                if (preData_SendToMyOBOnline != postData_SendToMyOBOnline && postData_SendToMyOBOnline && postImage.Contains("ols_readytoinvoice") && postImage.GetAttributeValue<bool>("ols_readytoinvoice"))
                {
                    AppendLog("------------------------MYOB_Online-Start-------------------------------------");
                    AppendLog(string.Format("MYOB Online- SendToMYOB {0}", InvoiceNumber));

                    var res = "";
                    try
                    {
                        AppendLog("MYOB Online - Execute");
                        res = ExecuteMYOB(true, Context.InitiatingUserId, postImage);

                        AppendLog("MYOB Online - Executed");
                        AppendLog("RES=" + res);
                        if (InvoiceNumber != string.Empty && !res.Contains("Error"))
                        {
                            AppendLog("MYOB Online - SetPaidAmount started");
                            res = SetPaidAmount(funeralNumber, Context.InitiatingUserId);
                            AppendLog("MYOB Online - SetPaidAmount completed:" + res);
                        }
                    }
                    catch (Exception ex)
                    {
                        res = ex.Message;
                        objResponse.Success = false;
                        objResponse.Message = ex.Message;
                    }
                    finally
                    {
                        AppendLog("MYOB - UpdateStatus");
                        UpdateStatus(postImage.Id, res, Paid, objResponse);
                        AppendLog("------------------------MYOB_Online-End-------------------------------------");
                    }

                }
                #endregion

                #region MYOB Getting Paid Amount
                bool preData_GetPaidAmount = preImage.Contains("ols_getpaidamount") ? preImage.GetAttributeValue<bool>("ols_getpaidamount") : false;
                bool postData_GetPaidAmount = postImage.Contains("ols_getpaidamount") ? postImage.GetAttributeValue<bool>("ols_getpaidamount") : false;

                if (preData_GetPaidAmount != postData_GetPaidAmount && postData_GetPaidAmount && !string.IsNullOrEmpty(InvoiceNumber))
                {
                    AppendLog("------------------------MYOB Getting Paid Amount-Start-------------------------------------");

                    var res = "";
                    try
                    {
                        isCloud = false;
                        res = SetPaidAmount(funeralNumber, Context.InitiatingUserId);
                    }
                    catch (Exception ex)
                    {
                        res = ex.Message;
                        objResponse.Success = false;
                        objResponse.Message = ex.Message;
                    }
                    finally { UpdateStatus(postImage.Id, res, Paid, objResponse); }
                    AppendLog("------------------------MYOB Getting Paid Amount-End-------------------------------------");

                }
                #endregion

                #region MYOB Online Getting Paid Amount
                bool preData_GetPaidAmountOnline = preImage.Contains("ols_getpaidamountonline") ? preImage.GetAttributeValue<bool>("ols_getpaidamountonline") : false;
                bool postData_GetPaidAmountOnline = postImage.Contains("ols_getpaidamountonline") ? postImage.GetAttributeValue<bool>("ols_getpaidamountonline") : false;

                if (preData_GetPaidAmountOnline != postData_GetPaidAmountOnline && postData_GetPaidAmountOnline && !string.IsNullOrEmpty(InvoiceNumber))
                {
                    AppendLog("------------------------MYOB Online Getting Paid Amount-Start-------------------------------------");

                    var res = "";
                    try
                    {
                        isCloud = true;
                        res = SetPaidAmount(funeralNumber, Context.InitiatingUserId);
                    }
                    catch (Exception ex)
                    {
                        res = ex.Message;
                        objResponse.Success = false;
                        objResponse.Message = ex.Message;
                    }
                    finally { UpdateStatus(postImage.Id, res, Paid, objResponse); }
                    AppendLog("------------------------MYOB Online Getting Paid Amount-End-------------------------------------");

                }
                #endregion
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #region [Public Methods]
        public void UpdateStatus(Guid opportunityId, string res, Money paid, Response objResponse)
        {
            try
            {
                Entity updateOpportunity = new Entity("opportunity", opportunityId);
                //if(!objResponse.Success)
                if (!res.Contains("Invoice Number: "))
                {
                    updateOpportunity["ols_sendtomyob"] = false;
                    updateOpportunity["ols_sendtomyobonline"] = false;
                    updateOpportunity["ols_getpaidamount"] = false;
                    updateOpportunity["ols_getpaidamountonline"] = false;
                    updateOpportunity["ols_myobstatus"] = objResponse.Message;
                }
                else
                {
                    updateOpportunity["ols_invoicenumber"] = res.Replace("Invoice Number: ", "");
                    updateOpportunity["ols_myobstatus"] = objResponse.Message;
                    updateOpportunity["ols_sendtomyob"] = false;
                    updateOpportunity["ols_sendtomyobonline"] = false;
                    updateOpportunity["ols_getpaidamount"] = false;
                    updateOpportunity["ols_getpaidamountonline"] = false;
                    updateOpportunity["ols_paid"] = paid;
                }

                Update(UserType.User, updateOpportunity);

            }
            catch (Exception ex) { AppendLog("updateStatusError:" + ex.Message); throw ex; }
        }
        public void MYOBRegistrationData()
        {
            isCloud = false;
            _config = new ConfigData(GetService(UserType.User));
        }
        public void MYOBRegistrationData(Guid clientId)
        {
            _config = new ConfigData(GetService(UserType.User));

            isCloud = true;

            AppendLog("Retrieving MYOBAccess Code from CRM settings");
            Entity setting = GetMYOBCode(clientId);

            if (setting != null && setting.Contains("ols_value"))
            {
                try
                {
                    string accessToken = setting.Contains("ols_securityvalue") ? setting.GetAttributeValue<string>("ols_securityvalue") : string.Empty;
                    if (string.IsNullOrEmpty(accessToken))
                    {
                        AppendLog(string.Format("MYOB Access Code={0}", setting.GetAttributeValue<string>("ols_value")));
                        string myOBCode = Uri.UnescapeDataString(setting.GetAttributeValue<string>("ols_value"));
                        string data = "client_id=" + _config.MYOBClientId + "&client_secret=" + _config.MYOBClientSecret + "&grant_type=authorization_code&code=" + myOBCode + "&redirect_uri=" + _config.MYOBRedirectURL;
                        string myOBTokenURL = "https://secure.myob.com/oauth2/v1/authorize/";
                        AppendLog("Getting Access Token");
                        string result = GetToken(data, myOBTokenURL);
                        if (!string.IsNullOrEmpty(result))
                        {
                            var Jserializer = new JavaScriptSerializer();
                            Response res = Jserializer.Deserialize<Response>(result);
                            AccessToken = res.access_token;
                            AppendLog("Access Token: " + AccessToken);
                            UpdateAccessToken(AccessToken, setting.Id);
                        }
                        else AppendLog("Access Token Could not retrieved.");
                    }
                    else
                    {
                        AccessToken = accessToken;
                        AppendLog("Pre Stored Access Token: " + AccessToken);
                    }
                }
                catch (Exception ex)
                {
                    AppendLog("Error occured in MYOBRegistrationData: " + ex.Message);
                    throw new InvalidPluginExecutionException(ex.Message);
                }
                finally
                {
                    //if (!CheckOtherMYOBToProcess())
                    //{
                    //    AppendLog(string.Format("Deleting MYOB Access Code from Settings"));
                    //    RemoveMYOBCode(setting.Id);
                    //    AppendLog(string.Format("MYOB Access Code was deleted from Settings"));
                    //}
                }
            }
            else AppendLog("MYOBAccess Code not found in CRM");

        }
        public Entity GetMYOBCode(Guid clientId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_setting");
                qe.ColumnSet = new ColumnSet("ols_name", "ols_value", "ols_securityvalue");
                qe.Criteria.AddCondition("ols_name", ConditionOperator.Equal, "MYOB_" + clientId.ToString().ToUpper());
                qe.Orders.Add(new OrderExpression("createdon", OrderType.Descending));
                EntityCollection settingColl = RetrieveMultiple(UserType.User, qe);
                if (settingColl != null && settingColl.Entities.Count > 0)
                {
                    #region Deleting old AccessCode settings
                    if (settingColl.Entities.Count > 1)
                        for (int i = settingColl.Entities.Count - 1; i > 0; i--)
                            RemoveMYOBCode(settingColl.Entities[i].Id);
                    #endregion

                    return settingColl.Entities.FirstOrDefault();
                }
                else
                    return null;
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in GetMYOBCode: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        public void RemoveMYOBCode(Guid SettingId)
        {
            try
            {
                Delete(UserType.User, "ols_setting", SettingId);
            }
            catch { }
        }
        public string SetPaidAmount(string funeralNumber, Guid UserId)
        {
            try
            {
                if (_config == null)
                {
                    if (isCloud)
                    {
                        // string code_value = Uri.UnescapeDataString(GetMYOBCode(svc, clientId));
                        AppendLog("UserId:" + UserId.ToString());
                        MYOBRegistrationData(UserId);
                    }
                    else
                        MYOBRegistrationData();

                    GetBasAPIURL();
                }

                Paid = GetPaidAmount(funeralNumber);
                AppendLog("Paid amount: " + Paid.Value);
                return "Invoice Number: " + InvoiceNumber;
            }
            catch (Exception ex) { return ex.Message; }
        }
        public EntityCollection GetOppProducts(Guid oppId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("opportunityproduct");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, oppId);
                return RetrieveMultiple(UserType.User, qe);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string ExecuteMYOB(bool isCloud, Guid UserId, Entity postImage)
        {
            try
            {
                AppendLog(string.Format("MYOB isCloud={0}", isCloud.ToString()));

                if (isCloud)
                {
                    AppendLog("UserId:" + UserId.ToString());
                    MYOBRegistrationData(UserId);
                }
                else
                    MYOBRegistrationData();

                GetBasAPIURL();
                if (objResponse.Success)
                    return Submit(postImage);
                else
                    return objResponse.Message;

            }
            catch (Exception ex) { return ex.Message; }
        }
        public void UpdateAccessToken(string accessToken, Guid settingId)
        {
            try
            {
                Entity updateSetting = new Entity("ols_setting", settingId);
                updateSetting["ols_securityvalue"] = accessToken;
                Update(UserType.User, updateSetting);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public bool CheckOtherMYOBToProcess()
        {
            try
            {
                QueryExpression qe = new QueryExpression("opportunity");
                qe.ColumnSet = new ColumnSet("opportunityid");
                qe.Criteria.AddCondition("ols_readytoinvoice", ConditionOperator.Equal, true);
                qe.Criteria.AddCondition("ols_sendtomyobonline", ConditionOperator.Equal, true);
                Entity opportunity = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                if (opportunity != null)
                {
                    AppendLog("Another Opportunity found to process with Id: " + opportunity.Id);
                    return true;
                }
                else return false;
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in GetMYOBCode: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }

        }
        public void GetBasAPIURL()
        {
            try
            {
                AppendLog("GetBasAPIURL() method started.");
                if (isCloud)
                {
                    int retryCount = 0;
                    bool retry = false;
                    do
                    {
                        AppendLog("GetBasAPIURL Retry Count: " + retryCount);
                        string result = CallGetAPI(_config.MYOBOnlieURL, AccessToken);
                        if (!string.IsNullOrEmpty(result) && !result.Contains("Error:"))
                        {
                            //AppendLog("GetBasAPIURL result: " + result);
                            var Jserializer = new JavaScriptSerializer();
                            List<CompanyFile> companyFileList = Jserializer.Deserialize<List<CompanyFile>>(result);
                            if (companyFileList != null && companyFileList.Count > 0)
                            {
                                CompanyFile companyFile = companyFileList.Where(a => a.Name == _config.MYOBOnlineCompanyName).FirstOrDefault();
                                if (companyFile != null)
                                {
                                    MYOBAPIBaseURL = _config.MYOBOnlieURL + "/" + companyFile.Id;
                                    AppendLog("API base URL: " + MYOBAPIBaseURL);
                                    objResponse.Success = true;
                                }
                                else
                                {
                                    objResponse.Success = false;
                                    objResponse.Message = objResponse.Message + " Company file not found";
                                    AppendLog("Company file not found");
                                }
                            }
                            else
                            {
                                objResponse.Success = false;
                                objResponse.Message = objResponse.Message + " Company file not found";
                                AppendLog("Company file not found");
                            }
                            retry = false;
                        }
                        else if (!string.IsNullOrEmpty(result))
                        {
                            AppendLog(result);
                            objResponse.Success = false;
                            objResponse.Message = result;
                            MYOBRegistrationData(Context.InitiatingUserId);
                            retry = true;
                            retryCount += 1;
                        }
                    } while (retryCount < 6 && retry);

                }
                else
                {
                    string result = CallGetAPI(_config.MYOBURL, string.Empty);
                    if (!string.IsNullOrEmpty(result) && !result.Contains("Error:"))
                    {
                        //AppendLog("GetBasAPIURL result: " + result);
                        var Jserializer = new JavaScriptSerializer();
                        List<CompanyFile> companyFileList = Jserializer.Deserialize<List<CompanyFile>>(result);
                        if (companyFileList != null && companyFileList.Count > 0)
                        {
                            //AppendLog("Data: " + _config.MYOBCompanyName + " " + _config.MYOBVersion + " " + _config.MYOBLibraryPath);
                            CompanyFile companyFile = companyFileList.Where(a => a.Name == _config.MYOBCompanyName && a.ProductVersion == _config.MYOBVersion && a.LibraryPath == _config.MYOBLibraryPath).FirstOrDefault();
                            if (companyFile != null)
                            {
                                MYOBAPIBaseURL = _config.MYOBURL + "/" + companyFile.Id;
                                AppendLog("API base URL: " + MYOBAPIBaseURL);
                            }
                            else
                            {
                                objResponse.Success = false;
                                objResponse.Message = objResponse.Message + " Company file not found";
                                AppendLog("Company file not found");
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(result))
                    {
                        objResponse.Success = false;
                        objResponse.Message = result;
                    }
                }
                AppendLog("GetBasAPIURL() method completed.");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region PaidAmount
        private Money GetPaidAmount(string funeralNumber)
        {
            try
            {
                AppendLog("GetPaidAmount method started.");
                var sum = new decimal(0);
                AppendLog("Calling Get API with Url: " + MYOBAPIBaseURL + "/Sale/CustomerPayment?$filter=Customer/DisplayID eq '" + funeralNumber + "'");
                string result = CallGetAPI(MYOBAPIBaseURL + "/Sale/CustomerPayment?$filter=Customer/DisplayID eq '" + funeralNumber + "'", AccessToken);
                AppendLog("Get API Result: " + result);
                if (!string.IsNullOrEmpty(result) && !result.Contains("Error:"))
                {

                    var Jserializer = new JavaScriptSerializer();
                    CustomerPaymentItems customerPaymentItems = Jserializer.Deserialize<CustomerPaymentItems>(result);

                    if (customerPaymentItems != null && customerPaymentItems.Items != null && customerPaymentItems.Items.Count > 0)
                    {
                        AppendLog("Deserealized into CustomerPaymentItems");
                        foreach (var customerPayment in customerPaymentItems.Items)
                        {
                            AppendLog("Total invoices: "+ customerPayment.Invoices.Count);
                            foreach (var line in customerPayment.Invoices)
                            {
                                //if (line.Number == InvoiceNumber)
                                    sum += line.AmountApplied;
                            }
                        }
                    }
                    else
                        AppendLog("Unable to deserealized into CustomerPaymentItems");

                }
                else
                {
                    objResponse.Success = false;
                    objResponse.Message = result;
                }
                AppendLog("GetPaidAmount method completed.");

                return new Money(sum);
            }
            catch (Exception ex) { AppendLog("Error in GetPaidAmount method:" + ex.Message); return new Money(0); }
        }
        #endregion

        #region[Private Methods]
        private string GetToken(string data, string url)
        {
            string result = string.Empty;
            try
            {
                var httpRequest = (HttpWebRequest)WebRequest.Create(url);
                httpRequest.Method = "POST";
                httpRequest.Timeout = 60000;//0.5 Minutes
                httpRequest.ContentType = "application/x-www-form-urlencoded";
                httpRequest.Accept = "*/*";
                httpRequest.Host = "secure.myob.com";
                httpRequest.Headers["Accept-Encoding"] = "gzip, deflate, br";

                using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
                {
                    streamWriter.Write(data);
                }

                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var errorResponse = (HttpWebResponse)wex.Response)
                    {
                        using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                        {
                            string res = reader.ReadToEnd();
                            var Jserializer = new JavaScriptSerializer();
                            APIErrorResponse APIErrorResponse = Jserializer.Deserialize<APIErrorResponse>(res);
                            if (APIErrorResponse != null && APIErrorResponse.Errors != null && APIErrorResponse.Errors.Count > 0)
                            {
                                foreach (var error in APIErrorResponse.Errors)
                                {
                                    if (string.IsNullOrEmpty(result))
                                        result = "Error: " + error.Message + " ";
                                    else
                                        result = result + error.Message + " ";
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        private Employee GenerateSalesPerson()
        {
            var myobEmployee = new Employee();
            //myobEmployee.CompanyName = companyFile.Name;
            myobEmployee.FirstName = arrangerFirstName;
            myobEmployee.LastName = arrangerLastName;
            myobEmployee.IsActive = true;
            myobEmployee.IsIndividual = true;
            return myobEmployee;
        }

        private string Submit(Entity postImage)
        {
            try
            {
                var result = string.Empty;

                //if (companyFile == null) return "Error: Could not find a Company File. Check config file.";

                #region taxCode
                TaxCode taxCodeGST = null;
                AppendLog("Retrieving TaxCode from MYOB");
                string resultTaxCode = CallGetAPI(MYOBAPIBaseURL + "/GeneralLedger/TaxCode", AccessToken);
                if (!string.IsNullOrEmpty(resultTaxCode) && !resultTaxCode.Contains("Error:"))
                {
                    TaxCodeItems objTaxCodeItems = null;
                    using (MemoryStream DeSerializememoryStream = new MemoryStream())
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(TaxCodeItems));
                        StreamWriter writer = new StreamWriter(DeSerializememoryStream);
                        writer.Write(resultTaxCode);
                        writer.Flush();
                        DeSerializememoryStream.Position = 0;
                        objTaxCodeItems = (TaxCodeItems)serializer.ReadObject(DeSerializememoryStream);
                    }

                    if (objTaxCodeItems != null && objTaxCodeItems.Items != null)
                    {
                        taxCodeGST = objTaxCodeItems.Items.Where(a => a.Code == "GST").FirstOrDefault();
                        AppendLog("TaxCode retrieved from MYOB");
                    }
                }
                else if (!string.IsNullOrEmpty(resultTaxCode))
                {
                    objResponse.Success = false;
                    objResponse.Message = objResponse.Message + resultTaxCode;
                    return resultTaxCode;
                }
                else
                {
                    objResponse.Success = false;
                    objResponse.Message = objResponse.Message + "Error: Could not find Tax Codes.";
                    AppendLog("ResultTaxCode: " + resultTaxCode);
                    return "Error: Could not find Tax Codes.";
                }

                #endregion


                #region SalesPerson
                string salesPersonId = FindSalesPersonByName(postImage);

                if (salesPersonId == string.Empty)
                {
                    AppendLog("Salesperson Id not found in MYOB, Creating Salesperson");
                    result = CreateSalesPerson();

                    salesPersonId = FindSalesPersonByName(postImage);
                    if (!string.IsNullOrEmpty(salesPersonId))
                        AppendLog("Salesperson created with Id: " + salesPersonId);
                    else
                    {
                        objResponse.Success = false;
                        objResponse.Message = objResponse.Message + "Error: Could not create salesperson. ";
                        return "Error: Could not create salesperson.  ";
                    }

                }
                #endregion

                #region   Customer
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
                string customerId = FindCustomerByFuneralNumber(funeralNumber);

                if (customerId == string.Empty)
                {
                    AppendLog("Customer Id not found in MYOB, Creating Customer");
                    result = CreateCustomer(taxCodeGST, salesPersonId, postImage);

                    customerId = FindCustomerByFuneralNumber(funeralNumber);

                    if (!string.IsNullOrEmpty(customerId))
                        AppendLog("Customer created with Id: " + customerId);
                    else
                    {
                        objResponse.Success = false;
                        objResponse.Message = objResponse.Message + "Error: Could not create customer. ";
                        return "Error: Could not create customer.  ";
                    }

                    //if (!result.Contains("Error"))
                    //{
                    //    customerId = result.Substring(result.Length - 36, 36);
                    //    AppendLog("Customer created with Id: " + customerId);
                    //}
                    //else return "Error: Could not create customer.  ";
                }
                #endregion

                #region   Job
                // Job job = new Job();
                var job = FindJobByNumber(funeralNumber);
                if (job == null)
                {
                    AppendLog("Job not found in MYOB, Creating Job");
                    job = CreateJob(postImage, customerId);
                    if (job == null)
                    {
                        objResponse.Success = false;
                        objResponse.Message = objResponse.Message + "Error:  Could not create Job. ";
                        return "Error:  Could not create Job.";
                    }
                }
                #endregion


                if (InvoiceNumber == string.Empty)
                {
                    AppendLog("Invoice number is empty.");
                    result = CreateInvoice(postImage, new Guid(customerId), new Guid(salesPersonId), taxCodeGST, job);
                }
                else result = UpdateInvoice(postImage, new Guid(customerId), new Guid(salesPersonId), taxCodeGST, job);
                return result;
            }

            catch (Exception ex)
            {
                AppendLog("Error occured in Submit: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private string UpdateInvoice(Entity postImage, Guid customerid, Guid salespersonid, TaxCode taxcode, Job job)
        {
            var res = string.Empty;
            try
            {
                AppendLog("UpdateInvoice method started.");
                EntityCollection oppProductColl = GetOppProducts(postImage.Id);
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
                ItemInvoice items = GetInvoice(InvoiceNumber);
                if (items == null)
                {
                    InvoiceNumber = GetInvoiceNumber(funeralNumber);
                    items = GetInvoice(InvoiceNumber);
                    objResponse.InvoiceNumber = InvoiceNumber;
                    AppendLog("Invoice: " + InvoiceNumber);
                }
                if (items != null)
                {
                    ItemInvoice MYOBItem = GenerateItem(oppProductColl, customerid, salespersonid, taxcode, job, funeralNumber);
                    if (MYOBItem != null)
                    {
                        AppendLog("Generated MYOBItems");
                        try
                        {
                            MYOBItem.UID = items.UID;
                            MYOBItem.RowVersion = items.RowVersion;
                            AppendLog("Rowversion: " + items.RowVersion);
                            int retryCount = 0;
                            bool retry = false;
                            do
                            {

                                AppendLog("UpdateInvoice RetryCount: " + retryCount);

                                if (retry)
                                {
                                    AppendLog("Retrieving Items for Update RetryCount:" + retryCount);
                                    //items = GetInvoice(InvoiceNumber);
                                    //if (items == null)
                                    //{
                                    InvoiceNumber = GetInvoiceNumber(funeralNumber);
                                    items = GetInvoice(InvoiceNumber);
                                    objResponse.InvoiceNumber = InvoiceNumber;
                                    AppendLog("Invoice: " + InvoiceNumber);
                                    //}
                                    if (items != null)
                                    {
                                        MYOBItem.UID = items.UID;
                                        MYOBItem.RowVersion = items.RowVersion;
                                        AppendLog("Rowversion for RetryCount: " + items.RowVersion);

                                    }
                                }

                                string jsonString = string.Empty;
                                using (var ms = new MemoryStream())
                                {
                                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(ItemInvoice));
                                    serializer.WriteObject(ms, MYOBItem);
                                    ms.Position = 0;
                                    jsonString = System.Text.Encoding.Default.GetString(ms.ToArray());
                                }
                                string result = CallPutAPI(jsonString, MYOBAPIBaseURL + "/Sale/Invoice/Item/" + items.UID, AccessToken);
                                if (string.IsNullOrEmpty(result))
                                {
                                    objResponse.Message = "Success: Updated to MYOB";
                                    objResponse.Success = true;
                                    AppendLog("Success: Updated to MYOB");
                                    retry = false;
                                }
                                else
                                {
                                    objResponse.Message = result;
                                    objResponse.Success = false;
                                    AppendLog(result);
                                    retryCount += 1;
                                    retry = true;
                                }
                            } while (retryCount < 6 && retry);

                            return InvoiceNumber;
                        }
                        catch (Exception ex)
                        {
                            AppendLog("err:" + ex.Message);
                            objResponse.Success = false;
                            objResponse.Message = objResponse.Message + "Error:" + ex.Message;
                            return "Error:" + ex.Message;
                        }
                    }
                    else return string.Empty;
                }
                else
                {
                    objResponse.Success = false;
                    objResponse.Message = objResponse.Message + "Error: No Invoice found with number: " + InvoiceNumber;
                    return "Error: No Invoice found with number: " + InvoiceNumber;
                }
                //else return "Error: Could not update Invoice. Invoice not found with number: " + InvoiceNumber;
                //}
            }
            catch (Exception ex) { return ex.Message; }
            //  try
            // {


            //  DeleteInvoice(contact.new_invoicenumber);

            //  return createInvoice(contact, customerid,salespersonid,taxcode,job);    
            // }
            //  catch (Exception ex) { return ex.Message; }
        }

        private ItemInvoice GetInvoice(string invoiceNumber)
        {
            string result = CallGetAPI(MYOBAPIBaseURL + "/Sale/Invoice/Item?$filter=Number eq '" + invoiceNumber + "'", AccessToken);
            if (!string.IsNullOrEmpty(result))
            {
                var Jserializer = new JavaScriptSerializer();
                InvoiceItems invoiceItems = Jserializer.Deserialize<InvoiceItems>(result);
                if (invoiceItems != null && invoiceItems.Items != null && invoiceItems.Items.Count > 0)
                {
                    return invoiceItems.Items.FirstOrDefault();
                }
                else
                    return null;
            }
            else
                return null;
        }

        private string CreateInvoice(Entity postImage, Guid customerId, Guid salesPersonId, TaxCode taxCode, Job job)
        {
            var res = string.Empty;
            try
            {
                AppendLog("CreateInvoice method started.");

                EntityCollection oppProductColl = GetOppProducts(postImage.Id);
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;

                ItemInvoice MYOBItem = GenerateItem(oppProductColl, customerId, salesPersonId, taxCode, job, funeralNumber);
                if (MYOBItem != null)
                {
                    AppendLog("MYOBItems generated.");

                    try
                    {
                        int retryCount = 0;
                        bool retry = false;
                        do
                        {
                            AppendLog("CreateInvoice Retry Count: " + retryCount);
                            string jsonString = string.Empty;
                            using (var ms = new MemoryStream())
                            {
                                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(ItemInvoice));
                                serializer.WriteObject(ms, MYOBItem);
                                ms.Position = 0;
                                jsonString = System.Text.Encoding.Default.GetString(ms.ToArray());
                            }
                            AppendLog("Calling Create Invoice API");
                            string result = CallPostAPI(jsonString, MYOBAPIBaseURL + "/Sale/Invoice/Item", AccessToken);
                            AppendLog("Calling Create Invoice API completed");
                            if (!string.IsNullOrEmpty(result))
                            {
                                AppendLog(result);
                                objResponse.Message = result;
                                objResponse.Success = false;
                                res = result;

                                retryCount += 1;
                                retry = true;
                            }
                            else
                            {
                                retry = false;
                                objResponse.Message = "Success: Sent to MYOB";
                                objResponse.Success = true;
                                AppendLog("Success: Sent to MYOB");
                            }
                        } while (retryCount < 6 && retry);
                    }
                    catch (Exception ex)
                    {
                        objResponse.Success = false;
                        objResponse.Message = objResponse.Message + "Error: " + ex.Message;
                        return "Error: " + ex.Message;
                    }


                    if (objResponse.Success)
                    {
                        //var invoiceid = res.Substring(res.Length - 36, 36);
                        InvoiceNumber = GetInvoiceNumber(funeralNumber);
                        if (InvoiceNumber != string.Empty)
                        {
                            objResponse.InvoiceNumber = InvoiceNumber;
                            AppendLog("Invoice Number: " + InvoiceNumber);
                            return "Invoice Number: " + InvoiceNumber;
                        }
                        else
                        {
                            objResponse.Success = false;
                            objResponse.Message = objResponse.Message + "Error: Could not get the invoice number.";
                            res = "Error: Could not get the invoice number.";
                        }
                    }
                    return res;

                }
                else
                {
                    objResponse.Success = false;
                    objResponse.Message = objResponse.Message + "Error: Could not generate Item. Unexpected error.";
                    return "Error: Could not generate Item. Unexpected error.";
                }
                // }
            }
            catch (Exception ex) { return ex.Message; }

        }

        private string GetInvoiceNumber(string funeralNumber)
        {
            string result = CallGetAPI(MYOBAPIBaseURL + "/Sale/Invoice/Item?$filter=CustomerPurchaseOrderNumber eq '" + funeralNumber.Replace("FN-", "") + "'", AccessToken);
            if (!string.IsNullOrEmpty(result))
            {
                InvoiceItems invoiceItems = null;
                var Jserializer = new JavaScriptSerializer();
                invoiceItems = Jserializer.Deserialize<InvoiceItems>(result);
                //using (MemoryStream DeSerializememoryStream = new MemoryStream())
                //{
                //    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(InvoiceItems));
                //    StreamWriter writer = new StreamWriter(DeSerializememoryStream);
                //    writer.Write(result);
                //    writer.Flush();
                //    DeSerializememoryStream.Position = 0;
                //    invoiceItems = (InvoiceItems)serializer.ReadObject(DeSerializememoryStream);
                //}

                if (invoiceItems != null && invoiceItems.Items != null && invoiceItems.Items.Count > 0)
                {
                    ItemInvoice invoice = invoiceItems.Items.FirstOrDefault();
                    if (invoice != null)
                        return invoice.Number;
                    else return string.Empty;
                }
                else return string.Empty;
            }
            else return string.Empty;
        }


        private ItemInvoice GenerateItem(EntityCollection items, Guid customerid, Guid salespersonid, TaxCode taxcode, Job job, string funeralnumber)
        {
            ItemInvoice itemInvoice = new ItemInvoice();
            Customer customer = new Customer();
            customer.UID = customerid;

            Employee salesperson = new Employee();
            salesperson.UID = salespersonid;

            itemInvoice.Customer = customer;
            itemInvoice.Salesperson = salesperson;

            itemInvoice.Terms = new InvoiceTerms();

            itemInvoice.Date = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            itemInvoice.CustomerPurchaseOrderNumber = funeralnumber.Replace("FN-", "");

            //itemInvoice.Number =  funeralnumber;
            List<ItemInvoiceLine> lines = new List<ItemInvoiceLine>();

            StringBuilder errors = new StringBuilder();
            var sorteditems = items.Entities.OrderBy(o => o.GetAttributeValue<int>("sequencenumber")).ToList();
            foreach (var item in sorteditems)
            {
                ItemInvoiceLine line = new ItemInvoiceLine();
                line.UnitPrice = item.Contains("priceperunit") ? item.GetAttributeValue<Money>("priceperunit").Value : 0;
                line.ShipQuantity = item.Contains("quantity") ? item.GetAttributeValue<decimal>("quantity") : 0;
                line.Total = line.UnitPrice * line.ShipQuantity;

                line.TaxCode = new TaxCode { UID = taxcode.UID };
                line.Job = new Job { UID = job.UID };
                string productName = item.Contains("productid") ? item.GetAttributeValue<EntityReference>("productid").Name : string.Empty;

                if (productName == "Professional Services")
                    line.Description = "Professional Services";
                else line.Description = productName;

                string productNumber = item.Contains("productid") ? GetProductNumber(item.GetAttributeValue<EntityReference>("productid").Id) : string.Empty;

                Item objItem = null;
                if (item.Contains("productid"))
                    objItem = GetItem(productNumber);

                if (objItem != null)
                {
                    line.Item = new Item { Name = objItem.Name, UID = objItem.UID, Number = objItem.Number };
                    lines.Add(line);
                }
                else
                {
                    errors.Append(productNumber + ",");
                }
            }
            if (errors.Length > 0)
            {
                objResponse.Message = "Error: Items with productcode " + errors.ToString().Substring(0, errors.Length - 1) + " does not exist.";
                objResponse.Success = false;
                AppendLog("Error: Items with productcode " + errors.ToString().Substring(0, errors.Length - 1) + " does not exist.");
                throw new Exception("Error: Items with productcode " + errors.ToString().Substring(0, errors.Length - 1) + " does not exist.");
            }
            itemInvoice.Lines = lines;
            return itemInvoice;
        }

        private string GetProductNumber(Guid productId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("product");
                qe.ColumnSet = new ColumnSet("productnumber");
                qe.Criteria.AddCondition("productid", ConditionOperator.Equal, productId);
                Entity prod = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                if (prod != null)
                    return prod.Contains("productnumber") ? prod.GetAttributeValue<string>("productnumber") : string.Empty;
                else
                    return string.Empty;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        private Item GetItem(string productcode)
        {
            try
            {
                ProductItems productItems = null;
                string result = CallGetAPI(MYOBAPIBaseURL + "/Inventory/Item?$filter=Number eq '" + productcode + "'", AccessToken);
                if (!string.IsNullOrEmpty(result))
                {
                    using (MemoryStream DeSerializememoryStream = new MemoryStream())
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(ProductItems));
                        StreamWriter writer = new StreamWriter(DeSerializememoryStream);
                        writer.Write(result);
                        writer.Flush();
                        DeSerializememoryStream.Position = 0;
                        productItems = (ProductItems)serializer.ReadObject(DeSerializememoryStream);
                    }
                }
                if (productItems != null && productItems.Items != null && productItems.Items.Count > 0)
                {
                    return productItems.Items.FirstOrDefault();
                }
                else
                    return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private Job FindJobByNumber(string funeralnumber)
        {
            try
            {
                JobItems jobItems = null;
                string result = CallGetAPI(MYOBAPIBaseURL + "/GeneralLedger/Job?$filter=Number eq '" + funeralnumber + "'", AccessToken);
                if (!string.IsNullOrEmpty(result) && !result.Contains("Error:"))
                {
                    using (MemoryStream DeSerializememoryStream = new MemoryStream())
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(JobItems));
                        StreamWriter writer = new StreamWriter(DeSerializememoryStream);
                        writer.Write(result);
                        writer.Flush();
                        DeSerializememoryStream.Position = 0;
                        jobItems = (JobItems)serializer.ReadObject(DeSerializememoryStream);
                    }

                }
                else if (!string.IsNullOrEmpty(result))
                {
                    objResponse.Success = false;
                    objResponse.Message = result;
                }
                if (jobItems != null && jobItems.Items != null && jobItems.Items.Count > 0)
                    return jobItems.Items.FirstOrDefault();
                else
                    return null;
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in FindJobByNumber: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private Job CreateJob(Entity postImage, string customerid)
        {
            Job retrievedJob = null;
            try
            {
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
                string name = postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty;
                string deceasedSurname = postImage.Contains("ols_deceasedfamilyname") ? postImage.GetAttributeValue<string>("ols_deceasedfamilyname") : string.Empty;
                DateTime? servicePlaceSession = postImage.Contains("ols_serviceplacesessionfrom") ? Util.LocalFromUTCUserDateTime(GetService(UserType.User), postImage.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom")) : (DateTime?)null;

                Job objJob = GenerateJob(funeralNumber, name, deceasedSurname, servicePlaceSession != null ? string.Format("{0:dd/MM/yyyy}", servicePlaceSession) : string.Empty, customerid);
                if (objJob != null)
                {
                    try
                    {
                        AppendLog("Creating Job for customerid: " + customerid);

                        string jsonString = string.Empty;
                        using (var ms = new MemoryStream())
                        {
                            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Job));
                            serializer.WriteObject(ms, objJob);
                            ms.Position = 0;
                            jsonString = System.Text.Encoding.Default.GetString(ms.ToArray());
                        }
                        string result = CallPostAPI(jsonString, MYOBAPIBaseURL + "/GeneralLedger/Job", AccessToken);
                        if (!string.IsNullOrEmpty(result) && result.ToLower().Contains("error"))
                            objResponse.Message = result;
                        retrievedJob = FindJobByNumber(funeralNumber);

                        if (retrievedJob != null)
                        {
                            AppendLog("Job is created with JobId: " + retrievedJob.UID);
                        }
                    }
                    catch (Exception e) { throw e; }
                }
                return retrievedJob;
            }

            catch (Exception ex)
            {
                AppendLog("Error occured in CreateJob: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private Job GenerateJob(string funeralNumber, string deceasedName, string deceasedSurname, string serviceDate, string customerId)
        {
            try
            {
                Job job = new Job();
                job.Number = funeralNumber;

                var name = deceasedName + " " + serviceDate;
                if (name.Length > 30)
                    name = deceasedSurname + " " + serviceDate;
                job.Name = name;


                job.IsHeader = false;

                Customer customer = new Customer();
                customer.UID = new Guid(customerId);

                job.LinkedCustomer = customer;

                return job;
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in GenerateJob: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private string FindSalesPersonByName(Entity postImage)
        {
            try
            {
                string SalesPersonUID = string.Empty;
                #region Get Arranger Details
                if (postImage.Contains("ownerid") && string.IsNullOrEmpty(arrangerFirstName))
                {
                    QueryExpression qe = new QueryExpression("systemuser");
                    qe.ColumnSet = new ColumnSet("firstname", "lastname", "title", "mobilephone", "internalemailaddress");
                    qe.Criteria.AddCondition("systemuserid", ConditionOperator.Equal, postImage.GetAttributeValue<EntityReference>("ownerid").Id);
                    Entity user = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                    if (user != null)
                    {
                        arrangerFirstName = user.Contains("firstname") ? user.GetAttributeValue<string>("firstname") : string.Empty;
                        arrangerLastName = user.Contains("lastname") ? user.GetAttributeValue<string>("lastname") : string.Empty;
                        arrangerTitle = user.Contains("title") ? user.GetAttributeValue<string>("title") : string.Empty;
                        arrangerMobile = user.Contains("mobilephone") ? user.GetAttributeValue<string>("mobilephone") : string.Empty;
                        arrangerEmail = user.Contains("internalemailaddress") ? user.GetAttributeValue<string>("internalemailaddress") : string.Empty;
                    }
                }
                #endregion

                string result = CallGetAPI(MYOBAPIBaseURL + "/Contact/Employee?$filter=LastName eq '" + arrangerLastName + "' and FirstName eq '" + arrangerFirstName + "'", AccessToken);
                if (!string.IsNullOrEmpty(result) && !result.Contains("Error:"))
                {
                    EmployeeItems employeeItems = null;
                    using (MemoryStream DeSerializememoryStream = new MemoryStream())
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(EmployeeItems));
                        StreamWriter writer = new StreamWriter(DeSerializememoryStream);
                        writer.Write(result);
                        writer.Flush();

                        DeSerializememoryStream.Position = 0;
                        employeeItems = (EmployeeItems)serializer.ReadObject(DeSerializememoryStream);
                    }
                    if (employeeItems != null && employeeItems.Items != null)
                    {
                        Employee emp = employeeItems.Items.FirstOrDefault();
                        if (emp != null)
                            SalesPersonUID = emp.UID.ToString();
                    }

                }
                else if (!string.IsNullOrEmpty(result))
                {
                    objResponse.Success = false;
                    objResponse.Message = result;
                }
                return SalesPersonUID;
            }

            catch (Exception ex)
            {
                AppendLog("Error occured in FindSalesPersonByName: " + ex.Message);
                objResponse.Success = false;
                objResponse.Message = "Error occured in FindSalesPersonByName: " + ex.Message;
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        private string CreateSalesPerson()
        {
            try
            {
                var myobEmployee = new Employee();
                myobEmployee.FirstName = arrangerFirstName;
                myobEmployee.LastName = arrangerLastName;
                myobEmployee.IsActive = true;
                myobEmployee.IsIndividual = true;
                Address address = new Address();
                address.Phone1 = arrangerMobile;
                address.Email = arrangerEmail;
                List<Address> addList = new List<Address>();
                addList.Add(address);
                myobEmployee.Addresses = addList;
                string jsonString = string.Empty;
                using (var ms = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Employee));
                    serializer.WriteObject(ms, myobEmployee);
                    ms.Position = 0;
                    jsonString = System.Text.Encoding.Default.GetString(ms.ToArray());
                }
                string result = CallPostAPI(jsonString, MYOBAPIBaseURL + "/Contact/Employee", AccessToken);
                if (!string.IsNullOrEmpty(result) && result.ToLower().Contains("error"))
                    objResponse.Message = result;
                return result;
            }
            catch (Exception ex)
            {
                return "Error:" + ex.Message;
            }
        }

        private string FindCustomerByFuneralNumber(string funeralnumber)
        {
            try
            {
                string customerId = string.Empty;
                string result = CallGetAPI(MYOBAPIBaseURL + "/Contact/Customer?$filter=DisplayID eq '" + funeralnumber + "'", AccessToken);
                if (!string.IsNullOrEmpty(result) && !result.Contains("Error:"))
                {
                    CustomerItems customerItems = null;
                    using (MemoryStream DeSerializememoryStream = new MemoryStream())
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(CustomerItems));
                        StreamWriter writer = new StreamWriter(DeSerializememoryStream);
                        writer.Write(result);
                        writer.Flush();
                        DeSerializememoryStream.Position = 0;
                        customerItems = (CustomerItems)serializer.ReadObject(DeSerializememoryStream);
                    }
                    if (customerItems != null && customerItems.Items != null)
                    {
                        Customer cx = customerItems.Items.FirstOrDefault();
                        if (cx != null)
                            customerId = cx.UID.ToString();
                    }
                }
                else if (!string.IsNullOrEmpty(result))
                {
                    objResponse.Success = false;
                    objResponse.Message = result;
                }
                return customerId;
            }
            catch (Exception ex)
            {
                objResponse.Success = false;
                objResponse.Message = ex.Message;
                return string.Empty;
            }
        }

        private string CreateCustomer(TaxCode taxCodeGST, string salesPersonId, Entity postImage)
        {
            try
            {
                Customer MYOBcontact = GenerateContact(postImage, taxCodeGST, salesPersonId);

                string jsonString = string.Empty;
                using (var ms = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Customer));
                    serializer.WriteObject(ms, MYOBcontact);
                    ms.Position = 0;
                    jsonString = System.Text.Encoding.Default.GetString(ms.ToArray());
                }
                string result = CallPostAPI(jsonString, MYOBAPIBaseURL + "/Contact/Customer", AccessToken);
                if (!string.IsNullOrEmpty(result))
                    objResponse.Message = objResponse.Message + " " + result;
                return result;
            }
            catch (Exception ex)
            {
                return "Error:" + ex.Message;
            }
        }

        private Customer GenerateContact(Entity postImage, TaxCode taxCode, string salesPersonId)
        {
            #region Get Account
            Entity account = null;
            if (postImage.Contains("customerid"))
                account = Retrieve(UserType.User, "account", postImage.GetAttributeValue<EntityReference>("customerid").Id, new ColumnSet(true));
            #endregion

            Customer myobContact = null;

            if (account != null)
            {
                myobContact = new Customer();

                //myobContact.CompanyName = companyFile.Name;

                var addresses = new List<Address>();
                var address = new Address();

                address.City = account.Contains("ols_suburbid") ? account.GetAttributeValue<EntityReference>("ols_suburbid").Name : string.Empty;
                address.ContactName = arrangerFirstName + " " + arrangerLastName;
                // address.Country = contact.informant__address2_county;
                address.Email = account.Contains("emailaddress1") ? account.GetAttributeValue<string>("emailaddress1") : string.Empty;
                address.Fax = "";
                address.Location = 1;
                address.Phone1 = account.Contains("address2_telephone1") ? account.GetAttributeValue<string>("address2_telephone1") : string.Empty;
                address.Phone2 = account.Contains("address2_telephone2") ? account.GetAttributeValue<string>("address2_telephone2") : string.Empty;
                address.Phone3 = "";
                address.PostCode = "";
                address.PostCode = account.Contains("ols_billtoaddresspostcodeid") ? account.GetAttributeValue<EntityReference>("ols_billtoaddresspostcodeid").Name : string.Empty;
                address.Salutation = arrangerTitle;
                address.State = account.Contains("ols_stateid") ? account.GetAttributeValue<EntityReference>("ols_stateid").Name : string.Empty;
                var streets = new System.Text.StringBuilder();
                string address1Line1 = account.Contains("address1_line1") ? account.GetAttributeValue<string>("address1_line1") : string.Empty;
                string address1Line2 = account.Contains("address1_line2") ? account.GetAttributeValue<string>("address1_line2") : string.Empty;
                address.Street = streets.AppendLine(address1Line1).Append(address1Line2).ToString();
                address.Website = postImage.Contains("ols_bpaynumberid") ? postImage.GetAttributeValue<EntityReference>("ols_bpaynumberid").Name : string.Empty;
                addresses.Add(address);
                myobContact.Addresses = addresses;
                myobContact.FirstName = account.Contains("ols_givennames") ? account.GetAttributeValue<string>("ols_givennames") : string.Empty;
                myobContact.LastName = account.Contains("ols_familyname") ? account.GetAttributeValue<string>("ols_familyname") : myobContact.FirstName;
                // myobContact.FirstName = contact.new_deceasedfirstname1;
                //myobContact.LastName = contact.new_deceasedsurname;
                myobContact.IsActive = true;
                myobContact.IsIndividual = true;
                myobContact.DisplayID = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty; // cardid
                string name = postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty;
                if (postImage.Contains("ols_serviceplacesessionfrom"))
                    myobContact.Notes = name.TrimEnd() + " on " + string.Format("{0:dd/M/yyyy}", Util.LocalFromUTCUserDateTime(GetService(UserType.User), postImage.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom")));
                else myobContact.Notes = name;

                //selling details
                SellingDetails sellingDetails = new SellingDetails();

                Employee employee = new Employee();
                employee.UID = new Guid(salesPersonId);

                TaxCode objTaxCode = new TaxCode();
                objTaxCode.UID = taxCode.UID;

                FreightTaxCode objFreightTaxCode = new FreightTaxCode();
                objFreightTaxCode.UID = taxCode.UID;

                sellingDetails.SalesPerson = employee;
                sellingDetails.TaxCode = objTaxCode;
                sellingDetails.FreightTaxCode = objFreightTaxCode;
                myobContact.SellingDetails = sellingDetails;
            }
            return myobContact;
        }


        private string CallPostAPI(string data, string url, string accessToken)
        {
            string result = string.Empty;

            try
            {
                var httpRequest = (HttpWebRequest)WebRequest.Create(url);
                httpRequest.Method = "POST";

                httpRequest.Headers["x-myobapi-version"] = "v2";
                httpRequest.Headers["Accept-Encoding"] = "gzip,deflate";
                if (isCloud)
                {
                    httpRequest.Headers["x-myobapi-key"] = _config.MYOBClientId;
                    httpRequest.Headers["Authorization"] = "Bearer " + accessToken;
                }
                else
                    httpRequest.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(_config.MYOBAccount + ":" + _config.MYOBPassword));

                httpRequest.ContentType = "application/json";
                httpRequest.Timeout = 60000;//0.5 Minutes


                using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
                {
                    streamWriter.Write(data);
                }

                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var errorResponse = (HttpWebResponse)wex.Response)
                    {
                        using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                        {
                            string res = reader.ReadToEnd();
                            var Jserializer = new JavaScriptSerializer();
                            APIErrorResponse APIErrorResponse = Jserializer.Deserialize<APIErrorResponse>(res);
                            if (APIErrorResponse != null && APIErrorResponse.Errors != null && APIErrorResponse.Errors.Count > 0)
                            {
                                foreach (var error in APIErrorResponse.Errors)
                                {
                                    if (string.IsNullOrEmpty(result))
                                        result = "Error: " + error.Message + " ";
                                    else
                                        result = result + error.Message + " ";
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        private string CallGetAPI(string url, string accessToken)
        {
            string result = string.Empty;

            try
            {
                var httpRequest = (HttpWebRequest)WebRequest.Create(url);
                httpRequest.Method = "GET";

                httpRequest.Headers["x-myobapi-version"] = "v2";
                //httpRequest.Headers["Accept-Encoding"] = "gzip,deflate";
                //httpRequest.ContentType = "application/json";
                if (isCloud)
                {
                    httpRequest.Headers["x-myobapi-key"] = _config.MYOBClientId;
                    httpRequest.Headers["Authorization"] = "Bearer " + accessToken;
                }
                else
                    httpRequest.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(_config.MYOBAccount + ":" + _config.MYOBPassword));

                httpRequest.Timeout = 60000;//0.5 Minutes

                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var errorResponse = (HttpWebResponse)wex.Response)
                    {
                        using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                        {
                            string res = reader.ReadToEnd();
                            var Jserializer = new JavaScriptSerializer();
                            APIErrorResponse APIErrorResponse = Jserializer.Deserialize<APIErrorResponse>(res);
                            if (APIErrorResponse != null && APIErrorResponse.Errors != null && APIErrorResponse.Errors.Count > 0)
                            {
                                foreach (var error in APIErrorResponse.Errors)
                                {
                                    if (string.IsNullOrEmpty(result))
                                        result = "Error: " + error.Message + " ";
                                    else
                                        result = result + error.Message + " ";
                                }
                            }
                            else if (!string.IsNullOrEmpty(res))
                            {
                                var Jserializer1 = new JavaScriptSerializer();
                                Errors error = Jserializer.Deserialize<Errors>(res);
                                result = "Error: " + error.Message;
                            }
                            else
                                result = "Error: " + wex.Message;
                        }
                    }
                }
            }
            return result;

        }

        private string CallPutAPI(string data, string url, string accessToken)
        {
            string result = string.Empty;

            try
            {
                var httpRequest = (HttpWebRequest)WebRequest.Create(url);
                httpRequest.Method = "PUT";

                httpRequest.Headers["x-myobapi-version"] = "v2";
                httpRequest.Headers["Accept-Encoding"] = "gzip,deflate";
                if (isCloud)
                {
                    httpRequest.Headers["x-myobapi-key"] = _config.MYOBClientId;
                    httpRequest.Headers["Authorization"] = "Bearer " + accessToken;
                }
                else
                    httpRequest.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(_config.MYOBAccount + ":" + _config.MYOBPassword));

                httpRequest.ContentType = "application/json";
                httpRequest.Timeout = 60000;//0.5 Minutes

                using (var streamWriter = new StreamWriter(httpRequest.GetRequestStream()))
                {
                    streamWriter.Write(data);
                }

                var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }

            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var errorResponse = (HttpWebResponse)wex.Response)
                    {
                        using (var reader = new StreamReader(errorResponse.GetResponseStream()))
                        {
                            string res = reader.ReadToEnd();
                            var Jserializer = new JavaScriptSerializer();
                            APIErrorResponse APIErrorResponse = Jserializer.Deserialize<APIErrorResponse>(res);
                            if (APIErrorResponse != null && APIErrorResponse.Errors != null && APIErrorResponse.Errors.Count > 0)
                            {
                                foreach (var error in APIErrorResponse.Errors)
                                {
                                    if (string.IsNullOrEmpty(result))
                                        result = "Error: " + error.Message + " ";
                                    else
                                        result = result + error.Message + " ";
                                }
                            }
                        }
                    }
                }
            }
            return result;
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
