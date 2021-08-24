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
    public class PostCreateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostCreate of Funeral
        /// Creates Childrens, Marriages based on Children & Marriage no. given
        /// Auto creates Music, BDM, Death Certificate, Mortuary Reg and Funeral Operation
        /// Creates different Funeral Products based on different conditions
        /// Updates all Funeral products to prepaid if funeral Prepaid status is Active
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="secureString"></param>

        public PostCreateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        #region [Variables]
        public Money ServiceFees;
        public Money CrematoriumServiceFees;
        public Money CemetaryServiceFees;
        private string Message = string.Empty;
        IOrganizationService svc = null;
        #endregion

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "create" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Opportunity PostCreate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];

                if ((target == null || target.LogicalName != "opportunity" || target.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }
                svc = GetService(UserType.User);
                if (svc == null)
                {
                    AppendLog("Service is null");
                    return;
                }

                #region Translate Fields for Fees
                if (target.Contains("ols_serviceplaceid"))
                    ServiceFees = Util.SetServiceFees(svc, "ols_serviceplaces", target.GetAttributeValue<EntityReference>("ols_serviceplaceid"), target);
                if (target.Contains("ols_cemeteryid"))
                    CemetaryServiceFees = Util.SetServiceFees(svc, "ols_cemetery", target.GetAttributeValue<EntityReference>("ols_cemeteryid"), target);
                if (target.Contains("ols_crematoriumid"))
                    CrematoriumServiceFees = Util.SetServiceFees(svc, "ols_crematorium", target.GetAttributeValue<EntityReference>("ols_crematoriumid"), target);
                #endregion

                CreateChildren(target);
                CreateMarriage(target);
                CreateMusic(target);
                CreateOppProductByOptions(target);
                SetDueDate(target);
                SetPrePaidLock(target);

                #region Create BDM, Death Certificate, Mortuary Reg and Funeral Operation
                string name = target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty;
                string funeralNumber = target.Contains("ols_funeralnumber") ? target.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;

                Entity updateOpportunity = new Entity(target.LogicalName, target.Id);
                updateOpportunity["ols_bdmid"] = new EntityReference("ols_bdm", CreateBDM(name, funeralNumber, target));
                updateOpportunity["ols_deathcertificateid"] = new EntityReference("ols_deathcertificate", CreateDeathCertificate(name, funeralNumber, target));
                Guid mortuaryId = CreateMortuaryReg(name, funeralNumber, target);

                updateOpportunity["ols_mortuaryregisterid"] = new EntityReference("ols_mortuaryregister", mortuaryId);
                updateOpportunity["ols_operationsid"] = new EntityReference("ols_funeraloperations", CreateOperations(name, funeralNumber, target));
                Update(UserType.User, updateOpportunity);
                #endregion

                #region DefaultDoctor
                if (target.Contains("ols_funeraltype") && target.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == 1) //Cremation
                {
                    //EntityReference mortuaryRef = target.Contains("ols_mortuaryregisterid") ? target.GetAttributeValue<EntityReference>("ols_mortuaryregisterid") : null;
                    if (mortuaryId != Guid.Empty)
                    {
                        string doctorname = "Downing, Dr Margaret";
                        Guid medicalrefereeId = GetDoctorIdByName(doctorname);
                        if (medicalrefereeId != Guid.Empty)
                        {
                            Entity updateMortuary = new Entity("ols_mortuaryregister", mortuaryId);
                            updateMortuary["ols_medicalrefereeid"] = new EntityReference("ols_medicalreferee", medicalrefereeId);
                            Update(UserType.User, updateMortuary);
                        }
                    }
                }
                #endregion

                #region Olsens Booking
                if (target.Contains("ols_serviceplacesessionfrom"))
                {
                    Entity olsensBooking = GetOlsensBooking(funeralNumber + "_From");
                    if (olsensBooking != null)
                    {
                        Entity updateOB = new Entity(olsensBooking.LogicalName, olsensBooking.Id);
                        updateOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), target.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom"));
                        Update(UserType.User, updateOB);
                    }
                    else
                    {
                        Entity createOB = new Entity("ols_olsensbooking");
                        createOB["ols_name"] = funeralNumber + "_From";
                        createOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), target.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom"));
                        createOB["ols_funeralid"] = new EntityReference(target.LogicalName, target.Id);
                        Create(UserType.User, createOB);
                    }
                }
                if (target.Contains("ols_serviceplacesessionto"))
                {
                    Entity olsensBooking = GetOlsensBooking(funeralNumber + "_To");
                    if (olsensBooking != null)
                    {
                        Entity updateOB = new Entity(olsensBooking.LogicalName, olsensBooking.Id);
                        updateOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), target.GetAttributeValue<DateTime>("ols_serviceplacesessionto"));
                        Update(UserType.User, updateOB);
                    }
                    else
                    {
                        Entity createOB = new Entity("ols_olsensbooking");
                        createOB["ols_name"] = funeralNumber + "_To";
                        createOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), target.GetAttributeValue<DateTime>("ols_serviceplacesessionto"));
                        createOB["ols_funeralid"] = new EntityReference(target.LogicalName, target.Id);
                        Create(UserType.User, createOB);
                    }
                }
                #endregion

                AppendLog("Opportunity PostCreate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #region [Public Methods]
        public Entity GetOlsensBooking(string name)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_olsensbooking");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_name", ConditionOperator.Equal, name);
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception ex) { throw ex; }
        }
        public void CreateChildren(Entity target)
        {
            try
            {
                int numberOfChildren = target.Contains("ols_numberofchildren") ? target.GetAttributeValue<int>("ols_numberofchildren") : 0;

                if (numberOfChildren > 0)
                {
                    string required = numberOfChildren.ToString();
                    for (int i = 1; i <= numberOfChildren; i++)
                    {
                        Entity createChildren = new Entity("ols_children");
                        createChildren["ols_name"] = "Children of " + (target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty) + "( " + i + " of " + required + " )";
                        createChildren["ols_funeralid"] = new EntityReference("opportunity", target.Id);
                        createChildren["ols_sequencenumber"] = i;
                        createChildren["ols_firstgivenname"] = "Not stated";
                        createChildren["ols_familyname"] = "Not stated";
                        Create(UserType.User, createChildren);
                    }

                }
            }
            catch (Exception e)
            {
                throw e;
            }
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
        public void CreateMarriage(Entity target)
        {
            try
            {
                int numberOfMarriages = target.Contains("ols_numberofadditionalmarriages") ? target.GetAttributeValue<int>("ols_numberofadditionalmarriages") : 0;

                #region Create Current Marriage
                int marriageStatus = target.Contains("ols_marriagestatus") ? target.GetAttributeValue<OptionSetValue>("ols_marriagestatus").Value : 0;
                if (marriageStatus != 4)// Never Married
                {
                    Entity createMarrDetails = new Entity("ols_marriagedetails");
                    if (numberOfMarriages > 0)
                        createMarrDetails["ols_name"] = "Marriage History of " + (target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty) + "( Required " + numberOfMarriages.ToString() + " )";
                    else
                        createMarrDetails["ols_name"] = "Marriage History of " + (target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty) + "( Required None )";
                    createMarrDetails["ols_funeralid"] = new EntityReference("opportunity", target.Id);
                    createMarrDetails["ols_marriagestatus"] = target.GetAttributeValue<OptionSetValue>("ols_marriagestatus");
                    Entity country = GetCountry("Australia");
                    if (country != null)
                        createMarrDetails["ols_countryid"] = new EntityReference(country.LogicalName, country.Id);
                    //createMarrDetails["ols_sequencenumber"] = i;
                    Create(UserType.User, createMarrDetails);
                }
                #endregion


                if (numberOfMarriages > 0)
                {
                    string required = numberOfMarriages.ToString();
                    for (int i = 1; i <= numberOfMarriages; i++)
                    {
                        Entity createMarrDetails = new Entity("ols_marriagedetails");
                        createMarrDetails["ols_name"] = "Marriage History of " + (target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty) + "( " + i + " of required " + required + " )";
                        createMarrDetails["ols_funeralid"] = new EntityReference("opportunity", target.Id);
                        //if (target.Contains("ols_marriagestatus") && i == 1)
                        //    createMarrDetails["ols_marriagestatus"] = target.GetAttributeValue<OptionSetValue>("ols_marriagestatus");
                        createMarrDetails["ols_sequencenumber"] = i;
                        Entity country = GetCountry("Australia");
                        if (country != null)
                            createMarrDetails["ols_countryid"] = new EntityReference(country.LogicalName, country.Id);
                        Create(UserType.User, createMarrDetails);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void CreateMusic(Entity target)
        {
            try
            {
                Entity createMusic = new Entity("ols_music");
                createMusic["ols_name"] = "Music for " + (target.Contains("name") ? target.GetAttributeValue<string>("name") : string.Empty);
                createMusic["ols_funeralid"] = new EntityReference("opportunity", target.Id);
                createMusic["ols_sequencenumber"] = 1;
                Create(UserType.User, createMusic);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #region Opportunity Product Methods
        public void CreateOppProductByOptions(Entity target)
        {
            try
            {
                CreateOppProduct("CH002-2", target.Id);
                if (target.Contains("pricelevelid") && GetPriceListName(target.GetAttributeValue<EntityReference>("pricelevelid").Id) != "Walter Carter Funerals Pty Ltd")
                {
                    CreateOppProduct("CH002-3", target.Id);
                    CreateOppProduct("CH002-4", target.Id);
                }
                if (target.Contains("ols_clergycelebrantid") || (target.Contains("ols_clergycelebrant") && !string.IsNullOrEmpty(target.GetAttributeValue<string>("ols_clergycelebrant"))))
                    CreateForCelebrant(target);
                if (target.Contains("ols_flowers") && target.GetAttributeValue<OptionSetValue>("ols_flowers").Value == Convert.ToInt32(Constants.FlowerStatus.FromHouse))
                    CreateOppProduct("CH008", target.Id); //flowers
                if (target.Contains("ols_carsrequired") && target.Contains("ols_qtycars"))
                    CreateOppProduct("CH010", target.Id, target.GetAttributeValue<int>("ols_qtycars"));
                if (target.Contains("ols_pressnoticeno") && target.GetAttributeValue<int>("ols_pressnoticeno") > 0)
                    for (int i = 0; i < target.GetAttributeValue<int>("ols_pressnoticeno"); i++)
                        CreateOppProduct("CH007", target.Id); //press Notice
                if (target.Contains("ols_funeraltype") && (target.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == Convert.ToInt32(Constants.FuneralType.Cremation) || target.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == Convert.ToInt32(Constants.FuneralType.Burial)))
                    CreateByFuneralType(target);
                if (target.Contains("ols_servicetypeid"))
                {
                    if (target.Contains("pricelevelid") && target.GetAttributeValue<EntityReference>("pricelevelid").Name == "Walter Carter Funerals Pty Ltd")
                    {
                        switch (target.GetAttributeValue<EntityReference>("ols_servicetypeid").Name)
                        {
                            case "Bespoke 1 location - Planning fee":
                                CreateOppProduct("CH-103", target.Id); //Hearse transportation – 1 location
                                break;
                            case "Bespoke 2 locations - Planning fee":
                                CreateOppProduct("CH-104", target.Id); //Hearse transportation – 2 location
                                break;
                            case "Immediate - Service fee":
                                CreateOppProduct("CH-105", target.Id); //Hearse transport
                                CreateOppProduct("CH0293", target.Id); //Ashes collect
                                break;
                            default:
                                break;
                        }
                    }
                    CreateOppProduct(GetProductNoByID(target.GetAttributeValue<EntityReference>("ols_servicetypeid").Id), target.Id);
                }
                if (target.Contains("ols_transferfrom"))
                    CreateOppProduct("CH002", target.Id); // Transfer
                if (target.GetAttributeValue<bool>("ols_ashesdeliveryinurn"))
                    CreateOppProduct("CH025", target.Id); //  Urn
                if (target.GetAttributeValue<bool>("ols_gravemarker"))
                    CreateOppProduct("CH024", target.Id); // Grave Marker 
                if (target.GetAttributeValue<bool>("ols_chapelhire"))
                {
                    if (ServiceFees != null && ServiceFees.Value != 0)
                        CreateOppProduct("CH014", target.Id, ServiceFees); //Chapel Hire
                    else CreateOppProduct("CH014", target.Id);
                }
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in CreateOppProductByOptions: " + ex.Message);
            }
        }
        public Guid CreateOppProduct(String productNumber, Guid OpportunityID)
        {

            Entity Ent = new Entity("opportunityproduct");
            Entity Prod = Util.GetProduct(productNumber, svc);

            try
            {
                if (Prod != null)
                {
                    Ent["productid"] = new EntityReference("product", Prod.Id);
                    if (Prod.Contains("defaultuomid"))
                        Ent["uomid"] = new EntityReference("uom", Prod.GetAttributeValue<EntityReference>("defaultuomid").Id);
                    if (Prod.Contains("price"))
                        Ent["priceperunit"] = Prod.GetAttributeValue<Money>("price");
                    Ent["quantity"] = new decimal(1.0);
                    Ent["opportunityid"] = new EntityReference("opportunity", OpportunityID);
                    Ent["ols_sequencenumber"] = Prod.Contains("ols_sequencenumber") ? Prod.GetAttributeValue<int>("ols_sequencenumber") : 1;
                    return Create(UserType.User, Ent);
                }
                else return Guid.Empty;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Ent = null;
            }
        }
        public Guid CreateOppProduct(String productNumber, Guid OpportunityID, int quantity)
        {
            Entity Ent = new Entity("opportunityproduct");
            Entity Prod = Util.GetProduct(productNumber, svc);
            try
            {
                if (Prod != null)
                {
                    Ent["productid"] = new EntityReference("product", Prod.Id);
                    if (Prod.Contains("defaultuomid"))
                        Ent["uomid"] = new EntityReference("uom", Prod.GetAttributeValue<EntityReference>("defaultuomid").Id);
                    if (Prod.Contains("price"))
                        Ent["priceperunit"] = Prod.GetAttributeValue<Money>("price");
                    Ent["quantity"] = Convert.ToDecimal(quantity);
                    Ent["opportunityid"] = new EntityReference("opportunity", OpportunityID);
                    Ent["ols_sequencenumber"] = Prod.Contains("ols_sequencenumber") ? Prod.GetAttributeValue<int>("ols_sequencenumber") : 1;
                    return Create(UserType.User, Ent);
                }
                else return Guid.Empty;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Ent = null;
            }
        }
        public Guid CreateOppProduct(String productNumber, Guid OpportunityID, Money price)
        {
            Entity Ent = new Entity("opportunityproduct");
            Entity Prod = Util.GetProduct(productNumber, svc);
            try
            {
                if (Prod != null)
                {
                    Ent["productid"] = new EntityReference("product", Prod.Id);
                    if (Prod.Contains("defaultuomid"))
                        Ent["uomid"] = new EntityReference("uom", Prod.GetAttributeValue<EntityReference>("defaultuomid").Id);
                    Ent["ispriceoverridden"] = true;
                    Ent["priceperunit"] = price;
                    Ent["quantity"] = new decimal(1.0);
                    Ent["opportunityid"] = new EntityReference("opportunity", OpportunityID);
                    Ent["ols_sequencenumber"] = Prod.Contains("ols_sequencenumber") ? Prod.GetAttributeValue<int>("ols_sequencenumber") : 1;
                    return Create(UserType.User, Ent);
                }
                else return Guid.Empty;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Ent = null;
            }
        }
        public void CreateForCelebrant(Entity target)
        {
            try
            {
                if (target.Contains("ols_clergycelebrantfee") && (target.GetAttributeValue<OptionSetValue>("ols_clergycelebrantfee").Value == Convert.ToInt32(Constants.CelebrantFees.NoPayment) || target.GetAttributeValue<OptionSetValue>("ols_clergycelebrantfee").Value == Convert.ToInt32(Constants.CelebrantFees.FamilyToPay)))
                    CreateOppProduct("CH006", target.Id, new Money(new decimal(0.0))); //create a Celebrant  but with no fees
                else
                {
                    if (target.GetAttributeValue<bool>("ols_doubleservice"))
                        CreateOppProduct("CH028", target.Id);
                    else
                        CreateOppProduct("CH006", target.Id);
                }
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in CreateForCelebrant: " + ex.Message);
            }
        }
        public void CreateByFuneralType(Entity target)
        {
            CreateOppProduct("CH009", target.Id);    //Doctors Fees-Death Cerficate CH009
            if (target.Contains("ols_funeraltype") && target.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == Convert.ToInt32(Constants.FuneralType.Cremation))
            {
                if (CrematoriumServiceFees != null && CrematoriumServiceFees.Value != 0)
                    CreateOppProduct("CH004", target.Id, CrematoriumServiceFees); //Cremation  Booking 
                else CreateOppProduct("CH004", target.Id);
                if (target.Contains("ols_transferfrom") && (target.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value == Convert.ToInt32(Constants.TransferFrom.Coroner) || target.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value == Convert.ToInt32(Constants.TransferFrom.PlaceOfDeath) || target.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value == Convert.ToInt32(Constants.TransferFrom.PlaceOfResidence)))
                    CreateOppProduct("CH005", target.Id);    //Statutory Documentation for Cremation
            }
            if (target.Contains("ols_funeraltype") && target.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == Convert.ToInt32(Constants.FuneralType.Burial))
            {
                if (CemetaryServiceFees != null && CemetaryServiceFees.Value != 0)
                    CreateOppProduct("CH021", target.Id, CemetaryServiceFees); //burial  Booking  
                else CreateOppProduct("CH021", target.Id);
            }
        }
        #region  Function product By ID 
        public String GetProductNoByID(Guid productId)
        {
            ColumnSet cSet = new ColumnSet(new string[] { "productnumber" });
            Entity ent = null;
            try
            {
                ent = Retrieve(UserType.User, "product", productId, cSet);
                if (ent != null)
                    return ent.Contains("productnumber") ? ent.GetAttributeValue<string>("productnumber") : string.Empty;
                else
                    return null;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                cSet = null;
                ent = null;
            }
        }
        #endregion

        #endregion

        public void SetDueDate(Entity target)
        {
            try
            {
                if (target.Contains("ols_serviceplacesessionfrom"))
                {
                    DateTime dt = target.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom").AddHours(5.5).AddDays(-1);
                    EntityCollection tasksColl = GetAllTask(target.Id);
                    if (tasksColl != null)
                    {
                        foreach (Entity ent in tasksColl.Entities)
                        {
                            ent["scheduledend"] = dt;
                            Update(UserType.User, ent);
                        }
                    }

                    EntityCollection apptColl = GetAllAppt(target.Id);
                    if (apptColl != null)
                    {
                        foreach (Entity ent in apptColl.Entities)
                        {
                            ent["scheduledend"] = dt;
                            Update(UserType.User, ent);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #region  GetAllTask
        public EntityCollection GetAllTask(Guid regardingObjectId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("task");
                qe.ColumnSet = new ColumnSet("activityid", "scheduledend");
                qe.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, regardingObjectId);
                qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                qe.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 2);
                return RetrieveMultiple(UserType.User, qe);
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        #endregion

        #region  GetAllAppt
        public EntityCollection GetAllAppt(Guid regardingObjectId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("appointment");
                qe.ColumnSet = new ColumnSet("activityid", "scheduledend");
                qe.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, regardingObjectId);
                qe.Criteria.AddCondition("statecode", ConditionOperator.Equal, 0);
                qe.Criteria.AddCondition("statuscode", ConditionOperator.Equal, 1);
                return RetrieveMultiple(UserType.User, qe);
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        #endregion

        public void SetPrePaidLock(Entity target)
        {
            try
            {
                if (target.Contains("ols_prepaidstatus") && target.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value == Convert.ToInt32(Constants.PrePaidStatus.Active))
                {
                    EntityCollection oppProdColl = GetOppProduct(target.Id);
                    if (oppProdColl != null)
                    {
                        foreach (Entity ent in oppProdColl.Entities)
                        {
                            ent["ols_prepaid"] = true;
                            ent["ispriceoverridden"] = true;
                            Update(UserType.User, ent);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public EntityCollection GetOppProduct(Guid opportunityid)
        {
            try
            {
                QueryExpression qe = new QueryExpression("opportunityproduct");
                qe.ColumnSet = new ColumnSet("ols_prepaid", "ispriceoverridden");
                qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityid);
                return RetrieveMultiple(UserType.User, qe);
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        public Guid CreateBDM(string name, string funeralNumber, Entity target)
        {
            try
            {
                Entity createBDM = new Entity("ols_bdm");
                createBDM["ols_name"] = name;
                createBDM["ols_funeralnumber"] = funeralNumber;
                createBDM["ols_funeralid"] = new EntityReference("opportunity", target.Id);
                if (target.Contains("ols_status"))
                    createBDM["ols_status"] = target.GetAttributeValue<OptionSetValue>("ols_status");
                if (target.Contains("ols_prepaidnumber"))
                    createBDM["ols_prepaidnumber"] = target.GetAttributeValue<string>("ols_prepaidnumber");
                if (target.Contains("customerid"))
                    createBDM["ols_customerid"] = target.GetAttributeValue<EntityReference>("customerid");
                return Create(UserType.User, createBDM);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public Guid CreateDeathCertificate(string name, string funeralNumber, Entity target)
        {
            try
            {
                Entity createDeathCertificate = new Entity("ols_deathcertificate");
                createDeathCertificate["ols_name"] = name;
                createDeathCertificate["ols_funeralnumber"] = funeralNumber;
                createDeathCertificate["ols_funeralid"] = new EntityReference("opportunity", target.Id);
                if (target.Contains("ols_status"))
                    createDeathCertificate["ols_status"] = target.GetAttributeValue<OptionSetValue>("ols_status");
                if (target.Contains("ols_prepaidnumber"))
                    createDeathCertificate["ols_prepaidnumber"] = target.GetAttributeValue<string>("ols_prepaidnumber");
                if (target.Contains("customerid"))
                    createDeathCertificate["ols_customerid"] = target.GetAttributeValue<EntityReference>("customerid");
                return Create(UserType.User, createDeathCertificate);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public string GetPriceListName(Guid priceLevelId)
        {
            QueryExpression qe = new QueryExpression("pricelevel");
            qe.ColumnSet = new ColumnSet("name");
            qe.Criteria.AddCondition("pricelevelid", ConditionOperator.Equal, priceLevelId);
            Entity priceLevel = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            if (priceLevel != null)
            {
                return priceLevel.Contains("name") ? priceLevel.GetAttributeValue<string>("name") : string.Empty;
            }
            else
                return string.Empty;
        }
        public Guid CreateMortuaryReg(string name, string funeralNumber, Entity target)
        {
            try
            {
                Entity createMortuaryRegister = new Entity("ols_mortuaryregister");
                createMortuaryRegister["ols_name"] = name;
                createMortuaryRegister["ols_funeralnumber"] = funeralNumber;
                createMortuaryRegister["ols_funeralid"] = new EntityReference(target.LogicalName, target.Id);
                return Create(UserType.User, createMortuaryRegister);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public Guid CreateOperations(string name, string funeralNumber, Entity target)
        {
            try
            {
                Entity createOperations = new Entity("ols_funeraloperations");
                createOperations["ols_name"] = name;
                createOperations["ols_funeralnumber"] = funeralNumber;
                createOperations["ols_funeralid"] = new EntityReference(target.LogicalName, target.Id);

                return Create(UserType.User, createOperations);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public Entity GetCountry(string name)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_country");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_name", ConditionOperator.Equal, name);
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion
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
