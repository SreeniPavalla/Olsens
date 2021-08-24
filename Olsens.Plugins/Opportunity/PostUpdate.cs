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
    public class PostUpdateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostUpdate of Funeral
        /// If LongStay is Yes, updating operations Embalming Type to Full, Creating Appointment
        /// On Change of Brand- Making BPayNumber Free, Updating BPNumberAccount, Updating BPNumberFuneral
        /// On Change of Informant- Updating account Bpay Number
        /// Updating Childs, Marriages based on count
        /// Calculating Child age
        /// On change of deceased fullname, Updating name in MortuaryReg, Operations, Music
        /// Updating all Activity tasks & Appointments scheduleEnd Date based on Funeral Service Place session date
        /// Deleting, Updating, Creating funeral products based on different options change
        /// Updating all Funeral products to Prepaid if funeral Prepaid status is Active
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>

        public PostUpdateWrapper(string unsecConfig, string securestring) : base(unsecConfig, securestring) { }

        #region [Variables]
        public Money ServiceFees;
        public Money CrematoriumServiceFees;
        public Money CemetaryServiceFees;
        private string Message = string.Empty;
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

                if (Context.Depth > 1) { return; }

                IOrganizationService svc = GetService(UserType.User);
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

                if (preImage.GetAttributeValue<bool>("ols_longstay") != postImage.GetAttributeValue<bool>("ols_longstay") && postImage.GetAttributeValue<bool>("ols_longstay"))
                {
                    UpdateOperations(postImage);
                    CreateAppointment(target, postImage);
                }

                Guid preData_BrandId = preImage.Contains("pricelevelid") ? preImage.GetAttributeValue<EntityReference>("pricelevelid").Id : Guid.Empty;
                Guid postData_BrandId = postImage.Contains("pricelevelid") ? postImage.GetAttributeValue<EntityReference>("pricelevelid").Id : Guid.Empty;
                if (!preData_BrandId.Equals(postData_BrandId))
                {
                    ChangeBrand(preImage, postImage);
                }

                Guid preData_CustomerId = preImage.Contains("customerid") ? preImage.GetAttributeValue<EntityReference>("customerid").Id : Guid.Empty;
                Guid postData_CustmerId = postImage.Contains("customerid") ? postImage.GetAttributeValue<EntityReference>("customerid").Id : Guid.Empty;
                if (!preData_CustomerId.Equals(postData_CustmerId) && preImage.Contains("customerid"))
                {
                    UpdateBPNumberAccount(preImage.GetAttributeValue<EntityReference>("customerid"), postImage);
                }

                UpdateChild(preImage, postImage);
                UpdateMarriageDetails(preImage, postImage);
                CalcChildAge(preImage, postImage);
                UpdateMortuaryReg(preImage, postImage);
                UpdateOperations(preImage, postImage);
                UpdateBDM(preImage, postImage);
                UpdateDeathCertificate(preImage, postImage);
                UpdateMusic(preImage, postImage);
                SetDueDate(preImage, postImage);

                int status = postImage.Contains("ols_status") ? postImage.GetAttributeValue<OptionSetValue>("ols_status").Value : 0;
                int prePaidstatus = postImage.Contains("ols_prepaidstatus") ? postImage.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value : 0;
                if (status == (int)Constants.FuneralStatus.PrePaid && prePaidstatus == (int)Constants.PrePaidStatus.AtNeed) return;

                ApplyOperations(preImage, postImage);
                SetPrePaidLock(preImage, postImage);

                #region DefaultDoctor
                if (target.Contains("ols_funeraltype") && target.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == 1) //Cremation
                {
                    EntityReference mortuaryRef = postImage.Contains("ols_mortuaryregisterid") ? postImage.GetAttributeValue<EntityReference>("ols_mortuaryregisterid") : null;
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

                #region Olsens Booking
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
                DateTime? preData_ServicePlaceSessionFrom = preImage.Contains("ols_serviceplacesessionfrom") ? preImage.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom") : (DateTime?)null;
                DateTime? postData_ServicePlaceSessionFrom = postImage.Contains("ols_serviceplacesessionfrom") ? postImage.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom") : (DateTime?)null;
                DateTime? preData_ServicePlaceSessionTo = preImage.Contains("ols_serviceplacesessionto") ? preImage.GetAttributeValue<DateTime>("ols_serviceplacesessionto") : (DateTime?)null;
                DateTime? postData_ServicePlaceSessionTo = postImage.Contains("ols_serviceplacesessionto") ? postImage.GetAttributeValue<DateTime>("ols_serviceplacesessionto") : (DateTime?)null;
                if (preData_ServicePlaceSessionFrom != postData_ServicePlaceSessionFrom)
                {
                    Entity olsensBooking = GetOlsensBooking(funeralNumber + "_From");
                    if (postData_ServicePlaceSessionFrom != null)
                    {
                        if (olsensBooking != null)
                        {
                            Entity updateOB = new Entity(olsensBooking.LogicalName, olsensBooking.Id);
                            updateOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), (DateTime)postData_ServicePlaceSessionFrom);
                            Update(UserType.User, updateOB);
                        }
                        else
                        {
                            Entity createOB = new Entity("ols_olsensbooking");
                            createOB["ols_name"] = funeralNumber + "_From";
                            createOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), (DateTime)postData_ServicePlaceSessionFrom);
                            createOB["ols_funeralid"] = new EntityReference(target.LogicalName, target.Id);
                            Create(UserType.User, createOB);
                        }
                    }
                    else if (olsensBooking != null)
                        Delete(UserType.User, olsensBooking.LogicalName, olsensBooking.Id);
                }
                if (preData_ServicePlaceSessionTo != postData_ServicePlaceSessionTo)
                {
                    Entity olsensBooking = GetOlsensBooking(funeralNumber + "_To");
                    if (postData_ServicePlaceSessionTo != null)
                    {
                        if (olsensBooking != null)
                        {
                            Entity updateOB = new Entity(olsensBooking.LogicalName, olsensBooking.Id);
                            updateOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), (DateTime)postData_ServicePlaceSessionTo);
                            Update(UserType.User, updateOB);
                        }
                        else
                        {
                            Entity createOB = new Entity("ols_olsensbooking");
                            createOB["ols_name"] = funeralNumber + "_To";
                            createOB["ols_servicedate"] = Util.LocalFromUTCUserDateTime(GetService(UserType.User), (DateTime)postData_ServicePlaceSessionTo);
                            createOB["ols_funeralid"] = new EntityReference(target.LogicalName, target.Id);
                            Create(UserType.User, createOB);
                        }
                    }
                    else if (olsensBooking != null)
                        Delete(UserType.User, olsensBooking.LogicalName, olsensBooking.Id);
                }
                #endregion

                AppendLog("Opportunity PostUpdate - Plugin Excecution is Completed.");
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
        public void UpdateOperations(Entity postImage)
        {
            if (postImage.Contains("ols_operationsid"))
            {
                try
                {
                    Entity operation = new Entity("ols_funeraloperations");
                    operation.Id = postImage.GetAttributeValue<EntityReference>("ols_operationsid").Id;
                    operation.LogicalName = postImage.GetAttributeValue<EntityReference>("ols_operationsid").LogicalName;
                    operation["ols_embalmingtype"] = new OptionSetValue(2); //Full
                    Update(UserType.User, operation);
                }
                catch (Exception ex)
                {
                    AppendLog("Error occured in UpdateOperations: " + ex.Message);
                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }
        }
        public void CreateAppointment(Entity target, Entity postImage)
        {
            Entity Ent = new Entity("appointment");

            try
            {
                Ent["subject"] = "LongStay_" + postImage.GetAttributeValue<string>("ols_funeralnumber") + "_" + postImage.GetAttributeValue<string>("name");

                DateTime createdOn = postImage.GetAttributeValue<DateTime>("createdon");
                var potentialDate = createdOn.AddDays(7);
                if (potentialDate.DayOfWeek == DayOfWeek.Saturday)
                    potentialDate = createdOn.AddDays(9);
                else if (potentialDate.DayOfWeek == DayOfWeek.Sunday)
                    potentialDate = createdOn.AddDays(8);

                Ent["scheduledstart"] = potentialDate.ToUniversalTime();

                Ent["scheduledend"] = potentialDate.AddHours(1).ToUniversalTime();
                Ent["regardingobjectid"] = new EntityReference("opportunity", target.Id);
                var party = new Entity("activityparty");
                var user = GetUserByDomainName();
                if (user != null)
                {
                    party["partyid"] = new EntityReference("systemuser", user.Id);
                    Ent["requiredattendees"] = new Entity[] { party };
                }
                Create(UserType.User, Ent);
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
        public Entity GetUserByDomainName()
        {
            try
            {
                QueryExpression qe = new QueryExpression("systemuser");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("domainname", ConditionOperator.Equal, Util.GetConfigSettingValue("LongStayAttendee", GetService(UserType.User)));
                qe.Criteria.AddCondition("isdisabled", ConditionOperator.Equal, false);
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception ex) { return null; }
        }
        public void ChangeBrand(Entity preImage, Entity postImage)
        {
            if (preImage.Contains("ols_bpaynumberid"))
                MakeBPNumberFree(preImage.GetAttributeValue<EntityReference>("ols_bpaynumberid").Id);
            Entity bpyNumber = GetBPayNumberEntity(postImage);
            if (bpyNumber == null) return;

            UpdateBPNumberEntity(postImage, bpyNumber);

            UpdateBPNumberAccount(postImage, bpyNumber.Id);
            // this.new_bpaynumberid = new EntityReference("new_bpaynumber", bpaynumberid.Id);

            UpdateBPNumberFuneral(bpyNumber, postImage);
        }
        public void MakeBPNumberFree(Guid bPayNumberId)
        {
            try
            {
                Entity updateBPayNumber = new Entity("ols_bpaynumber", bPayNumberId);
                updateBPayNumber["ols_funeralnumber"] = null;
                Update(UserType.User, updateBPayNumber);
            }
            catch { }
        }
        public Entity GetBPayNumberEntity(Entity postImage)
        {
            var bpaynumberEnt = new Entity();
            // ConfigData config = new ConfigData(svc);

            if (!postImage.Contains("ols_bpaynumberid"))
            {
                QueryExpression qe = new QueryExpression("ols_bpaynumber");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralnumber", ConditionOperator.Null);
                if (postImage.Contains("pricelevelid"))
                    qe.Criteria.AddCondition("ols_brand", ConditionOperator.Equal, Util.GetBrandValue(GetService(UserType.User), postImage.GetAttributeValue<EntityReference>("pricelevelid").Id));
                else
                    qe.Criteria.AddCondition("ols_brand", ConditionOperator.Equal, 0); // HN Olsen Funerals Pty Ltd
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            else
            {
                QueryExpression qe = new QueryExpression("ols_bpaynumber");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralnumber", ConditionOperator.Equal, postImage.GetAttributeValue<string>("ols_funeralnumber"));
                if (postImage.Contains("pricelevelid"))
                    qe.Criteria.AddCondition("ols_brand", ConditionOperator.Equal, Util.GetBrandValue(GetService(UserType.User), postImage.GetAttributeValue<EntityReference>("pricelevelid").Id));
                else
                    qe.Criteria.AddCondition("ols_brand", ConditionOperator.Equal, 0); // HN Olsen Funerals Pty Ltd
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
        }
        public void UpdateBPNumberEntity(Entity postImage, Entity bPayNumber)
        {
            try
            {
                if (!bPayNumber.Contains("ols_funeralnumber"))
                {
                    bPayNumber["ols_funeralnumber"] = postImage.GetAttributeValue<string>("ols_funeralnumber");
                    Update(UserType.User, bPayNumber);
                }
            }
            catch { throw new Exception("Cannot update bpnumber entity."); }
        }
        public void UpdateBPNumberAccount(Entity postImage, Guid bPayNumberId)
        {
            try
            {
                if (postImage.Contains("customerid"))
                {
                    Entity updateAccount = new Entity("account", postImage.GetAttributeValue<EntityReference>("customerid").Id);
                    updateAccount["ols_bpaynumberid"] = new EntityReference("ols_bpaynumber", bPayNumberId);
                    updateAccount["ols_funeralnumber"] = postImage.GetAttributeValue<string>("ols_funeralnumber");
                    Update(UserType.User, updateAccount);
                }
            }
            catch { throw new Exception("Cannot update Informant's funeralnumber/bpaynumber."); }

        }
        public void UpdateBPNumberFuneral(Entity bPayNumber, Entity postImage)
        {
            try
            {
                Entity updateFuneral = new Entity("opportunity", postImage.Id);
                updateFuneral["ols_bpaynumberid"] = new EntityReference(bPayNumber.LogicalName, bPayNumber.Id);
                Update(UserType.User, updateFuneral);
            }
            catch { throw new Exception("Cannot update Funeral's bpaynumber."); }
        }
        public void UpdateBPNumberAccount(EntityReference oldCustomerId, Entity postImage)
        {
            var bpayNumberEnt = GetBPayNumberEntity(postImage);
            if (bpayNumberEnt != null)
            {
                ClearAccount(oldCustomerId);
                UpdateBPNumberAccount(postImage, bpayNumberEnt.Id);
            }
        }
        public void ClearAccount(EntityReference oldCustomerId)
        {
            if (oldCustomerId != null)
            {
                Entity updateAccount = new Entity(oldCustomerId.LogicalName, oldCustomerId.Id);
                updateAccount["ols_bpaynumberid"] = null;
                updateAccount["ols_funeralnumber"] = null;
                Update(UserType.User, updateAccount);
            }
        }
        public void UpdateChild(Entity preImage, Entity postImage)
        {
            try
            {
                if (!postImage.Contains("ols_childcountupdatefromjs") || (postImage.Contains("ols_childcountupdatefromjs") && !postImage.GetAttributeValue<bool>("ols_childcountupdatefromjs")))
                {
                    int preData_numberOfChildren = preImage.Contains("ols_numberofchildren") ? preImage.GetAttributeValue<int>("ols_numberofchildren") : 0;
                    int postData_numberOfChildren = postImage.Contains("ols_numberofchildren") ? postImage.GetAttributeValue<int>("ols_numberofchildren") : 0;

                    if (preData_numberOfChildren < postData_numberOfChildren)
                    {
                        int extraRequired = postData_numberOfChildren - preData_numberOfChildren;
                        string required = postData_numberOfChildren.ToString();
                        for (int i = 1; i <= extraRequired; i++)
                        {
                            int count = preData_numberOfChildren + i;
                            Entity createChildren = new Entity("ols_children");
                            createChildren["ols_name"] = "Children of " + (postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty) + "( " + count + " of " + required + " )";
                            createChildren["ols_funeralid"] = new EntityReference("opportunity", postImage.Id);
                            createChildren["ols_firstgivenname"] = "Not stated";
                            createChildren["ols_familyname"] = "Not stated";
                            createChildren["ols_sequencenumber"] = count;
                            Create(UserType.User, createChildren);
                        }

                    }
                    if (preData_numberOfChildren > postData_numberOfChildren)
                    {
                        int recordsToDelete = preData_numberOfChildren - postData_numberOfChildren;
                        EntityCollection childColl = GetChildrens(postImage.Id);
                        if (childColl != null && childColl.Entities.Count > 0)
                        {
                            if (recordsToDelete > childColl.Entities.Count)
                                recordsToDelete = childColl.Entities.Count;
                            List<Entity> childCollDscList = childColl.Entities.OrderByDescending(a => a.GetAttributeValue<int>("ols_sequencenumber")).ToList();
                            if (childCollDscList != null && childCollDscList.Count > 0)
                            {
                                for (int i = 0; i < recordsToDelete; i++)
                                {
                                    Delete(UserType.User, childCollDscList[i].LogicalName, childCollDscList[i].Id);
                                }
                            }
                        }
                    }

                    if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") || preData_numberOfChildren != postData_numberOfChildren)
                    {
                        if (postData_numberOfChildren > 0)
                        {
                            EntityCollection childColl = GetChildrens(postImage.Id);
                            if (childColl != null && childColl.Entities.Count > 0)
                            {
                                string required = postData_numberOfChildren.ToString();
                                for (int i = 1; i <= childColl.Entities.Count; i++)
                                {
                                    Entity updateChildren = new Entity("ols_children", childColl.Entities[i - 1].Id);
                                    updateChildren["ols_name"] = "Children of " + (postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty) + "( " + i + " of " + required + " )";
                                    Update(UserType.User, updateChildren);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void UpdateMarriageDetails(Entity preImage, Entity postImage)
        {
            try
            {
                int preData_MarriageStatus = preImage.Contains("ols_marriagestatus") ? preImage.GetAttributeValue<OptionSetValue>("ols_marriagestatus").Value : -1;
                int postData_MarriageStatus = postImage.Contains("ols_marriagestatus") ? postImage.GetAttributeValue<OptionSetValue>("ols_marriagestatus").Value : -1;

                int preData_numberOfMarriage = preImage.Contains("ols_numberofadditionalmarriages") ? preImage.GetAttributeValue<int>("ols_numberofadditionalmarriages") : 0;
                int postData_numberOfMarriage = postImage.Contains("ols_numberofadditionalmarriages") ? postImage.GetAttributeValue<int>("ols_numberofadditionalmarriages") : 0;

                if (!postImage.Contains("ols_marriagecountupdatefromjs") || (postImage.Contains("ols_marriagecountupdatefromjs") && !postImage.GetAttributeValue<bool>("ols_marriagecountupdatefromjs")))
                {

                    if (preData_numberOfMarriage < postData_numberOfMarriage)
                    {
                        int extraRequired = postData_numberOfMarriage - preData_numberOfMarriage;

                        //if (preData_numberOfMarriage == 1)
                        //    extraRequired = postData_numberOfMarriage;

                        string required = postData_numberOfMarriage.ToString();
                        for (int i = 1; i <= extraRequired; i++)
                        {
                            int count = preData_numberOfMarriage + i;
                            Entity createMarrDetails = new Entity("ols_marriagedetails");
                            createMarrDetails["ols_name"] = "Marriage History of " + (postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty) + "( " + count + " of required " + required + " )";
                            createMarrDetails["ols_funeralid"] = new EntityReference("opportunity", postImage.Id);
                            //if (postImage.Contains("ols_marriagestatus") && preData_numberOfMarriage + i == 1)
                            //    createMarrDetails["ols_marriagestatus"] = postImage.GetAttributeValue<OptionSetValue>("ols_marriagestatus");
                            createMarrDetails["ols_sequencenumber"] = count;
                            //createMarrDetails["ols_firstgivenname"] = "Not stated";
                            Entity country = GetCountry("Australia");
                            if (country != null)
                                createMarrDetails["ols_countryid"] = new EntityReference(country.LogicalName, country.Id);
                            Create(UserType.User, createMarrDetails);
                        }
                    }
                    if (preData_numberOfMarriage > postData_numberOfMarriage)
                    {
                        int recordsToDelete = preData_numberOfMarriage - postData_numberOfMarriage;
                        EntityCollection marriageColl = GetMarriages(postImage.Id);
                        if (marriageColl != null && marriageColl.Entities.Count > 0)
                        {
                            if (recordsToDelete > marriageColl.Entities.Count)
                                recordsToDelete = marriageColl.Entities.Count;
                            List<Entity> marriageCollDscList = marriageColl.Entities.OrderByDescending(a => a.GetAttributeValue<int>("ols_sequencenumber")).ToList();
                            if (marriageCollDscList != null && marriageCollDscList.Count > 0)
                            {
                                for (int i = 0; i < recordsToDelete; i++)
                                {
                                    Delete(UserType.User, marriageCollDscList[i].LogicalName, marriageCollDscList[i].Id);
                                }
                            }
                        }
                    }

                    if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") || preData_numberOfMarriage != postData_numberOfMarriage)
                    {
                        if (postData_numberOfMarriage > 0)
                        {
                            EntityCollection marriageColl = GetMarriages(postImage.Id);
                            if (marriageColl != null && marriageColl.Entities.Count > 0)
                            {
                                string required = postData_numberOfMarriage.ToString();
                                for (int i = 1; i <= marriageColl.Entities.Count; i++)
                                {
                                    Entity updateMarrDetails = new Entity("ols_marriagedetails", marriageColl[i - 1].Id);
                                    if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") || preData_numberOfMarriage != postData_numberOfMarriage)
                                        updateMarrDetails["ols_name"] = "Marriage History of " + (postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty) + "( " + i + " of required " + required + " )";
                                    //if (preData_MarriageStatus != postData_MarriageStatus && postImage.Contains("ols_marriagestatus") && i == 1)
                                    //    updateMarrDetails["ols_marriagestatus"] = postImage.GetAttributeValue<OptionSetValue>("ols_marriagestatus");
                                    Update(UserType.User, updateMarrDetails);
                                }
                            }
                        }
                    }
                }

                if (preData_MarriageStatus != postData_MarriageStatus)
                {
                    Entity currentMarriage = GetCurrentMarriage(preImage.Id);
                    if (currentMarriage != null)
                    {
                        Entity updateMarrDetails = new Entity("ols_marriagedetails", currentMarriage.Id);
                        if (postImage.Contains("ols_marriagestatus"))
                            updateMarrDetails["ols_marriagestatus"] = postImage.GetAttributeValue<OptionSetValue>("ols_marriagestatus");
                        else
                            updateMarrDetails["ols_marriagestatus"] = null;

                        Update(UserType.User, updateMarrDetails);
                    }

                }
                if (preData_numberOfMarriage != postData_numberOfMarriage)
                {
                    Entity currentMarriage = GetCurrentMarriage(preImage.Id);
                    if (currentMarriage != null)
                    {
                        Entity updateMarrDetails = new Entity("ols_marriagedetails", currentMarriage.Id);
                        if (postData_numberOfMarriage > 0)
                            updateMarrDetails["ols_name"] = "Marriage History of " + (postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty) + "( Required " + postData_numberOfMarriage.ToString() + " )";
                        else
                            updateMarrDetails["ols_name"] = "Marriage History of " + (postImage.Contains("name") ? postImage.GetAttributeValue<string>("name") : string.Empty) + "( Required None )";
                        Update(UserType.User, updateMarrDetails);
                    }

                }
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

        #region GetChild
        public EntityCollection GetChildrens(Guid oppId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_children");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, oppId);
                qe.Orders.Add(new OrderExpression("ols_sequencenumber", OrderType.Ascending));
                return RetrieveMultiple(UserType.User, qe);
            }
            catch { return null; }

        }
        #endregion

        #region GetMarriages
        public EntityCollection GetMarriages(Guid oppId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_marriagedetails");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, oppId);
                qe.Criteria.AddCondition("ols_sequencenumber", ConditionOperator.NotNull);
                qe.Orders.Add(new OrderExpression("ols_sequencenumber", OrderType.Ascending));
                return RetrieveMultiple(UserType.User, qe);
            }
            catch (Exception ex) { throw ex; }

        }
        public Entity GetCurrentMarriage(Guid oppId)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_marriagedetails");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, oppId);
                qe.Criteria.AddCondition("ols_sequencenumber", ConditionOperator.Null);
                qe.Orders.Add(new OrderExpression("createdon", OrderType.Ascending));
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception ex) { throw ex; }

        }
        #endregion

        #region CalcChildAge
        public void CalcChildAge(Entity preImage, Entity postImage)
        {
            try
            {
                DateTime? predata_dod = preImage.Contains("ols_dateofdeath") ? preImage.GetAttributeValue<DateTime>("ols_dateofdeath") : (DateTime?)null;
                DateTime? postdata_dod = postImage.Contains("ols_dateofdeath") ? postImage.GetAttributeValue<DateTime>("ols_dateofdeath") : (DateTime?)null;
                DateTime? predata_dodTo = preImage.Contains("ols_dateofdeathto") ? preImage.GetAttributeValue<DateTime>("ols_dateofdeathto") : (DateTime?)null;
                DateTime? postdata_dodTo = postImage.Contains("ols_dateofdeathto") ? postImage.GetAttributeValue<DateTime>("ols_dateofdeathto") : (DateTime?)null;
                if (predata_dod == postdata_dod && predata_dodTo == postdata_dodTo) return;

                postdata_dod = postdata_dod != null ? postdata_dod : postdata_dodTo;

                EntityCollection childColl = GetChildrens(postImage.Id);
                if (childColl != null && childColl.Entities.Count > 0)
                {
                    foreach (var child in childColl.Entities)
                    {
                        if (child.Contains("ols_dateofbirth"))
                        {
                            if ((child.Contains("ols_lifestatus") && child.GetAttributeValue<OptionSetValue>("ols_lifestatus").Value != 1) || !child.Contains("ols_lifestatus"))
                            {
                                Entity updateChild = new Entity(child.LogicalName, child.Id);
                                updateChild["ols_age"] = null;
                                updateChild["ols_ageunit"] = null;
                                Update(UserType.User, updateChild);
                            }
                            else
                            {
                                var age = CalcAge(child.GetAttributeValue<DateTime>("ols_dateofbirth"), (DateTime)postdata_dod);
                                Entity updateChild = new Entity(child.LogicalName, child.Id);
                                updateChild["ols_age"] = Convert.ToString(age.Value);
                                updateChild["ols_ageunit"] = new OptionSetValue(age.UnitValue);
                                Update(UserType.User, updateChild);

                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public Age CalcAge(DateTime startDate, DateTime EndDate)
        {

            var dateDiff = new DateDiff(startDate, EndDate);
            var years = dateDiff.inYears;

            if (years < 0) return new Age(0, "years");
            var months = dateDiff.inMonths;

            var weeks = dateDiff.inWeeks;

            var days = dateDiff.inDays;

            if (years >= 1)
                return new Age(years, "years");
            else
               if (months >= 1) return new Age(months, "months");

            else
           if (weeks >= 1)
                return new Age(weeks, "weeks");
            if (days >= 1) return new Age(days, "days");
            else return new Age(0, "hours");
        }

        #endregion

        #region UpdateMortuaryReg
        public void UpdateMortuaryReg(Entity preImage, Entity postImage)
        {
            try
            {
                if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") && postImage.Contains("ols_mortuaryregisterid"))
                {
                    Entity updateMortuary = new Entity("ols_mortuaryregister", postImage.GetAttributeValue<EntityReference>("ols_mortuaryregisterid").Id);
                    updateMortuary["ols_name"] = postImage.GetAttributeValue<string>("name");
                    Update(UserType.User, updateMortuary);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region UpdateOperations
        public void UpdateOperations(Entity preImage, Entity postImage)
        {
            try
            {
                if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") && postImage.Contains("ols_operationsid"))
                {
                    Entity updateOperation = new Entity("ols_funeraloperations", postImage.GetAttributeValue<EntityReference>("ols_operationsid").Id);
                    updateOperation["ols_name"] = postImage.GetAttributeValue<string>("name");
                    Update(UserType.User, updateOperation);
                }
            }
            catch (Exception ex) { throw ex; }
        }

        #endregion

        #region UpdateBDM
        public void UpdateBDM(Entity preImage, Entity postImage)
        {
            try
            {
                if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") && postImage.Contains("ols_bdmid"))
                {
                    Entity updateBDM = new Entity("ols_bdm", postImage.GetAttributeValue<EntityReference>("ols_bdmid").Id);
                    updateBDM["ols_name"] = postImage.GetAttributeValue<string>("name");
                    Update(UserType.User, updateBDM);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region UpdateDeathCertificate
        public void UpdateDeathCertificate(Entity preImage, Entity postImage)
        {
            try
            {
                if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") && postImage.Contains("ols_deathcertificateid"))
                {
                    Entity UpdateDeathCertificate = new Entity("ols_deathcertificate", postImage.GetAttributeValue<EntityReference>("ols_deathcertificateid").Id);
                    UpdateDeathCertificate["ols_name"] = postImage.GetAttributeValue<string>("name");
                    Update(UserType.User, UpdateDeathCertificate);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

        #region Update Music
        public void UpdateMusic(Entity preImage, Entity postImage)
        {
            try
            {
                if (preImage.GetAttributeValue<string>("name") != postImage.GetAttributeValue<string>("name") && postImage.Contains("ols_operationsid"))
                {
                    QueryExpression qe = new QueryExpression("ols_music");
                    qe.ColumnSet = new ColumnSet(true);
                    qe.Criteria.AddCondition("ols_funeralid", ConditionOperator.Equal, postImage.Id);
                    EntityCollection musicColl = RetrieveMultiple(UserType.User, qe);
                    if (musicColl != null)
                    {
                        foreach (var music in musicColl.Entities)
                        {
                            Entity updateMusic = new Entity("ols_music", music.Id);
                            updateMusic["ols_name"] = postImage.GetAttributeValue<string>("name");
                            Update(UserType.User, updateMusic);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        #endregion

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

        #region  ApplyOperations Methods
        public void ApplyOperations(Entity preImage, Entity postImage)
        {
            try
            {
                DeleteByOptions(preImage, postImage);

                CreateByOptions(preImage, postImage);

                UpdateByOptions(preImage, postImage);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void DeleteOppProduct(string ProductNum, Guid OpportunityID)
        {
            try
            {
                Entity prod = Util.GetProduct(ProductNum, GetService(UserType.User));
                if (prod != null)
                {
                    Entity oppProduct = GetOppProductId(prod.Id, OpportunityID);
                    if (oppProduct != null)
                    {
                        if (oppProduct.Contains("ols_prepaid") && oppProduct.GetAttributeValue<bool>("ols_prepaid"))
                        {
                            AppendLog("Cannot delete a Pre Paid Item.");
                            throw new InvalidPluginExecutionException("Cannot delete a Pre Paid Item.");
                        }
                        else
                            Delete(UserType.User, oppProduct.LogicalName, oppProduct.Id);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public Entity GetOppProductId(Guid productid, Guid opportunityid)
        {
            try
            {
                QueryExpression qe = new QueryExpression("opportunityproduct");
                qe.ColumnSet = new ColumnSet("ols_prepaid", "ispriceoverridden");
                qe.Criteria.AddCondition("productid", ConditionOperator.Equal, productid);
                qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityid);
                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        public void DeleteByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                Guid preData_ClergyCelebrantId = preImage.Contains("ols_clergycelebrantid") ? preImage.GetAttributeValue<EntityReference>("ols_clergycelebrantid").Id : Guid.Empty;
                Guid postData_ClergyCelebrantId = postImage.Contains("ols_clergycelebrantid") ? postImage.GetAttributeValue<EntityReference>("ols_clergycelebrantid").Id : Guid.Empty;
                if (preData_ClergyCelebrantId != Guid.Empty && !preData_ClergyCelebrantId.Equals(postData_ClergyCelebrantId))
                {
                    DeleteOppProduct("CH006", preImage.Id); //Celebrant
                }


                string preData_ClergyCelebrant = preImage.Contains("ols_clergycelebrant") ? preImage.GetAttributeValue<string>("ols_clergycelebrant") : string.Empty;
                string postData_ClergyCelebrant = postImage.Contains("ols_clergycelebrant") ? postImage.GetAttributeValue<string>("ols_clergycelebrant") : string.Empty;
                if ((preData_ClergyCelebrant != string.Empty) && (postData_ClergyCelebrant != preData_ClergyCelebrant))
                {
                    DeleteOppProduct("CH006", preImage.Id); //Celebrant
                }

                int preData_flowerStatus = preImage.Contains("ols_flowers") ? preImage.GetAttributeValue<OptionSetValue>("ols_flowers").Value : 0;
                int postData_flowerStatus = postImage.Contains("ols_flowers") ? postImage.GetAttributeValue<OptionSetValue>("ols_flowers").Value : 0;
                if ((preData_flowerStatus == Convert.ToInt32(Constants.FlowerStatus.FromHouse)) && (preData_flowerStatus != postData_flowerStatus))
                {
                    DeleteOppProduct("CH008", postImage.Id); //flowers
                }

                int preData_CarsRequired = preImage.Contains("ols_carsrequired") ? preImage.GetAttributeValue<OptionSetValue>("ols_carsrequired").Value : 0;
                int postData_CarsRequired = postImage.Contains("ols_carsrequired") ? postImage.GetAttributeValue<OptionSetValue>("ols_carsrequired").Value : 0;
                if ((preData_CarsRequired != 0) && (preData_CarsRequired != postData_CarsRequired))
                {
                    DeleteOppProduct("CH010", postImage.Id); //Cars Required 
                }

                int preData_FuneralType = preImage.Contains("ols_funeraltype") ? preImage.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value : 0;
                int postData_FuneralType = postImage.Contains("ols_funeraltype") ? postImage.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value : 0;

                int preData_TransferFrom = preImage.Contains("ols_transferfrom") ? preImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value : 0;
                int postData_TransferFrom = postImage.Contains("ols_transferfrom") ? postImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value : 0;
                if ((preData_TransferFrom == Convert.ToInt32(Constants.TransferFrom.Coroner) || preData_TransferFrom == Convert.ToInt32(Constants.TransferFrom.PlaceOfDeath) || preData_TransferFrom == Convert.ToInt32(Constants.TransferFrom.PlaceOfResidence)) && (preData_TransferFrom != postData_TransferFrom))
                {
                    if (preData_FuneralType == Convert.ToInt32(Constants.FuneralType.Cremation))
                    {
                        DeleteOppProduct("CH005", preImage.Id);    //Statutory Documentation for Cremation
                    }
                }


                if (preData_FuneralType != postData_FuneralType)
                {
                    DeleteOppProduct("CH009", preImage.Id); // Death Certificate
                    if (preData_FuneralType == Convert.ToInt32(Constants.FuneralType.Burial))
                        DeleteOppProduct("CH021", preImage.Id); // Burial Fee
                    else if (preData_FuneralType == Convert.ToInt32(Constants.FuneralType.Cremation))
                    {
                        DeleteOppProduct("CH004", preImage.Id); // Cremation
                        DeleteOppProduct("CH005", preImage.Id); // Doctors Fee
                    }
                    if (postData_FuneralType == Convert.ToInt32(Constants.FuneralType.Unknown) || postData_FuneralType == 0)
                    {
                        DeleteOppProduct("CH021", preImage.Id); // Burial Fee
                        DeleteOppProduct("CH004", preImage.Id); // Cremation
                        DeleteOppProduct("CH005", preImage.Id); // Doctors Fee
                    }
                }

                Guid preData_ServiceTypeId = preImage.Contains("ols_servicetypeid") ? preImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Id : Guid.Empty;
                Guid postData_ServiceTypeId = postImage.Contains("ols_servicetypeid") ? postImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Id : Guid.Empty;
                if ((preData_ServiceTypeId != Guid.Empty) && (!preData_ServiceTypeId.Equals(postData_ServiceTypeId)))
                {

                    if (postImage.Contains("pricelevelid") && postImage.GetAttributeValue<EntityReference>("pricelevelid").Name == "Walter Carter Funerals Pty Ltd")
                    {
                        switch (preImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Name)
                        {
                            case "Bespoke 1 location - Planning fee":
                                DeleteOppProduct("CH-103", postImage.Id); //Hearse transportation – 1 location
                                break;
                            case "Bespoke 2 locations - Planning fee":
                                DeleteOppProduct("CH-104", postImage.Id); //Hearse transportation – 2 location
                                break;
                            case "Immediate - Service fee":
                                DeleteOppProduct("CH-105", postImage.Id); //Hearse transport
                                DeleteOppProduct("CH0293", postImage.Id); //Ashes collect
                                break;
                            default:
                                break;
                        }
                    }

                    DeleteOppProduct(GetProductNoByID(preImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Id), postImage.Id);
                }

                if (preData_TransferFrom != postData_TransferFrom)
                {
                    DeleteOppProduct("CH002", postImage.Id); // Transfer
                }

                bool preData_AshesDeliveryInUrn = preImage.Contains("ols_ashesdeliveryinurn") ? preImage.GetAttributeValue<bool>("ols_ashesdeliveryinurn") : false;
                bool postData_AshesDeliveryInUrn = postImage.Contains("ols_ashesdeliveryinurn") ? postImage.GetAttributeValue<bool>("ols_ashesdeliveryinurn") : false;
                if ((!postData_AshesDeliveryInUrn) && (preData_AshesDeliveryInUrn != postData_AshesDeliveryInUrn))
                {
                    DeleteOppProduct("CH025", postImage.Id); //  Urn
                }

                bool preData_GraveMarker = preImage.Contains("ols_gravemarker") ? preImage.GetAttributeValue<bool>("ols_gravemarker") : false;
                bool postData_GraveMarker = postImage.Contains("ols_gravemarker") ? postImage.GetAttributeValue<bool>("ols_gravemarker") : false;
                if ((!postData_GraveMarker) && (preData_GraveMarker != postData_GraveMarker))
                {
                    DeleteOppProduct("CH024", postImage.Id); // Grave Marker
                }

                bool preData_ChapelHire = preImage.Contains("ols_chapelhire") ? preImage.GetAttributeValue<bool>("ols_chapelhire") : false;
                bool postData_ChapelHire = postImage.Contains("ols_chapelhire") ? postImage.GetAttributeValue<bool>("ols_chapelhire") : false;
                if (!postData_ChapelHire && preData_ChapelHire != postData_ChapelHire)
                {
                    DeleteOppProduct("CH014", postImage.Id);
                }

                bool preData_DoubleService = preImage.Contains("ols_doubleservice") ? preImage.GetAttributeValue<bool>("ols_doubleservice") : false;
                bool postData_DoubleService = postImage.Contains("ols_doubleservice") ? postImage.GetAttributeValue<bool>("ols_doubleservice") : false;
                if ((!preData_DoubleService) && (postData_DoubleService))
                {
                    DeleteOppProduct("CH006", postImage.Id); //Celebrant 
                }

                //no deletes cause once it is created user cannot touch Press Notice field
                //if (preData.new_pressnoticeno != postData.new_pressnoticeno)
                //{
                //    DeletePNoticeTask(svc, postData.name + " - Press Notice.", postData.opportunityid);
                //    DeleteNoticeOppProduct(svc, "CH007", postData.opportunityid); //flowers
                //}

                //Not needed Logic
                //if ((preData.transferfrom == TransferFrom.Coroner) && (preData.funeraltype == FuneralType.Cremation) && (preData.new_transferfrom != postData.new_transferfrom) {
                //    DeleteOppProduct(svc, "CH005", preData.opportunityid);    //Statutory Documentation for Cremation
                //}

                //because you cannot change funeral type
                //if (preData.new_funeraltype != postData.new_funeraltype)
                //{
                //    DeleteByFuneralType(svc, preData, postData);
                //}

            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public string GetProductNoByID(Guid productId)
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
        public void CreateByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                bool preData_DoubleService = preImage.Contains("ols_doubleservice") ? preImage.GetAttributeValue<bool>("ols_doubleservice") : false;
                bool postData_DoubleService = postImage.Contains("ols_doubleservice") ? postImage.GetAttributeValue<bool>("ols_doubleservice") : false;
                if (postData_DoubleService && !preData_DoubleService)
                {
                    if (!CheckOppProdExist(postImage.Id, "CH006"))
                        CreateOppProduct("CH028", postImage.Id);
                }

                Guid preData_ClergyCelebrantId = preImage.Contains("ols_clergycelebrantid") ? preImage.GetAttributeValue<EntityReference>("ols_clergycelebrantid").Id : Guid.Empty;
                Guid postData_ClergyCelebrantId = postImage.Contains("ols_clergycelebrantid") ? postImage.GetAttributeValue<EntityReference>("ols_clergycelebrantid").Id : Guid.Empty;
                if (postData_ClergyCelebrantId != Guid.Empty && !postData_ClergyCelebrantId.Equals(preData_ClergyCelebrantId))
                {
                    if (!CheckOppProdExist(postImage.Id, "CH028"))
                        CreateOppProduct("CH006", postImage.Id);
                }

                string preData_ClergyCelebrant = preImage.Contains("ols_clergycelebrant") ? preImage.GetAttributeValue<string>("ols_clergycelebrant") : string.Empty;
                string postData_ClergyCelebrant = postImage.Contains("ols_clergycelebrant") ? postImage.GetAttributeValue<string>("ols_clergycelebrant") : string.Empty;
                if (postData_ClergyCelebrant != string.Empty && postData_ClergyCelebrant != preData_ClergyCelebrant)
                {
                    if (!CheckOppProdExist(postImage.Id, "CH028"))
                        CreateOppProduct("CH006", postImage.Id);
                }

                int preData_flowerStatus = preImage.Contains("ols_flowers") ? preImage.GetAttributeValue<OptionSetValue>("ols_flowers").Value : 0;
                int postData_flowerStatus = postImage.Contains("ols_flowers") ? postImage.GetAttributeValue<OptionSetValue>("ols_flowers").Value : 0;
                if (postData_flowerStatus == Convert.ToInt32(Constants.FlowerStatus.FromHouse) && postData_flowerStatus != preData_flowerStatus)
                {
                    CreateOppProduct("CH008", postImage.Id); //flowers
                }

                int preData_CarsRequired = preImage.Contains("ols_carsrequired") ? preImage.GetAttributeValue<OptionSetValue>("ols_carsrequired").Value : 0;
                int postData_CarsRequired = postImage.Contains("ols_carsrequired") ? postImage.GetAttributeValue<OptionSetValue>("ols_carsrequired").Value : 0;
                if (postData_CarsRequired != 0 && preData_CarsRequired != postData_CarsRequired)
                {
                    CreateOppProduct("CH010", postImage.Id); //Cars Required 
                }

                int preData_PressNoticeNo = preImage.Contains("ols_pressnoticeno") ? preImage.GetAttributeValue<int>("ols_pressnoticeno") : 0;
                int postData_PressNoticeNo = postImage.Contains("ols_pressnoticeno") ? postImage.GetAttributeValue<int>("ols_pressnoticeno") : 0;
                if (postData_PressNoticeNo > 0 && preData_PressNoticeNo != postData_PressNoticeNo)
                {
                    //if (postData.new_pressnoticeno > 15) { postData.new_pressnoticeno = 15; }
                    for (int i = 0; i < postData_PressNoticeNo; i++)
                    {
                        CreateOppProduct("CH007", postImage.Id); //press Notice
                    }
                }

                int preData_FuneralType = preImage.Contains("ols_funeraltype") ? preImage.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value : 0;
                int postData_FuneralType = postImage.Contains("ols_funeraltype") ? postImage.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value : 0;
                if ((postData_FuneralType == Convert.ToInt32(Constants.FuneralType.Burial) || postData_FuneralType == Convert.ToInt32(Constants.FuneralType.Cremation)) && preData_FuneralType != postData_FuneralType)
                {
                    CreateByFuneralType(postImage);
                }

                int preData_TransferFrom = preImage.Contains("ols_transferfrom") ? preImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value : 0;
                int postData_TransferFrom = postImage.Contains("ols_transferfrom") ? postImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value : 0;
                if ((postData_TransferFrom == Convert.ToInt32(Constants.TransferFrom.Coroner) || postData_TransferFrom == Convert.ToInt32(Constants.TransferFrom.PlaceOfDeath) || postData_TransferFrom == Convert.ToInt32(Constants.TransferFrom.PlaceOfResidence)) && preData_TransferFrom != postData_TransferFrom)
                {
                    if (postData_FuneralType == Convert.ToInt32(Constants.FuneralType.Cremation))
                    {
                        if (!CheckOppProdExist(postImage.Id, "CH005"))
                            CreateOppProduct("CH005", postImage.Id);    //Statutory Documentation for Cremation
                    }
                }

                Guid preData_ServiceTypeId = preImage.Contains("ols_servicetypeid") ? preImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Id : Guid.Empty;
                Guid postData_ServiceTypeId = postImage.Contains("ols_servicetypeid") ? postImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Id : Guid.Empty;
                if (postData_ServiceTypeId != Guid.Empty && (!postData_ServiceTypeId.Equals(preData_ServiceTypeId)))
                {
                    if (postImage.Contains("pricelevelid") && postImage.GetAttributeValue<EntityReference>("pricelevelid").Name == "Walter Carter Funerals Pty Ltd")
                    {
                        switch (postImage.GetAttributeValue<EntityReference>("ols_servicetypeid").Name)
                        {
                            case "Bespoke 1 location - Planning fee":
                                CreateOppProduct("CH-103", postImage.Id); //Hearse transportation – 1 location
                                break;
                            case "Bespoke 2 locations - Planning fee":
                                CreateOppProduct("CH-104", postImage.Id); //Hearse transportation – 2 location
                                break;
                            case "Immediate - Service fee":
                                CreateOppProduct("CH-105", postImage.Id); //Hearse transport
                                CreateOppProduct("CH0293", postImage.Id); //Ashes collect
                                break;
                            default:
                                break;
                        }
                    }
                    CreateOppProduct(GetProductNoByID(postData_ServiceTypeId), postImage.Id);
                }

                if (postData_TransferFrom > 0 && preData_TransferFrom != postData_TransferFrom)
                {
                    CreateOppProduct("CH002", postImage.Id); // Transfer
                }

                bool preData_AshesDeliveryInUrn = preImage.Contains("ols_ashesdeliveryinurn") ? preImage.GetAttributeValue<bool>("ols_ashesdeliveryinurn") : false;
                bool postData_AshesDeliveryInUrn = postImage.Contains("ols_ashesdeliveryinurn") ? postImage.GetAttributeValue<bool>("ols_ashesdeliveryinurn") : false;
                if (postData_AshesDeliveryInUrn && preData_AshesDeliveryInUrn != postData_AshesDeliveryInUrn)
                {
                    CreateOppProduct("CH025", postImage.Id); //  Urn
                }

                bool preData_GraveMarker = preImage.Contains("ols_gravemarker") ? preImage.GetAttributeValue<bool>("ols_gravemarker") : false;
                bool postData_GraveMarker = postImage.Contains("ols_gravemarker") ? postImage.GetAttributeValue<bool>("ols_gravemarker") : false;
                if (postData_GraveMarker && preData_GraveMarker != postData_GraveMarker)
                {
                    CreateOppProduct("CH024", postImage.Id); // Grave Marker 
                }

                bool preData_ChapelHire = preImage.Contains("ols_chapelhire") ? preImage.GetAttributeValue<bool>("ols_chapelhire") : false;
                bool postData_ChapelHire = postImage.Contains("ols_chapelhire") ? postImage.GetAttributeValue<bool>("ols_chapelhire") : false;
                if (postData_ChapelHire && preData_ChapelHire != postData_ChapelHire)
                {
                    if (ServiceFees != null && ServiceFees.Value != 0)
                        CreateOppProduct("CH014", postImage.Id, ServiceFees); //Chapel Hire
                    else CreateOppProduct("CH014", postImage.Id);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public bool CheckOppProdExist(Guid opportunityid, string productNo)
        {
            try
            {
                QueryExpression qe = new QueryExpression("opportunityproduct");
                qe.ColumnSet = new ColumnSet("opportunityproductid");
                qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityid);
                qe.Criteria.AddCondition("productid", ConditionOperator.Equal, Util.GetProduct(productNo, GetService(UserType.User)).Id);
                RetrieveMultiple(UserType.User, qe);
                Entity oppProd = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                if (oppProd != null)
                    return true;
                else
                    return false;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        public Guid CreateOppProduct(string productNumber, Guid OpportunityID)
        {

            Entity Ent = new Entity("opportunityproduct");
            Entity Prod = Util.GetProduct(productNumber, GetService(UserType.User));

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
        public void CreateByFuneralType(Entity postImage)
        {
            CreateOppProduct("CH009", postImage.Id);    //Doctors Fees-Death Cerficate CH009
            if (postImage.Contains("ols_funeraltype") && postImage.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == Convert.ToInt32(Constants.FuneralType.Cremation))
            {
                if (CrematoriumServiceFees != null && CrematoriumServiceFees.Value != 0)
                    CreateOppProduct("CH004", postImage.Id, CrematoriumServiceFees); //Cremation  Booking 
                else CreateOppProduct("CH004", postImage.Id);
                if (postImage.Contains("ols_transferfrom") && (postImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value == Convert.ToInt32(Constants.TransferFrom.Coroner) || postImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value == Convert.ToInt32(Constants.TransferFrom.PlaceOfDeath) || postImage.GetAttributeValue<OptionSetValue>("ols_transferfrom").Value == Convert.ToInt32(Constants.TransferFrom.PlaceOfResidence)))
                    CreateOppProduct("CH005", postImage.Id);     // Doctors Fee
            }
            if (postImage.Contains("ols_funeraltype") && postImage.GetAttributeValue<OptionSetValue>("ols_funeraltype").Value == Convert.ToInt32(Constants.FuneralType.Burial))
            {
                if (CemetaryServiceFees != null && CemetaryServiceFees.Value != 0)
                    CreateOppProduct("CH021", postImage.Id, CemetaryServiceFees); //burial  Booking  
                else CreateOppProduct("CH021", postImage.Id);
            }
        }
        public Guid CreateOppProduct(string productNumber, Guid OpportunityID, Money price)
        {
            Entity Ent = new Entity("opportunityproduct");
            Entity Prod = Util.GetProduct(productNumber, GetService(UserType.User));
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
        public void UpdateByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                int preData_ClergyCelebrantfee = preImage.Contains("ols_clergycelebrantfee") ? preImage.GetAttributeValue<OptionSetValue>("ols_clergycelebrantfee").Value : 0;
                int postData_ClergyCelebrantfee = postImage.Contains("ols_clergycelebrantfee") ? postImage.GetAttributeValue<OptionSetValue>("ols_clergycelebrantfee").Value : 0;
                if (postData_ClergyCelebrantfee != preData_ClergyCelebrantfee)
                {
                    if (postData_ClergyCelebrantfee == (int)Constants.CelebrantFees.FamilyToPay || postData_ClergyCelebrantfee == (int)Constants.CelebrantFees.NoPayment)
                    {
                        UpdateOppProduct("CH006", postImage.Id, new Money(new decimal(0.0)));
                        UpdateOppProduct("CH028", postImage.Id, new Money(new decimal(0.0)));
                    }
                    else
                    {
                        UpdateOppProduct("CH006", postImage.Id);
                        UpdateOppProduct("CH028", postImage.Id);
                    }
                }

                bool postData_ChapelHire = postImage.Contains("ols_chapelhire") ? postImage.GetAttributeValue<bool>("ols_chapelhire") : false;
                Guid preData_ServicePlaceId = preImage.Contains("ols_serviceplaceid") ? preImage.GetAttributeValue<EntityReference>("ols_serviceplaceid").Id : Guid.Empty;
                Guid postData_ServicePlaceId = postImage.Contains("ols_serviceplaceid") ? postImage.GetAttributeValue<EntityReference>("ols_serviceplaceid").Id : Guid.Empty;
                if (postData_ChapelHire && preData_ServicePlaceId != postData_ServicePlaceId)
                {
                    if (ServiceFees.Value != 0)
                        UpdateOppProduct("CH014", postImage.Id, ServiceFees); //Chapel Hire
                    else
                        UpdateOppProduct("CH014", postImage.Id);
                }

                Guid preData_CrematoriumId = preImage.Contains("ols_crematoriumid") ? preImage.GetAttributeValue<EntityReference>("ols_crematoriumid").Id : Guid.Empty;
                Guid postData_CrematoriumId = postImage.Contains("ols_crematoriumid") ? postImage.GetAttributeValue<EntityReference>("ols_crematoriumid").Id : Guid.Empty;
                if (preData_CrematoriumId != postData_CrematoriumId)
                {
                    if (CrematoriumServiceFees.Value != 0)
                        UpdateOppProduct("CH004", postImage.Id, CrematoriumServiceFees);
                    else
                        UpdateOppProduct("CH004", postImage.Id);
                }

                Guid preData_CemetaryId = preImage.Contains("ols_cemeteryid") ? preImage.GetAttributeValue<EntityReference>("ols_cemeteryid").Id : Guid.Empty;
                Guid postData_CemetaryId = postImage.Contains("ols_cemeteryid") ? postImage.GetAttributeValue<EntityReference>("ols_cemeteryid").Id : Guid.Empty;
                if (preData_CemetaryId != postData_CemetaryId)
                {
                    if (CemetaryServiceFees.Value != 0)
                        UpdateOppProduct("CH021", postImage.Id, CemetaryServiceFees);
                    else
                        UpdateOppProduct("CH021", postImage.Id);
                }

                bool preData_DoubleService = preImage.Contains("ols_doubleservice") ? preImage.GetAttributeValue<bool>("ols_doubleservice") : false;
                bool postData_DoubleService = postImage.Contains("ols_doubleservice") ? postImage.GetAttributeValue<bool>("ols_doubleservice") : false;
                if (preData_DoubleService && !postData_DoubleService)
                {
                    DeleteOppProduct("CH028", postImage.Id); //Celebrant 

                    if (postImage.Contains("ols_clergycelebrantid") || postImage.Contains("ols_clergycelebrant"))
                    {
                        CreateOppProduct("CH006", postImage.Id);
                    }
                }

                int preData_QtyCars = preImage.Contains("ols_qtycars") ? preImage.GetAttributeValue<int>("ols_qtycars") : 0;
                int postData_QtyCars = postImage.Contains("ols_qtycars") ? postImage.GetAttributeValue<int>("ols_qtycars") : 0;
                if (preData_QtyCars != postData_QtyCars)
                {
                    UpdateOppProductQuant("CH010", postImage.Id, postData_QtyCars);
                }

            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void UpdateOppProductQuant(string productNum, Guid opportunityID, int quant)
        {
            try
            {
                Entity Prod = Util.GetProduct(productNum, GetService(UserType.User));
                if (Prod != null)
                {
                    QueryExpression qe = new QueryExpression("opportunityproduct");
                    qe.ColumnSet = new ColumnSet("opportunityproductid", "quantity");
                    qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityID);
                    qe.Criteria.AddCondition("productid", ConditionOperator.Equal, Prod.Id);

                    Entity ent = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                    if (ent != null)
                    {
                        Entity updateOppProduct = new Entity(ent.LogicalName, ent.Id);
                        ent["quantity"] = new decimal(quant);
                        Update(UserType.User, ent);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateOppProduct(string productNum, Guid opportunityID, Money price)
        {
            try
            {
                Entity Prod = Util.GetProduct(productNum, GetService(UserType.User));
                if (Prod != null)
                {
                    QueryExpression qe = new QueryExpression("opportunityproduct");
                    qe.ColumnSet = new ColumnSet(true);
                    qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityID);
                    qe.Criteria.AddCondition("productid", ConditionOperator.Equal, Prod.Id);

                    EntityCollection ec = RetrieveMultiple(UserType.User, qe);
                    if (ec != null && ec.Entities.Count > 0)
                        foreach (Entity en in ec.Entities)
                        {
                            UpdatePrice(en.Id, price);
                        }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateOppProduct(string productNum, Guid opportunityID)
        {
            try
            {
                Entity Prod = Util.GetProduct(productNum, GetService(UserType.User));
                if (Prod != null)
                {
                    QueryExpression qe = new QueryExpression("opportunityproduct");
                    qe.ColumnSet = new ColumnSet(true);
                    qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityID);
                    qe.Criteria.AddCondition("productid", ConditionOperator.Equal, Prod.Id);

                    EntityCollection ec = RetrieveMultiple(UserType.User, qe);
                    if (ec != null && ec.Entities.Count > 0)
                        foreach (Entity en in ec.Entities)
                        {
                            UpdateOriginalPrice(en.Id, Prod.Contains("price") ? Prod.GetAttributeValue<Money>("price") : null);
                        }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdatePrice(Guid productid, Money serviceFees)
        {
            try
            {
                Entity updateOppProduct = new Entity("opportunityproduct", productid);
                updateOppProduct["priceperunit"] = serviceFees;
                updateOppProduct["ispriceoverridden"] = true;
                Update(UserType.User, updateOppProduct);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void UpdateOriginalPrice(Guid productid, Money serviceFees)
        {
            try
            {
                Entity updateOppProduct = new Entity("opportunityproduct", productid);
                updateOppProduct["priceperunit"] = serviceFees;
                updateOppProduct["ispriceoverridden"] = false;
                Update(UserType.User, updateOppProduct);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        public void SetDueDate(Entity preImage, Entity postImage)
        {
            try
            {
                DateTime? predata_ServicePlaceSession = preImage.Contains("ols_serviceplacesessionfrom") ? preImage.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom") : (DateTime?)null;
                DateTime? postdata_ServicePlaceSession = postImage.Contains("ols_serviceplacesessionfrom") ? postImage.GetAttributeValue<DateTime>("ols_serviceplacesessionfrom") : (DateTime?)null;

                if (predata_ServicePlaceSession != postdata_ServicePlaceSession && postdata_ServicePlaceSession != null)
                {
                    DateTime dt = postdata_ServicePlaceSession.Value.AddDays(-1);
                    EntityCollection tasksColl = GetAllTask(postImage.Id);
                    if (tasksColl != null)
                    {
                        foreach (Entity ent in tasksColl.Entities)
                        {
                            ent["scheduledend"] = dt;
                            Update(UserType.User, ent);
                        }
                    }

                    EntityCollection apptColl = GetAllAppt(postImage.Id);
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
        public void SetPrePaidLock(Entity preImage, Entity postImage)
        {
            try
            {
                int preData_prePaidstatus = preImage.Contains("ols_prepaidstatus") ? preImage.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value : 0;
                int postData_prePaidstatus = postImage.Contains("ols_prepaidstatus") ? postImage.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value : 0;

                if (postData_prePaidstatus == (int)Constants.PrePaidStatus.Active && preData_prePaidstatus != postData_prePaidstatus)
                {
                    EntityCollection oppProdColl = GetOppProduct(postImage.Id);
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
            finally
            {

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
        #endregion
    }
    public class Age
    {
        private int _value;
        private string _unit;
        private int _unitvalue;

        public int Value
        {
            get { return _value; }
            set { _value = value; }
        }
        public string Unit
        {
            get { return _unit; }
            set { _unit = value; }
        }
        public int UnitValue
        {
            get { return _unitvalue; }
            set { _unitvalue = value; }
        }
        public Age(int number, string unit)
        {
            _value = number;
            _unit = unit;
            switch (unit)
            {
                case "years":
                    _unitvalue = 6;
                    break;
                case "months":
                    _unitvalue = 4;
                    break;
                case "weeks":
                    _unitvalue = 5;
                    break;
                case "days":
                    _unitvalue = 1;
                    break;
                case "hours":
                    _unitvalue = 2;
                    break;
            }
        }

    }
    public class DateDiff
    {
        public int inYears
        {
            get; set;
        }
        public int inDays
        {
            get; set;
        }
        public int inWeeks
        {
            get; set;
        }
        public int inMonths
        {
            get; set;
        }

        public DateDiff(DateTime d1, DateTime d2)
        {
            TimeSpan span = d2 - d1;
            inDays = span.Days;//Convert.ToInt32((d2 - d1)) / (1000 * 60 * 60 * 24);
            inWeeks = inDays / 7;// Convert.ToInt32(Math.Floor(Convert.ToDecimal(Convert.ToInt32(d2 - d1) / 1000 / 60 / 60 / 24 / 7)));

            var d1Y = d1.Year;
            var d2Y = d2.Year;
            var d1M = d1.Month;
            var d2M = d2.Month;
            var d2d = d2.Date;
            var d1d = d1.Date;
            if ((d2M + 12 * d2Y) - (d1M + 12 * d1Y) > 1 || (d2d > d1d))
                inMonths = (d2M + 12 * d2Y) - (d1M + 12 * d1Y);
            else if (((d2M + 12 * d2Y) - (d1M + 12 * d1Y) == 1) && (d2d == d1d)) inMonths = 1;
            else inMonths = 0;

            var m1 = d1.Month;
            var m2 = d2.Month;
            var y1 = d1.Year;
            var y2 = d2.Year;
            var dd1 = d1.Date;
            var dd2 = d2.Date;
            if (y2 > y1)
            {
                if (m2 > m1) inYears = y2 - y1;
                else
                    if (m2 < m1) inYears = y2 - y1 - 1;
                else
                        if (m2 == m1)
                {
                    if (dd2 >= dd1) inYears = y2 - y1;
                    else inYears = y2 - y1 - 1;
                }
            }
            else if (y1 > y2) inYears = -1;
            else inYears = 0;
        }
    }

    public class PostUpdate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string Securestring = string.Empty;
        public PostUpdate(string unsecConfig, string securestring)
        {
            UnsecConfig = unsecConfig;
            Securestring = securestring;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostUpdateWrapper(UnsecConfig, Securestring);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
