using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Operations
{
    public class PostUpdateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostUpdate of Funeral Operation. 
        /// Creates, deletes Activity Tasks based on different options/fields
        /// Throws Error if Funeral is Pre Paid active
        /// Creates, updates, deletes Funeral products based on different options/fields
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>

        public PostUpdateWrapper(string unsecConfig, string securestring) : base(unsecConfig, securestring) { }

        #region [Variables]
        public string deceasedName = string.Empty;
        public Guid opportunityId = Guid.Empty;
        public int prePaidStatus = 0;
        public int funeralStatus = -1;
        #endregion

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("FuneralOperations PostUpdate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];
                Entity preImage = Context.PreEntityImages.Contains("PreImage") ? (Entity)Context.PreEntityImages["PreImage"] : null;
                Entity postImage = Context.PostEntityImages.Contains("PostImage") ? (Entity)Context.PostEntityImages["PostImage"] : null;

                if ((target == null || target.LogicalName != "ols_funeraloperations" || target.Id == Guid.Empty) || (preImage == null || preImage.Id == Guid.Empty) || (postImage == null || postImage.Id == Guid.Empty))
                {
                    AppendLog("Target/PreImage/PostImage is null");
                    return;
                }

                #region Translate Fields 
                Entity opportunity = GetOppByFuneralOppId(target.Id);
                if (opportunity != null)
                {
                    opportunityId = opportunity.Id;
                    deceasedName = opportunity.Contains("name") ? opportunity.GetAttributeValue<string>("name") : string.Empty;
                    if (opportunity.Contains("ols_prepaidstatus"))
                        prePaidStatus = opportunity.GetAttributeValue<OptionSetValue>("ols_prepaidstatus").Value;
                    if (opportunity.Contains("ols_status"))
                        funeralStatus = opportunity.GetAttributeValue<OptionSetValue>("ols_status").Value;
                }
                #endregion

                ApplyTaskOperations(preImage, postImage);

                if (prePaidStatus == (int)Constants.PrePaidStatus.Active)
                {
                    AppendLog("This is a Pre Paid active record. No Changes are allowed");
                    throw new InvalidPluginExecutionException("This is a Pre Paid active record. No Changes are allowed");
                }
                if ((funeralStatus == (int)Constants.FuneralStatus.PrePaid) && (prePaidStatus != (int)Constants.PrePaidStatus.InActive))
                {
                    AppendLog("This is a Pre Paid & PrePaidStatus inactive record.");
                    return;
                }

                ApplyOppProductOperations(preImage, postImage);

                AppendLog("FuneralOperations PostUpdate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #region [Public Methods]

        #region Activity Task Operation Methods
        public void ApplyTaskOperations(Entity preImage, Entity postImage)
        {
            try
            {
                DeleteByOptions(preImage, postImage);
                CreateByOptions(preImage, postImage);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void DeleteByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                DateTime? preData_ViewingDate = preImage.Contains("ols_viewingdate") ? preImage.GetAttributeValue<DateTime>("ols_viewingdate") : (DateTime?)null;
                DateTime? postData_ViewingDate = postImage.Contains("ols_viewingdate") ? postImage.GetAttributeValue<DateTime>("ols_viewingdate") : (DateTime?)null;
                if (postData_ViewingDate == null && preData_ViewingDate != null)
                {
                    DeleteTask(opportunityId, deceasedName, "Preparation/Embalming.", 92);
                    DeleteTask(opportunityId, deceasedName, "Viewing Place & Time.", 24);
                }

                bool preData_Photo = preImage.Contains("ols_photo") ? preImage.GetAttributeValue<bool>("ols_photo") : false;
                bool postData_Photo = postImage.Contains("ols_photo") ? postImage.GetAttributeValue<bool>("ols_photo") : false;
                if (!postData_Photo && postData_Photo != preData_Photo)
                {
                    DeleteTask(opportunityId, deceasedName, "Photo Presentation.", 91);
                }

                int preData_oosStyle = preImage.Contains("ols_oosstyle") ? preImage.GetAttributeValue<OptionSetValue>("ols_oosstyle").Value : 0;
                int postData_oosStyle = postImage.Contains("ols_oosstyle") ? postImage.GetAttributeValue<OptionSetValue>("ols_oosstyle").Value : 0;
                if (postData_oosStyle != (int)Constants.OOSStyle.A && postData_oosStyle != (int)Constants.OOSStyle.B && postData_oosStyle != (int)Constants.OOSStyle.C && postData_oosStyle != preData_oosStyle)
                {
                    DeleteTask(opportunityId, deceasedName, "Order of Service.", 90);
                }

                bool preData_LocksOfHair = preImage.Contains("ols_locksofhair") ? preImage.GetAttributeValue<bool>("ols_locksofhair") : false;
                bool postData_LocksOfHair = postImage.Contains("ols_locksofhair") ? postImage.GetAttributeValue<bool>("ols_locksofhair") : false;
                if (!postData_LocksOfHair && postData_LocksOfHair != preData_LocksOfHair)
                {
                    DeleteTask(opportunityId, deceasedName, "Locks of Hair.", 54);
                }

                if (!IsClothing(postImage))
                {
                    DeleteTask(opportunityId, deceasedName, "Clothing Sheet.", 86);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void CreateByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                IOrganizationService svc = GetService(UserType.User);
                Guid opsQId = Util.GetQueueId(svc, "Operations");
                Guid AdminQId = Util.GetQueueId(svc, "Administration");
                Guid FAdminQId = Util.GetQueueId(svc, "Funeral Admin");
                string funeralNumber = postImage.Contains("ols_funeralnumber") ? postImage.GetAttributeValue<string>("ols_funeralnumber") : string.Empty;
                DateTime? preData_ViewingDate = preImage.Contains("ols_viewingdate") ? preImage.GetAttributeValue<DateTime>("ols_viewingdate") : (DateTime?)null;
                DateTime? postData_ViewingDate = postImage.Contains("ols_viewingdate") ? postImage.GetAttributeValue<DateTime>("ols_viewingdate") : (DateTime?)null;
                if (postData_ViewingDate != null && preData_ViewingDate == null)
                {
                    CreateTask(opsQId, opportunityId, deceasedName, "Preparation/Embalming.", 92, funeralNumber);
                    CreateTask(opsQId, opportunityId, deceasedName, "Viewing Place & Time.", 24, funeralNumber);
                }

                bool preData_Photo = preImage.Contains("ols_photo") ? preImage.GetAttributeValue<bool>("ols_photo") : false;
                bool postData_Photo = postImage.Contains("ols_photo") ? postImage.GetAttributeValue<bool>("ols_photo") : false;
                if (postData_Photo && postData_Photo != preData_Photo)
                {
                    CreateTask(FAdminQId, opportunityId, deceasedName, "Photo Presentation.", 91, funeralNumber);
                }

                int preData_oosStyle = preImage.Contains("ols_oosstyle") ? preImage.GetAttributeValue<OptionSetValue>("ols_oosstyle").Value : 0;
                int postData_oosStyle = postImage.Contains("ols_oosstyle") ? postImage.GetAttributeValue<OptionSetValue>("ols_oosstyle").Value : 0;
                if ((postData_oosStyle == (int)Constants.OOSStyle.A || postData_oosStyle == (int)Constants.OOSStyle.B || postData_oosStyle == (int)Constants.OOSStyle.C) && postData_oosStyle != preData_oosStyle)
                {
                    if (!ChkTaskExist(deceasedName, "Order of Service.", opportunityId, 90))
                        CreateTask(opportunityId, deceasedName, "Order of Service.", 90, funeralNumber);
                }

                bool preData_LocksOfHair = preImage.Contains("ols_locksofhair") ? preImage.GetAttributeValue<bool>("ols_locksofhair") : false;
                bool postData_LocksOfHair = postImage.Contains("ols_locksofhair") ? postImage.GetAttributeValue<bool>("ols_locksofhair") : false;
                if (postData_LocksOfHair == true && postData_LocksOfHair != preData_LocksOfHair)
                {
                    CreateTask(opsQId, opportunityId, deceasedName, "Locks of Hair.", 54, funeralNumber);
                }

                if (IsClothing(postImage))
                {
                    if (!ChkTaskExist(deceasedName, "Clothing Sheet.", opportunityId, 86))
                        CreateTask(opportunityId, deceasedName, "Clothing Sheet.", 86, funeralNumber);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public Guid CreateTask(Guid id, string name, string subject, int taskOrder, string funeralNo)
        {

            Entity Ent = new Entity("task");

            try
            {
                Ent["subject"] = GetSubject(name, subject, taskOrder);
                Ent["regardingobjectid"] = new EntityReference("opportunity", id);
                Ent["ols_funeralnumber"] = funeralNo;
                return Create(UserType.User, Ent);
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
        public bool ChkTaskExist(string name, string subject, Guid id, int taskOrder)
        {
            Entity ent;

            try
            {
                QueryExpression qe = new QueryExpression("task");
                qe.ColumnSet = new ColumnSet("activityid");
                qe.Criteria.AddCondition("subject", ConditionOperator.Equal, GetSubject(name, subject, taskOrder));
                qe.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, id);

                ent = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();

                if (ent != null)
                    return true;
                else
                    return false;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        public void CreateTask(Guid queueId, Guid oppId, string Name, string subject, int taskOrder, string funeralNo)
        {
            AddToQueueRequest req;
            try
            {
                Guid TaskId = CreateTask(oppId, Name, subject, taskOrder, funeralNo);
                req = new AddToQueueRequest();
                req.Target = new EntityReference("task", TaskId);
                req.DestinationQueueId = queueId;
                Execute(UserType.User, req);
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                req = null;
            }
        }
        public bool IsClothing(Entity postImage)
        {
            try
            {
                if ((postImage.Contains("ols_dress") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_dress"))) ||
                 (postImage.Contains("ols_jacket") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_jacket"))) ||
                 (postImage.Contains("ols_slacks") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_slacks"))) ||
                 (postImage.Contains("ols_socks") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_socks"))) ||
                 (postImage.Contains("ols_suit") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_suit"))) ||
                 (postImage.Contains("ols_underwear") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_underwear"))) ||
                 (postImage.Contains("ols_shoes") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_shoes"))) ||
                 (postImage.Contains("ols_shirt") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_shirt"))) ||
                 (postImage.Contains("ols_skirt") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_skirt"))) ||
                 (postImage.Contains("ols_valuables") && !string.IsNullOrEmpty(postImage.GetAttributeValue<string>("ols_valuables"))))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void DeleteTask(Guid ID, string name, string subject, int taskOrder)
        {
            try
            {
                Guid id = GetActivityID(GetSubject(name, subject, taskOrder), ID);
                if (id != Guid.Empty)
                    Delete(UserType.User, "task", id);

            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public Guid GetActivityID(string subject, Guid id)
        {
            Entity ent;
            try
            {
                QueryExpression qe = new QueryExpression("task");
                qe.ColumnSet = new ColumnSet("activityid");
                qe.Criteria.AddCondition("subject", ConditionOperator.Equal, subject);
                qe.Criteria.AddCondition("regardingobjectid", ConditionOperator.Equal, id);

                ent = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();

                if (ent != null)
                    return ent.Id;
                else
                    return Guid.Empty;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        #endregion

        #region Opportunity Product Operation Methods
        public void ApplyOppProductOperations(Entity preImage, Entity postImage)
        {
            try
            {
                DeleteOppProductByOptions(preImage, postImage);
                CreateOppProductByOptions(preImage, postImage);
                UpdateByOptions(preImage, postImage);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void UpdateByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                int preData_NoToBePrinted = preImage.Contains("ols_notobeprinted") ? preImage.GetAttributeValue<int>("ols_notobeprinted") : 0;
                int postData_NoToBePrinted = postImage.Contains("ols_notobeprinted") ? postImage.GetAttributeValue<int>("ols_notobeprinted") : 0;
                if (postData_NoToBePrinted != 0 && preData_NoToBePrinted != postData_NoToBePrinted)
                {
                    if (preData_NoToBePrinted != 0)
                        UpdateOppProductQuant("CH033", opportunityId, postData_NoToBePrinted);
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
        public void CreateOppProductByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                int preData_EmbalmingType = preImage.Contains("ols_embalmingtype") ? preImage.GetAttributeValue<OptionSetValue>("ols_embalmingtype").Value : 0;
                int postData_EmbalmingType = postImage.Contains("ols_embalmingtype") ? postImage.GetAttributeValue<OptionSetValue>("ols_embalmingtype").Value : 0;
                if (postData_EmbalmingType == (int)Constants.EmbalmingType.Partial && preData_EmbalmingType != postData_EmbalmingType)
                {
                    CreateOppProduct("CH027", opportunityId);
                }

                if (postData_EmbalmingType == (int)Constants.EmbalmingType.Full && preData_EmbalmingType != postData_EmbalmingType)
                {
                    CreateOppProduct("CH025", opportunityId);
                }

                string preData_Dressing = preImage.Contains("ols_dressing") ? preImage.GetAttributeValue<string>("ols_dressing") : string.Empty;
                string postData_Dressing = postImage.Contains("ols_dressing") ? postImage.GetAttributeValue<string>("ols_dressing") : string.Empty;
                if (!string.IsNullOrEmpty(postData_Dressing) && string.IsNullOrEmpty(preData_Dressing))
                {
                    if ((!ChkOppProdExist(opportunityId, "CH025")) && (!ChkOppProdExist(opportunityId, "CH027")))  //if embalming does not exist create 
                        CreateOppProduct("CH026", opportunityId); //Dressing 
                }

                int preData_NoToBePrinted = preImage.Contains("ols_notobeprinted") ? preImage.GetAttributeValue<int>("ols_notobeprinted") : 0;
                int postData_NoToBePrinted = postImage.Contains("ols_notobeprinted") ? postImage.GetAttributeValue<int>("ols_notobeprinted") : 0;
                if (postData_NoToBePrinted != 0 && preData_NoToBePrinted != postData_NoToBePrinted)
                {
                    if (preData_NoToBePrinted == 0)
                        CreateOppProduct("CH033", opportunityId, postData_NoToBePrinted); //Viewing
                }

                Guid preData_CoffinId = preImage.Contains("ols_coffinid") ? preImage.GetAttributeValue<EntityReference>("ols_coffinid").Id : Guid.Empty;
                Guid postData_CoffinId = postImage.Contains("ols_coffinid") ? postImage.GetAttributeValue<EntityReference>("ols_coffinid").Id : Guid.Empty;
                if (postData_CoffinId != Guid.Empty && (!postData_CoffinId.Equals(preData_CoffinId)))
                {
                    CreateOppProduct(GetProductNoByID(postData_CoffinId), opportunityId); //Viewing
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public Guid CreateOppProduct(string productNum, Guid opportunityID, int quantity)
        {

            Entity Ent = new Entity("opportunityproduct");
            Entity Prod = Util.GetProduct(productNum, GetService(UserType.User));

            try
            {
                if (Prod != null)
                {
                    Ent["productid"] = new EntityReference("product", Prod.Id);
                    if (Prod.Contains("defaultuomid"))
                        Ent["uomid"] = new EntityReference("uom", Prod.GetAttributeValue<EntityReference>("defaultuomid").Id);
                    if (Prod.Contains("price"))
                        Ent["priceperunit"] = Prod.GetAttributeValue<Money>("price");
                    Ent["quantity"] = new decimal(quantity);
                    Ent["opportunityid"] = new EntityReference("opportunity", opportunityID);
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
        public bool ChkOppProdExist(Guid opportunityid, string productNo)
        {
            Entity ent = null;
            try
            {
                Entity prod = Util.GetProduct(productNo, GetService(UserType.User));
                if (prod != null)
                {
                    QueryExpression qe = new QueryExpression("opportunityproduct");
                    qe.ColumnSet = new ColumnSet("opportunityproductid");
                    qe.Criteria.AddCondition("productid", ConditionOperator.Equal, prod.Id);
                    qe.Criteria.AddCondition("opportunityid", ConditionOperator.Equal, opportunityid);
                    ent = RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
                }
                if (ent != null)
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
        public void DeleteOppProductByOptions(Entity preImage, Entity postImage)
        {
            try
            {
                string preData_Dressing = preImage.Contains("ols_dressing") ? preImage.GetAttributeValue<string>("ols_dressing") : string.Empty;
                string postData_Dressing = postImage.Contains("ols_dressing") ? postImage.GetAttributeValue<string>("ols_dressing") : string.Empty;
                if (!string.IsNullOrEmpty(preData_Dressing) && string.IsNullOrEmpty(postData_Dressing))
                {
                    DeleteOppProduct("CH026", opportunityId); //Dressing
                }

                int preData_NoToBePrinted = preImage.Contains("ols_notobeprinted") ? preImage.GetAttributeValue<int>("ols_notobeprinted") : 0;
                int postData_NoToBePrinted = postImage.Contains("ols_notobeprinted") ? postImage.GetAttributeValue<int>("ols_notobeprinted") : 0;
                if (postData_NoToBePrinted == 0 && preData_NoToBePrinted != postData_NoToBePrinted)
                {
                    DeleteOppProduct("CH033", opportunityId);
                }

                Guid preData_CoffinId = preImage.Contains("ols_coffinid") ? preImage.GetAttributeValue<EntityReference>("ols_coffinid").Id : Guid.Empty;
                Guid postData_CoffinId = postImage.Contains("ols_coffinid") ? postImage.GetAttributeValue<EntityReference>("ols_coffinid").Id : Guid.Empty;
                if (preData_CoffinId != Guid.Empty && !preData_CoffinId.Equals(postData_CoffinId))
                {
                    DeleteOppProduct(GetProductNoByID(preData_CoffinId), opportunityId); //Viewing
                }

                int preData_EmbalmingType = preImage.Contains("ols_embalmingtype") ? preImage.GetAttributeValue<OptionSetValue>("ols_embalmingtype").Value : 0;
                int postData_EmbalmingType = postImage.Contains("ols_embalmingtype") ? postImage.GetAttributeValue<OptionSetValue>("ols_embalmingtype").Value : 0;
                if (postData_EmbalmingType != (int)Constants.EmbalmingType.Full && preData_EmbalmingType != postData_EmbalmingType)
                {
                    DeleteOppProduct("CH025", opportunityId);
                    DeleteOppProduct("CH026", opportunityId);
                }

                if (postData_EmbalmingType != (int)Constants.EmbalmingType.Partial && preData_EmbalmingType != postData_EmbalmingType)
                {
                    DeleteOppProduct("CH027", opportunityId);
                    DeleteOppProduct("CH026", opportunityId);
                }
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
        #endregion

        public string AddZeroPrefix(int num)
        {

            try
            {
                string numStr = num.ToString();

                if (numStr.Length == 1)
                {
                    return "0" + numStr;
                }
                else
                {
                    return numStr;
                }

                return numStr;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetSubject(string name, string subject, int taskOrder)
        {

            try
            {
                return name + " - " + AddZeroPrefix(taskOrder) + " " + subject;
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        public Entity GetOppByFuneralOppId(Guid id)
        {

            try
            {
                QueryExpression qe = new QueryExpression("opportunity");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_operationsid", ConditionOperator.Equal, id);

                return RetrieveMultiple(UserType.User, qe).Entities.FirstOrDefault();
            }
            catch (Exception e)
            {
                throw (e);
            }
        }

        #endregion
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
