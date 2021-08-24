using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Association
{
    public class PostAssociateWrapper : PluginHelper
    {
        /// <summary>
        /// Triggers on PostAssociate of Funeral with Bearers, Conductors & Hearses. 
        /// Retrieves all associated Bearers and stores Bearer names ";" separated in string field- Bearer
        /// Retrieves all associated Conductors and stores Conductor names ";" separated in string field- Conductor
        /// Retrieves all associated Hearses and stores Hearse names ";" separated in string field- Hearse
        /// </summary>
        /// <param name="unsecConfig"></param>
        /// <param name="securestring"></param>

        public PostAssociateWrapper(string unsecConfig, string secureString) : base(unsecConfig, secureString) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "associate" || !Context.InputParameters.Contains("Target")) return;

                AppendLog("PostAssociate - Plugin Excecution is Started.");

                Relationship relationship = (Relationship)Context.InputParameters["Relationship"];
                EntityReference target = (EntityReference)Context.InputParameters["Target"];
                EntityReferenceCollection relatedEnts = (EntityReferenceCollection)Context.InputParameters["RelatedEntities"];
                if (relatedEnts == null || relatedEnts.Count == 0) return;
                if ((target == null || target.LogicalName != "opportunity" || target.Id == Guid.Empty))
                {
                    AppendLog("Target is null");
                    return;
                }

                if (relationship.SchemaName == "ols_opportunity_ols_bearer")
                    PostAssociateBearer(target.Id);

                else if (relationship.SchemaName == "ols_opportunity_ols_conductor")
                    PostAssociateConductor(target.Id);

                else if (relationship.SchemaName == "ols_opportunity_ols_hearse")
                    PostAssociateHearse(target.Id);

                AppendLog("PostAssociate - Plugin Excecution is Completed.");
            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        #region [Public Methods]

        #region Bearers Methods
        public void PostAssociateBearer(Guid Id) //RelatedEntities = ols_bearer, Target = opportunity
        {
            try
            {
                AppendLog("PostAssociateBearer method strated.");
                var bearersList = new StringBuilder();
                var bearers = GetAssociatedBearers(Id);
                if (bearers != null && bearers.Entities != null && bearers.Entities.Count > 0)
                {
                    AppendLog("Total Beares count: " + bearers.Entities.Count);
                    foreach (Entity bearer in bearers.Entities)
                    {
                        if (bearer.Contains("ols_name"))
                            bearersList.Append(bearer.GetAttributeValue<string>("ols_name")).Append("; ");
                    }
                }
                AppendLog("bearersList: " + bearersList);
                UpdateOpportunityBearer(bearersList.ToString(), Id);
                AppendLog("PostAssociateBearer method completed.");
            }
            catch (Exception ex) { throw ex; }
        }
        public EntityCollection GetAssociatedBearers(Guid id)
        {
            try
            {
                string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                          <entity name='ols_bearer'>
                            <attribute name='ols_bearerid' />
                            <attribute name='ols_name' />
                            <attribute name='createdon' />
                            <attribute name='ols_fullname' />
                            <order attribute='createdon' descending='true' />
                            <link-entity name='ols_opportunity_ols_bearer' from='ols_bearerid' to='ols_bearerid' visible='false' intersect='true'>
                              <link-entity name='opportunity' from='opportunityid' to='opportunityid' alias='ac'>
                                <filter type='and'>
                                  <condition attribute='opportunityid' operator='eq'  uitype='opportunity' value='{0}' />
                                </filter>
                              </link-entity>
                            </link-entity>
                          </entity>
                        </fetch>";
                return RetrieveMultiple(UserType.User, new FetchExpression(string.Format(fetch, id)));
            }
            catch (Exception ex) { throw ex; }
        }
        public void UpdateOpportunityBearer(string names, Guid opportunityId)
        {
            try
            {
                AppendLog("Updating opportunity for Bearers started");
                Entity updateOpportunity = new Entity("opportunity", opportunityId);
                updateOpportunity["ols_bearers"] = names;
                Update(UserType.User, updateOpportunity);
                AppendLog("Updating opportunity for Bearers completed");
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion

        #region Conductors Methods
        public void PostAssociateConductor(Guid Id) //RelatedEntities = ols_conductor, Target = opportunity
        {
            try
            {
                AppendLog("PostAssociateConductor method strated.");
                var conductorsList = new StringBuilder();
                var conductors = GetAssociatedConductors(Id);
                if (conductors != null && conductors.Entities.Count > 0)
                {
                    foreach (Entity conductor in conductors.Entities)
                    {
                        if (conductor.Contains("ols_name"))
                            conductorsList.Append(conductor.GetAttributeValue<string>("ols_name")).Append("; ");
                    }
                }
                UpdateOpportunityConductor(conductorsList.ToString(), Id);
                AppendLog("PostAssociateConductor method completed.");
            }
            catch (Exception ex) { throw ex; }
        }
        public EntityCollection GetAssociatedConductors(Guid id)
        {
            try
            {
                string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                      <entity name='ols_conductor'>
                        <attribute name='ols_conductorid' />
                        <attribute name='ols_name' />
                        <attribute name='createdon' />
                        <order attribute='createdon' descending='true' />
                        <link-entity name='ols_opportunity_ols_conductor' from='ols_conductorid' to='ols_conductorid' visible='false' intersect='true'>
                          <link-entity name='opportunity' from='opportunityid' to='opportunityid' alias='ab'>
                            <filter type='and'>
                              <condition attribute='opportunityid' operator='eq'  uitype='opportunity' value='{0}' />
                            </filter>
                          </link-entity>
                        </link-entity>
                      </entity>
                    </fetch>";
                return RetrieveMultiple(UserType.User, new FetchExpression(string.Format(fetch, id)));
            }
            catch (Exception ex) { throw ex; }
        }
        public void UpdateOpportunityConductor(string names, Guid opportunityId)
        {
            try
            {
                AppendLog("Updating opportunity for Conductors started");
                Entity updateOpportunity = new Entity("opportunity", opportunityId);
                updateOpportunity["ols_conductors"] = names;
                Update(UserType.User, updateOpportunity);
                AppendLog("Updating opportunity for Conductors completed");
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion

        #region Hearses Methods
        public void PostAssociateHearse(Guid Id) //RelatedEntities = ols_hearse, Target = opportunity
        {
            try
            {
                AppendLog("PostAssociateHearse method strated.");
                var hearsesList = new StringBuilder();
                var hearsesColl = GetAssociatedHearses(Id);
                if (hearsesColl != null && hearsesColl.Entities.Count > 0)
                {
                    foreach (Entity hearse in hearsesColl.Entities)
                    {
                        if (hearse.Contains("ols_name"))
                            hearsesList.Append(hearse.GetAttributeValue<string>("ols_name")).Append("; ");
                    }
                }
                UpdateOpportunityHearse(hearsesList.ToString(), Id);
                AppendLog("PostAssociateHearse method completed.");
            }
            catch (Exception ex) { throw ex; }
        }
        public EntityCollection GetAssociatedHearses(Guid id)
        {
            try
            {
                string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='ols_hearse'>
                                <attribute name='ols_name' />
                                <attribute name='createdon' />
                                <attribute name='ols_hearseid' />
                                <order attribute='createdon' descending='true' />
                                <link-entity name='ols_opportunity_ols_hearse' from='ols_hearseid' to='ols_hearseid' visible='false' intersect='true'>
                                  <link-entity name='opportunity' from='opportunityid' to='opportunityid' alias='ad'>
                                    <filter type='and'>
                                      <condition attribute='opportunityid' operator='eq'  uitype='opportunity' value='{0}' />
                                    </filter>
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>";
                return RetrieveMultiple(UserType.User, new FetchExpression(string.Format(fetch, id)));
            }
            catch (Exception ex) { throw ex; }
        }
        public void UpdateOpportunityHearse(string names, Guid opportunityId)
        {
            try
            {
                AppendLog("Updating opportunity for Hearses started");
                Entity updateOpportunity = new Entity("opportunity", opportunityId);
                updateOpportunity["ols_hearses"] = names;
                Update(UserType.User, updateOpportunity);
                AppendLog("Updating opportunity for Hearses completed");
            }
            catch (Exception ex) { throw ex; }
        }
        #endregion

        #endregion

    }
    public class PostAssociate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string SecureString = string.Empty;
        public PostAssociate(string unsecConfig, string secureString)
        {
            UnsecConfig = unsecConfig;
            SecureString = secureString;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostAssociateWrapper(UnsecConfig, SecureString);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
