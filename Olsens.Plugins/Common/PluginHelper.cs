using System;
using System.Text;
using System.Xml;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Collections.Generic;
using System.IO;
using System.ServiceModel;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace Olsens.Plugins.Common
{
    public abstract class PluginHelper : IDisposable
    {
        #region Variables

        public enum UserType
        {
            System, User, InitiatingUser
        }

        protected IPluginExecutionContext Context;
        protected IOrganizationServiceFactory ServiceFactory;

        private readonly bool _enableLogging;
        private readonly bool _throwError;
        private readonly bool _logImages;
        private readonly bool _logInputParameters;
        private bool _logged;

        private bool _isFaulted;
        private Exception _faultException;

        private StringBuilder _logBuilder;
        private readonly string _pluginName;

        protected ITracingService TracingService;

        private Guid TransactionId;
        private int _sequence;

        private Dictionary<Guid, IOrganizationService> _serviceCollection;

        #endregion Variables

        public PluginHelper(string unsecConfig, string secureString)
        {
            _pluginName = this.GetType().FullName;
            if (!String.IsNullOrEmpty(unsecConfig))
            {
                XmlDocument pluginConfig = new XmlDocument();
                pluginConfig.LoadXml(unsecConfig);
                _enableLogging = PluginConfiguration.GetConfigDataBool(pluginConfig, "EnableLogging");
                _throwError = PluginConfiguration.GetConfigDataBool(pluginConfig, "ThrowError");
                _logInputParameters = PluginConfiguration.GetConfigDataBool(pluginConfig, "LogInputParameters");
                _logImages = PluginConfiguration.GetConfigDataBool(pluginConfig, "LogImages");
            }
            else
            {
                _enableLogging = false;
                _throwError = true;
                _logImages = false;
                _logInputParameters = false;
            }

            _logBuilder = new StringBuilder();
            TransactionId = Guid.Empty;
            _sequence = 0;
            _serviceCollection = new Dictionary<Guid, IOrganizationService>();
        }

        protected abstract void Execute();

        public void Execute(IServiceProvider serviceProvider)
        {
            ClearParameters();
            Context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ServiceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));

            TracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            AppendLog("Plugin Name: " + _pluginName);
            try
            {
                TransactionId = Context.CorrelationId;

                if (!int.TryParse(GetSharedVaribleValue(Context, "$Sequence$").ToString(), out _sequence) || _sequence == 0)
                {
                    _sequence = 1;
                    Context.SharedVariables["$Sequence$"] = _sequence;
                }
                else
                {
                    _sequence += 1;
                    Context.SharedVariables["$Sequence$"] = _sequence;
                }

                AppendLog("Entering Execute");
                Execute();
                AppendLog("Completed Execution");
            }
            catch (InvalidPluginExecutionException ex)
            {
                CreateErrorLog(ex);
                throw;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                CreateErrorLog(ex);
                throw;
            }
            catch (Exception ex)
            {
                CreateErrorLog(ex);
                if (_throwError)
                    throw;
            }
            finally
            {
                if (_isFaulted && _faultException != null && Context.IsInTransaction)
                {
                    CreateErrorLog(_faultException);
                    ClearParameters(false);
                    throw _faultException;
                }
                else
                {
                    CreateLog();
                    ClearParameters();
                }
            }
        }

        protected IOrganizationService GetService(Guid? userId, bool recreate = false)
        {
            Guid userGuid = userId.HasValue ? userId.Value : new Guid("FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF");

            IOrganizationService service = null;
            if (!_serviceCollection.TryGetValue(userGuid, out service) || service == null || recreate)
            {
                service = ServiceFactory.CreateOrganizationService(userId);
                if (_serviceCollection.ContainsKey(userGuid))
                    _serviceCollection.Remove(userGuid);
                _serviceCollection.Add(userGuid, service);
            }
            return service;
        }

        protected IOrganizationService GetService(UserType type)
        {
            Guid? userId = GetUserId(type);
            return GetService(userId);
        }

        private void ClearParameters(bool clearFaultException = true)
        {
            _logBuilder = new StringBuilder();
            _logged = false;
            TransactionId = Guid.Empty;
            _sequence = 0;
            _serviceCollection = new Dictionary<Guid, IOrganizationService>();
            TracingService = null;
            Context = null;
            ServiceFactory = null;
            _isFaulted = false;
            if (clearFaultException)
                _faultException = null;
        }

        #region Logging

        protected void CreateLog()
        {
            if (!_logged && _enableLogging)
            {
                CreateLog(string.Empty, Constants.LogMode.Debug);
            }
        }

        protected void CreateErrorLog(Exception ex)
        {
            CreateLog(GenerateErrorMessage(ex), Constants.LogMode.Error);
        }

        protected void AppendLog(string format, params object[] args)
        {
            if (_logBuilder == null)
                _logBuilder = new StringBuilder();
            _logBuilder.AppendFormatLine(format, args);

            if (args.Length > 0)
                TracingService.Trace(format, args);
            else
                TracingService.Trace(format);
        }

        private void CreateLog(string internalLog, Constants.LogMode logMode)
        {
            try
            {
                Entity log = new Entity("ols_pluginlog");
                log["ols_name"] = _pluginName;
                log["ols_type"] = new OptionSetValue((int)logMode);

                log["ols_depth"] = Context.Depth;
                log["ols_initiatinguserid"] = new EntityReference("systemuser", Context.InitiatingUserId);
                log["ols_isexecutingoffline"] = Context.IsExecutingOffline;
                log["ols_isintransaction"] = Context.IsInTransaction;
                log["ols_isofflineplayback"] = Context.IsOfflinePlayback;
                log["ols_isolationmode"] = Context.IsolationMode;
                log["ols_messagename"] = Context.MessageName;
                log["ols_mode"] = Context.Mode;
                log["ols_operationcreatedon"] = Context.OperationCreatedOn;
                log["ols_primaryentityid"] = Context.PrimaryEntityId.ToString();
                log["ols_primaryentityname"] = Context.PrimaryEntityName;
                log["ols_secondaryentityname"] = Context.SecondaryEntityName;
                log["ols_userid"] = new EntityReference("systemuser", Context.UserId);
                log["ols_stage"] = Context.Stage;
                log["ols_sharedvariables"] = Context.CorrelationId.ToString();

                log["ols_transactionid"] = TransactionId.ToString();
                log["ols_sequence"] = _sequence;

                if (_logImages)
                {
                    log["ols_preimages"] = GenerateTextFromEntityImages(Context.PreEntityImages);
                    log["ols_postimages"] = GenerateTextFromEntityImages(Context.PostEntityImages);
                }

                if (_logInputParameters)
                {
                    log["ols_inputparameters"] = GenerateTextFromParameters(Context.InputParameters);
                }

                log["ols_sharedvariables"] = GenerateTextFromSharedVariables(Context.SharedVariables);

                //if (_logImages)
                //{
                //    log["ols_preimages"] = GenerateTextForCRMTypes(Context.PreEntityImages);
                //    log["ols_postimages"] = GenerateTextForCRMTypes(Context.PostEntityImages);
                //}

                //if (_logInputParameters)
                //{
                //    log["ols_inputparameters"] = GenerateTextForCRMTypes(Context.InputParameters);
                //}

                //log["ols_sharedvariables"] = GenerateTextForCRMTypes(Context.SharedVariables);

                log["ols_log"] = _logBuilder.ToString() + internalLog;

                ExecuteMultipleRequest req = new ExecuteMultipleRequest
                {
                    Requests = new OrganizationRequestCollection(),
                    Settings = new ExecuteMultipleSettings { ContinueOnError = true }
                };
                req.Requests.Add(new CreateRequest { Target = log });

                ExecuteMultipleResponse resp = (ExecuteMultipleResponse)Execute(UserType.System, req);
                _logged = true;
            }
            catch (Exception) { throw; }
        }

        private string GenerateErrorMessage(Exception serviceException)
        {
            StringBuilder messageBuilder = new StringBuilder();
            try
            {
                try
                {
                    throw serviceException;
                }
                catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
                {
                    messageBuilder.AppendFormatLine("Fault Exception");
                    messageBuilder.AppendFormatLine("---------------");
                    messageBuilder.AppendFormatLine("Timestamp: {0}", ex.Detail.Timestamp);
                    messageBuilder.AppendFormatLine("Code: {0}", ex.Detail.ErrorCode);
                    messageBuilder.AppendFormatLine("Message: {0}", ex.Detail.Message);
                    messageBuilder.AppendFormatLine("Inner Fault: {0}", null == ex.Detail.InnerFault ? "No Inner Fault" : ex.Detail.InnerFault.Message);
                }
                catch (System.TimeoutException ex)
                {
                    messageBuilder.AppendFormatLine("Timeout Exception");
                    messageBuilder.AppendFormatLine("-----------------");
                    messageBuilder.AppendFormatLine("Message: {0}", ex.Message);
                    messageBuilder.AppendFormatLine("Stack Trace: {0}", ex.StackTrace);
                    messageBuilder.AppendFormatLine("Inner Fault: {0}", null == ex.InnerException ? "No Inner Fault" : ex.InnerException.Message);
                }
                catch (System.Exception ex)
                {
                    messageBuilder.AppendFormatLine("Exception");
                    messageBuilder.AppendFormatLine("---------");
                    messageBuilder.AppendFormatLine(ex.Message);

                    // Display the details of the inner exception.
                    if (ex.InnerException != null)
                    {
                        messageBuilder.AppendFormatLine(ex.InnerException.Message);

                        FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe = ex.InnerException
                            as FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>;
                        if (fe != null)
                        {
                            messageBuilder.AppendFormatLine("Timestamp: {0}", fe.Detail.Timestamp);
                            messageBuilder.AppendFormatLine("Code: {0}", fe.Detail.ErrorCode);
                            messageBuilder.AppendFormatLine("Message: {0}", fe.Detail.Message);
                            messageBuilder.AppendFormatLine("Trace: {0}", fe.Detail.TraceText);
                            messageBuilder.AppendFormatLine("Inner Fault: {0}",
                                null == fe.Detail.InnerFault ? "No Inner Fault" : fe.Detail.InnerFault.Message);
                        }
                    }
                }
                finally
                {
                    messageBuilder.AppendFormatLine("StackTrace");
                    messageBuilder.AppendFormatLine("---------");
                    messageBuilder.AppendFormatLine(serviceException.StackTrace);

                }
                //messageBuilder.AppendFormatLine("The Exception is:-");
                //messageBuilder.AppendFormatLine("Exception :: " + serviceException.ToString());
                //if (serviceException.InnerException != null)
                //{
                //    messageBuilder.AppendFormatLine("InnerException :: " + serviceException.InnerException.ToString());
                //}
                return messageBuilder.ToString();
            }
            catch
            {
                messageBuilder.AppendFormatLine("Exception:: Unknown Exception.");
                return messageBuilder.ToString();
            }
        }

        #endregion Logging

        #region Common Static Methods

        public static T GetAttributeValue<T>(string attributeLogicalName, Entity primary, Entity secondary = null) where T : new()
        {
            var value = primary != null && primary.Contains(attributeLogicalName) ? primary.GetAttributeValue<T>(attributeLogicalName) :
                 secondary != null && secondary.Contains(attributeLogicalName) ? secondary.GetAttributeValue<T>(attributeLogicalName) : default(T);
            if (value == null)
                value = new T();
            return value;
        }

        public static object GetSharedVaribleValue(IPluginExecutionContext context, string key)
        {
            if (context != null)
            {
                if (context.SharedVariables.Contains(key))
                {
                    return context.SharedVariables[key];
                }
                else
                {
                    return GetSharedVaribleValue(context.ParentContext, key);
                }
            }
            else
            {
                return string.Empty;
            }
        }

        private static string GenerateTextFromSharedVariables(ParameterCollection sharedVariables)
        {
            StringBuilder imageBuilder = new StringBuilder();
            if (sharedVariables != null && sharedVariables.Count > 0)
            {
                foreach (var variable in sharedVariables)
                {
                    if (variable.Value == null)
                    {
                        imageBuilder.AppendFormatLine(variable.Key + " : Null");
                        continue;
                    }
                    imageBuilder.AppendFormatLine(variable.Key + " : " + GetTextValue(variable.Value));
                }
            }
            return imageBuilder.ToString();
        }

        private static string GenerateTextFromEntityImages(EntityImageCollection images)
        {
            StringBuilder imageBuilder = new StringBuilder();
            if (images != null && images.Count > 0)
            {
                foreach (var image in images)
                {
                    if (image.Value == null)
                        continue;
                    string name = "Name: " + image.Key;
                    imageBuilder.AppendFormatLine("Name: " + image.Key);
                    imageBuilder.AppendFormatLine(GenerateCharacter('-', name.Length));
                    imageBuilder.AppendFormatLine(GenerateTextFromEntity(image.Value));
                    imageBuilder.AppendLine();
                }
            }
            return imageBuilder.ToString();
        }

        protected static string GenerateTextFromEntity(Entity record)
        {
            StringBuilder imageBuilder = new StringBuilder();
            foreach (var att in record.Attributes)
            {
                if (att.Value != null)
                    imageBuilder.AppendFormatLine(att.Key + " (" + att.Value.GetType().Name + ") : " + GetTextValue(att.Value));
                else
                    imageBuilder.AppendFormatLine(att.Key + " : NULL");

            }
            return imageBuilder.ToString();
        }

        private static string GenerateTextFromParameters(ParameterCollection inputParameters)
        {
            StringBuilder imageBuilder = new StringBuilder();
            if (inputParameters != null && inputParameters.Count > 0)
            {
                foreach (var parameter in inputParameters)
                {
                    if (parameter.Value == null)
                    {
                        imageBuilder.AppendFormatLine(parameter.Key + " : NULL");
                        continue;
                    }
                    imageBuilder.AppendFormatLine(parameter.Key + " (" + parameter.Value.GetType().Name + ")");

                    if (parameter.Value is Entity)
                    {
                        imageBuilder.AppendLine(GenerateTextFromEntity((Entity)parameter.Value));
                    }
                    else
                    {
                        imageBuilder.AppendLine(GetTextValue(parameter.Value));
                    }
                    imageBuilder.AppendLine();
                }
            }
            return imageBuilder.ToString();
        }

        private static string GetTextValue(object value)
        {
            if (value == null)
            {
                return "NULL";
            }
            var type = value.GetType();
            if (type == typeof(OptionSetValue))
            {
                return ((OptionSetValue)value).Value.ToString();
            }
            else if (type == typeof(EntityReference))
            {
                return ((EntityReference)value).GetNameId();
            }
            else
            {
                return value.ToString();
            }
        }

        private static string GenerateCharacter(char character, int count)
        {
            StringBuilder str = new StringBuilder();
            for (int i = 0; i < str.Length; i++)
            {
                str.Append(character);
            }
            return str.ToString();
        }

        private static string UpdateFetchXmlWithPagingInfo(string xml, string cookie, int page, int count)
        {
            var stringReader = new StringReader(xml);
            var reader = new XmlTextReader(stringReader);
            var doc = new XmlDocument();
            doc.Load(reader);
            return UpdateFetchXmlWithPagingInfo(doc, cookie, page, count);
        }

        private static string UpdateFetchXmlWithPagingInfo(XmlDocument doc, string cookie, int page, int count)
        {
            if (doc.DocumentElement != null)
            {
                XmlAttributeCollection attrs = doc.DocumentElement.Attributes;
                if (cookie != null)
                {
                    XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                    pagingAttr.Value = cookie;
                    attrs.Append(pagingAttr);
                }
                XmlAttribute pageAttr = doc.CreateAttribute("page");
                pageAttr.Value = System.Convert.ToString(page);
                attrs.Append(pageAttr);
                XmlAttribute countAttr = doc.CreateAttribute("count");
                countAttr.Value = System.Convert.ToString(count);
                attrs.Append(countAttr);
            }
            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);
            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();
            return sb.ToString();
        }

        private static string FormatJSON(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            text = text.Replace(System.Environment.NewLine, string.Empty).Replace("\t", string.Empty);

            var offset = 0;
            var output = new StringBuilder();
            Action<StringBuilder, int> tabs = (sb, pos) => { for (var i = 0; i < pos; i++) { sb.Append("  "); } };
            Func<string, int, Nullable<Char>> previousNotEmpty = (s, i) =>
            {
                if (string.IsNullOrEmpty(s) || i <= 0) return null;

                Nullable<Char> prev = null;

                while (i > 0 && prev == null)
                {
                    prev = s[i - 1];
                    if (prev.ToString() == " ") prev = null;
                    i--;
                }

                return prev;
            };
            Func<string, int, Nullable<Char>> nextNotEmpty = (s, i) =>
            {
                if (string.IsNullOrEmpty(s) || i >= (s.Length - 1)) return null;

                Nullable<Char> next = null;
                i++;

                while (i < (s.Length - 1) && next == null)
                {
                    next = s[i++];
                    if (next.ToString() == " ") next = null;
                }

                return next;
            };

            var inQuote = false;
            var ignoreQuote = false;

            for (var i = 0; i < text.Length; i++)
            {
                var chr = text[i];

                if (chr == '"' && !ignoreQuote) inQuote = !inQuote;
                if (chr == '\'' && inQuote) ignoreQuote = !ignoreQuote;
                if (inQuote)
                {
                    output.Append(chr);
                }
                else if (chr.ToString() == "{")
                {
                    offset++;
                    output.Append(chr);
                    output.Append(System.Environment.NewLine);
                    tabs(output, offset);
                }
                else if (chr.ToString() == "}")
                {
                    offset--;
                    output.Append(System.Environment.NewLine);
                    tabs(output, offset);
                    output.Append(chr);

                }
                else if (chr.ToString() == ",")
                {
                    output.Append(chr);
                    output.Append(System.Environment.NewLine);
                    tabs(output, offset);
                }
                else if (chr.ToString() == "[")
                {
                    output.Append(chr);

                    var next = nextNotEmpty(text, i);

                    if (next != null && next.ToString() != "]")
                    {
                        offset++;
                        output.Append(System.Environment.NewLine);
                        tabs(output, offset);
                    }
                }
                else if (chr.ToString() == "]")
                {
                    var prev = previousNotEmpty(text, i);

                    if (prev != null && prev.ToString() != "[")
                    {
                        offset--;
                        output.Append(System.Environment.NewLine);
                        tabs(output, offset);
                    }

                    output.Append(chr);
                }
                else
                    output.Append(chr);
            }

            return output.ToString().Trim();
        }

        private static string GenerateTextForCRMTypes(object obj)
        {
            if (obj == null)
                return string.Empty;
            try
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(obj.GetType());
                    serializer.WriteObject(stream, obj);
                    var jsonText = Encoding.Default.GetString(stream.ToArray());

                    return obj.GetType().ToString() + Environment.NewLine + FormatJSON(jsonText);
                }
            }
            catch (Exception)
            {

                return obj.GetType().ToString();
            }
            //return FormatJSON(obj.SerializeByJson(Util.KnownTypes));
        }

        //private static string GenerateTextFromQueryExpressionTypes(object obj)
        //{
        //    return FormatJSON(obj.SerializeByJson(KnownTypesProvider.QueryExpressionKnownTypes));
        //}

        #endregion Common Static Methods

        #region CRM Service Methods

        protected void Associate(Guid? userId, string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            AssociateRequest request = new AssociateRequest
            {
                Target = new EntityReference(entityName, entityId),
                Relationship = relationship,
                RelatedEntities = relatedEntities
            };
            Execute(userId, request);
        }

        protected Guid Create(Guid? userId, Entity entity)
        {
            CreateRequest request = new CreateRequest
            {
                Target = entity
            };
            return ((CreateResponse)Execute(userId, request)).id;
        }

        protected void Delete(Guid? userId, string entityName, Guid id)
        {
            IOrganizationService service = GetService(userId);
            DeleteRequest request = new DeleteRequest
            {
                Target = new EntityReference(entityName, id),
            };
            Execute(userId, request);
        }

        protected void Disassociate(Guid? userId, string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            IOrganizationService service = GetService(userId);
            DisassociateRequest request = new DisassociateRequest
            {
                Target = new EntityReference(entityName, entityId),
                Relationship = relationship,
                RelatedEntities = relatedEntities
            };
            Execute(userId, request);
        }

        protected OrganizationResponse Execute(Guid? userId, OrganizationRequest request)
        {
            IOrganizationService service = GetService(userId);
            try
            {
                AppendLog("Executing Message : {0} UserId: {1}", request.RequestName, userId ?? Guid.Empty);
                OrganizationResponse resp = service.Execute(request);
                AppendLog("Excecuting Message Successful");
                return resp;
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                _faultException = ex;
                _isFaulted = true;
                AppendLog("Error :" + ex.Message);
                AppendLog("Request :\n" + GenerateTextForCRMTypes(request));
                throw;
            }
            catch (Exception ex)
            {
                _faultException = ex;
                _isFaulted = true;
                AppendLog("Error :" + ex.Message);
                AppendLog("Request :\n" + GenerateTextForCRMTypes(request));
                throw;
            }
        }

        protected Entity Retrieve(Guid? userId, string entityName, Guid id, ColumnSet columnSet)
        {
            IOrganizationService service = GetService(userId);
            RetrieveRequest request = new RetrieveRequest
            {
                Target = new EntityReference(entityName, id),
                ColumnSet = columnSet,
            };
            return ((RetrieveResponse)Execute(userId, request)).Entity;
        }

        protected EntityCollection RetrieveMultiple(Guid? userId, QueryBase query)
        {
            IOrganizationService service = GetService(userId);
            RetrieveMultipleRequest request = new RetrieveMultipleRequest
            {
                Query = query
            };
            return ((RetrieveMultipleResponse)Execute(userId, request)).EntityCollection;

        }

        protected void Update(Guid? userId, Entity record)
        {
            IOrganizationService service = GetService(userId);
            UpdateRequest req = new UpdateRequest
            {
                Target = record
            };
            Execute(userId, req);
        }

        protected Guid? GetUserId(UserType userType)
        {
            Guid? userId = null;
            switch (userType)
            {
                case UserType.User:
                    userId = Context.UserId;
                    break;

                case UserType.InitiatingUser:
                    userId = Context.InitiatingUserId;
                    break;

            }
            return userId;
        }

        protected void Associate(UserType userType, string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            Associate(GetUserId(userType), entityName, entityId, relationship, relatedEntities);
        }

        protected Guid Create(UserType userType, Entity entity)
        {
            return Create(GetUserId(userType), entity);
        }

        protected void Delete(UserType userType, string entityName, Guid id)
        {
            Delete(GetUserId(userType), entityName, id);
        }

        protected void Disassociate(UserType userType, string entityName, Guid entityId, Relationship relationship, EntityReferenceCollection relatedEntities)
        {
            Disassociate(GetUserId(userType), entityName, entityId, relationship, relatedEntities);
        }

        protected OrganizationResponse Execute(UserType userType, OrganizationRequest request)
        {
            return Execute(GetUserId(userType), request);
        }

        protected Entity Retrieve(UserType userType, string entityName, Guid id, ColumnSet columnSet)
        {
            return Retrieve(GetUserId(userType), entityName, id, columnSet);
        }

        protected EntityCollection RetrieveMultiple(UserType userType, QueryBase query)
        {
            return RetrieveMultiple(GetUserId(userType), query);
        }

        protected void Update(UserType userType, Entity record)
        {
            Update(GetUserId(userType), record);
        }

        protected List<Entity> RetrieveAll(UserType userType, QueryExpression query)
        {
            var allRecords = new List<Entity>();
            var moreRecords = false;

            query.PageInfo.PageNumber = 1;
            query.PageInfo.PagingCookie = null;

            do
            {
                EntityCollection coll = RetrieveMultiple(userType, query);
                allRecords.AddRange(coll.Entities);
                if (coll.MoreRecords)
                {
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = coll.PagingCookie;
                }
                moreRecords = coll.MoreRecords;
            } while (moreRecords);

            return allRecords;
        }

        protected List<Entity> RetrieveAll(UserType userType, string fetchXml)
        {
            var totalRecords = new List<Entity>();
            var fetchCount = 5000;
            var pageNumber = 1;
            string pagingCookie = null;
            while (true)
            {
                string xml = UpdateFetchXmlWithPagingInfo(fetchXml, pagingCookie, pageNumber, fetchCount);
                var fetchRequest = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };
                var returnCollection = ((RetrieveMultipleResponse)Execute(userType, fetchRequest)).EntityCollection;
                if (returnCollection.Entities.Count > 0)
                {
                    totalRecords.AddRange(returnCollection.Entities);
                }
                if (returnCollection.MoreRecords)
                {
                    pageNumber++;
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    break;
                }
            }
            return totalRecords;
        }

        #endregion CRM Service Methods

        public void Dispose()
        {
        }
    }
}
