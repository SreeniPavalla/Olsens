using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Common
{
    public static class Util
    {
        public static IEnumerable<Type> KnownTypes = new List<Type>()
         {
                typeof(AliasedValue),
                typeof(Dictionary<string, string>),
                typeof(Entity),
                typeof(Entity[]),
                typeof(ColumnSet),
                typeof(EntityReferenceCollection),
                typeof(QueryBase),
                typeof(QueryExpression),
                typeof(QueryExpression[]),
                typeof(LocalizedLabel[]),
                typeof(PagingInfo),
                typeof(Relationship),
                typeof(AttributePrivilegeCollection),
                typeof(RelationshipQueryCollection),


                typeof(bool),
                typeof(bool[]),
                typeof(int),
                typeof(int[]),
                typeof(string),
                typeof(string[]),
                typeof(string[][]),
                typeof(double),
                typeof(double[]),
                typeof(decimal),
                typeof(decimal[]),
                typeof(Guid),
                typeof(Guid[]),
                typeof(DateTime),
                typeof(DateTime[]),
                typeof(Money),
                typeof(Money[]),
                typeof(EntityReference),
                typeof(EntityReference[]),
                typeof(OptionSetValue),
                typeof(OptionSetValue[]),
                typeof(EntityCollection),
                typeof(Money),
                typeof(Label),
                typeof(LocalizedLabel),
                typeof(LocalizedLabelCollection),
                typeof(EntityMetadata[]),
                typeof(EntityMetadata),
                typeof(AttributeMetadata[]),
                typeof(AttributeMetadata),
                typeof(RelationshipMetadataBase[]),
                typeof(RelationshipMetadataBase),
                typeof(EntityFilters),
                typeof(OptionSetMetadataBase),
                typeof(OptionSetMetadataBase[]),
                typeof(OptionSetMetadata),
                typeof(BooleanOptionSetMetadata),
                typeof(OptionSetType),
                typeof(ManagedPropertyMetadata),
                typeof(ManagedPropertyMetadata[]),
                typeof(BooleanManagedProperty),
                typeof(AttributeRequiredLevelManagedProperty)
            };

        public static IEnumerable<Type> EntityKnownTypes;

        public static T GetAttributeValue<T>(string attributeLogicalName, Entity primary, Entity secondary = null) where T : new()
        {
            var value = primary != null && primary.Contains(attributeLogicalName) ? primary.GetAttributeValue<T>(attributeLogicalName) :
                 secondary != null && secondary.Contains(attributeLogicalName) ? secondary.GetAttributeValue<T>(attributeLogicalName) : default(T);
            if (value == null)
                value = new T();
            return value;
        }

        private static string GenerateContext(IPluginExecutionContext context)
        {
            var currentContext = context;
            var details = string.Empty;
            while (currentContext != null)
            {
                if (currentContext.Stage != 30)//Core Operation
                    details = "Entity: " + context.PrimaryEntityName + "Id: " + context.PrimaryEntityId + "Stage: " +
                               context.Stage + "Message Name: " + context.MessageName + Environment.NewLine + details;
                currentContext = context.ParentContext;
            }
            return details;
        }

        public static StringBuilder AppendFormatLine(this StringBuilder builder, string format, params object[] args)
        {
            return args != null && args.Length > 0 ? builder.AppendFormat(format + Environment.NewLine, args) : builder.AppendLine(format);
        }

        private static string CreateErrorMessage(Exception serviceException)
        {
            StringBuilder messageBuilder = new StringBuilder();
            try
            {
                messageBuilder.AppendFormatLine("The Exception is:-");
                messageBuilder.AppendFormatLine("Exception :: " + serviceException.ToString());
                if (serviceException.InnerException != null)
                {
                    messageBuilder.AppendFormatLine("InnerException :: " + serviceException.InnerException.ToString());
                }
                return messageBuilder.ToString();
            }
            catch
            {
                messageBuilder.AppendFormatLine("Exception:: Unknown Exception.");
                return messageBuilder.ToString();
            }
        }

        public static string GetNameId(this EntityReference reference)
        {
            return reference == null ? "Null" : "Id : " + reference.Id.ToString() + (String.IsNullOrEmpty(reference.Name) ? String.Empty : ", Name : " + reference.Name);
        }

        public static Entity CloneEntity(this Entity record)
        {
            return new Entity(record.LogicalName, record.Id);
        }

        public static Entity CreateEntity(this EntityReference record)
        {
            return new Entity(record.LogicalName, record.Id);
        }

        public static T GetValue<T>(this Entity entity, string attribute)
        {
            if (!entity.Attributes.Contains(attribute))
                return default(T);

            object attributeValue = entity[attribute];
            if (attributeValue is AliasedValue)
            {
                attributeValue = (attributeValue as AliasedValue).Value;
            }
            return (T)attributeValue;
        }

        private static byte[] ReadByteArray(Stream s)
        {
            byte[] rawLength = new byte[sizeof(int)];
            if (s.Read(rawLength, 0, rawLength.Length) != rawLength.Length)
            {
                throw new SystemException("Stream did not contain properly formatted byte array");
            }

            byte[] buffer = new byte[BitConverter.ToInt32(rawLength, 0)];
            if (s.Read(buffer, 0, buffer.Length) != buffer.Length)
            {
                throw new SystemException("Did not read byte array properly");
            }

            return buffer;
        }

        public static string GetUserNameById(IOrganizationService service, Guid userId)
        {
            string name = string.Empty;
            EntityCollection usersCol = null;
            QueryExpression query = new QueryExpression("systemuser");
            query.ColumnSet = new ColumnSet("fullname");
            query.Criteria.AddCondition(new ConditionExpression("systemuserid", ConditionOperator.Equal, userId));
            usersCol = service.RetrieveMultiple(query);

            if (usersCol != null && usersCol.Entities.Count > 0)
                name = usersCol.Entities.FirstOrDefault().Contains("fullname") ? usersCol.Entities.FirstOrDefault().GetAttributeValue<string>("fullname") : string.Empty;

            return name;
        }

        public static EntityReference GetBusinessUnitByName(IOrganizationService service, string recName)
        {
            EntityReference buRef = new EntityReference();
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='businessunit'>
                                    <attribute name='name' />
                                    <attribute name='createdon' />
                                    <order attribute='createdon' descending='false' />
                                    <filter type='and'>
                                      <condition attribute='name' operator='like' value='%{0}%' />
                                    </filter>
                                  </entity>
                                </fetch>";


            fetchXML = string.Format(fetchXML, recName);
            Entity FieldCollRec = service.RetrieveMultiple(new FetchExpression(fetchXML)).Entities.FirstOrDefault();

            if (FieldCollRec != null && FieldCollRec.Id != Guid.Empty)
                buRef = FieldCollRec.ToEntityReference();

            return buRef;

        }

        public static String FormatNum(Int32 Num, int Digits)
        {

            try
            {
                String numStr = Num.ToString();

                for (int i = numStr.Length; i < Digits; i++)
                {
                    numStr = "0" + numStr;
                }

                return numStr;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        #region Create Quote Methods
        public static void CreateQuote(Guid accountId, string prepaidNumber, IOrganizationService service)
        {
            try
            {
                Entity Ent = new Entity("quote");
                Ent["name"] = "Pre Paid No " + prepaidNumber + " Quote ";
                Ent["customerid"] = new EntityReference("account", accountId);
                Ent["pricelevelid"] = new EntityReference("pricelevel", GetPricelist("Olsens Prepaid", service));
                Guid QuoteID = service.Create(Ent);

                CreateQuoteDetail("CH002", QuoteID, service);
                CreateQuoteDetail("CH003", QuoteID, service);
                CreateQuoteDetail("CH009", QuoteID, service);
                CreateQuoteDetail("CH012", QuoteID, service);
                CreateQuoteDetail("CH014", QuoteID, service);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static Guid GetPricelist(String name, IOrganizationService service)
        {

            List<ConditionExpression> lst = new List<ConditionExpression>();

            try
            {
                QueryExpression qe = new QueryExpression("pricelevel");
                qe.ColumnSet = new ColumnSet("pricelevelid");
                qe.Criteria.AddCondition("name", ConditionOperator.Equal, name);
                Entity priceList = service.RetrieveMultiple(qe).Entities.FirstOrDefault();
                if (priceList != null)
                    return priceList.Id;
                else
                    return Guid.Empty;

            }
            catch (Exception e)
            {
                throw (e);
            }
            finally
            {
                //lst = null; cSet = null;
            }
        }
        public static void CreateQuoteDetail(String productNum, Guid quoteID, IOrganizationService service)
        {
            Entity Ent = new Entity("quotedetail");
            Entity Prod = GetProduct(productNum, service);
            try
            {
                if (Prod != null)
                {
                    Ent["productid"] = new EntityReference(Prod.LogicalName, Prod.Id);
                    if (Prod.Contains("defaultuomid"))
                        Ent["uomid"] = new EntityReference("uom", Prod.GetAttributeValue<EntityReference>("defaultuomid").Id);
                    if (Prod.Contains("price"))
                        Ent["priceperunit"] = Prod["price"];
                    Ent["quantity"] = new decimal(1.0);
                    Ent["quoteid"] = new EntityReference("quote", quoteID);
                    service.Create(Ent);
                }

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
        public static Entity GetProduct(String productnumber, IOrganizationService service)
        {
            try
            {
                QueryExpression qe = new QueryExpression("product");
                qe.ColumnSet = new ColumnSet("defaultuomid", "price", "ols_sequencenumber");
                qe.Criteria.AddCondition("productnumber", ConditionOperator.Equal, productnumber);
                return service.RetrieveMultiple(qe).Entities.FirstOrDefault();
            }
            catch (Exception e)
            {
                throw (e);
            }
            finally
            {
            }
        }
        #endregion

        #region Translate Fields for Fees Methods
        #region ServiceFees
        public static Money SetServiceFees(IOrganizationService svc, string entityName, EntityReference entity, Entity target)
        {

            if (entity == null) return new Money(0);

            String FeeType = GetFeeType(target);
            ColumnSet cSet = new ColumnSet(new string[] { FeeType });

            try
            {
                var ServicePlace = svc.Retrieve(entityName, entity.Id, cSet);

                if (ServicePlace != null && ServicePlace.Contains(FeeType) && ServicePlace[FeeType] != null)
                    return new Money(Convert.ToDecimal(ServicePlace[FeeType]));
                else
                    return new Money(0);

            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                cSet = null;
            }
        }
        #endregion
        #region GetFeeType
        public static String GetFeeType(Entity target)
        {
            try
            {
                if (GetDeceasedAge(target) < 1)
                {
                    return "ols_0to1yrsfees";
                }
                else if (GetDeceasedAge(target) <= 3)
                {
                    return "ols_1to3yrsfees";
                }
                else if (GetDeceasedAge(target) <= 6)
                {
                    return "ols_4to6yrsfees";
                }
                else if (GetDeceasedAge(target) <= 12)
                {
                    return "ols_7to12yrsfees";
                }
                else
                {
                    return "ols_servicefee";
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
        #endregion
        #region GetDeceasedAge
        public static int GetDeceasedAge(Entity target)
        {
            try
            {
                if (!target.Contains("ols_ageatdod") || string.IsNullOrEmpty(target.GetAttributeValue<string>("ols_ageatdod")))
                    return 20; //return as adult if no age calculated on form 

                if (!string.IsNullOrEmpty(target.GetAttributeValue<string>("ols_ageatdod")) && target.GetAttributeValue<string>("ols_ageatdod").IndexOf("yrs") == -1)
                    return 0;
                else
                    return int.Parse(target.GetAttributeValue<string>("ols_ageatdod").Substring(0, target.GetAttributeValue<string>("ols_ageatdod").IndexOf("yrs")));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion
        #endregion

        public static string GetConfigSettingValue(string name, IOrganizationService service)
        {
            try
            {
                QueryExpression qe = new QueryExpression("ols_setting");
                qe.ColumnSet = new ColumnSet(true);
                qe.Criteria.AddCondition("ols_name", ConditionOperator.Equal, name);
                Entity configRec = service.RetrieveMultiple(qe).Entities.FirstOrDefault();
                if (configRec != null)
                {
                    return configRec.Contains("ols_value") ? configRec.GetAttributeValue<string>("ols_value") : string.Empty;
                }
                else return string.Empty;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static int GetBrandValue(IOrganizationService svc, Guid brandId)
        {
            var brandvalue = -1;
            List<ConditionExpression> lst = new List<ConditionExpression>();
            var brand = svc.Retrieve("pricelevel", brandId, new ColumnSet("ols_companyname"));

            if (brand != null && brand.Contains("ols_companyname"))
            {
                var brand_companyname = brand.GetAttributeValue<string>("ols_companyname").ToLower();
                return brand_companyname.IndexOf("olsen") != -1 ? 0 : brand_companyname.IndexOf("kelly") != -1 ? 1 : brand_companyname.IndexOf("kennedy") != -1 ? 2 : brand_companyname.IndexOf("walters") != -1 ? 3 : brand_companyname.LastIndexOf("walter") != -1 ? 4 : -1;
            }
            return brandvalue;
        }
        public static Guid GetQueueId(IOrganizationService svc, String name)
        {
            Entity ent;
            try
            {
                QueryExpression qe = new QueryExpression("queue");
                qe.ColumnSet = new ColumnSet("queueid");
                qe.Criteria.AddCondition("name", ConditionOperator.Equal, name);
                ent = svc.RetrieveMultiple(qe).Entities.FirstOrDefault();
                if (ent != null)
                {
                    return ent.Id;
                }
                else
                {
                    return Guid.Empty;
                }
            }
            catch (Exception e)
            {
                throw (e);
            }
        }
        #region Date Helper Methods
        public static DateTime LocalFromUTCUserDateTime(IOrganizationService svc, DateTime convertDate)
        {

            int? getTimeZoneCode = RetrieveCurrentUsersSettings(svc);
            DateTime localDateTime = RetrieveLocalTimeFromUTCTime(convertDate, getTimeZoneCode, svc);
            return localDateTime;
        }
        private static int? RetrieveCurrentUsersSettings(IOrganizationService service)

        {

            var currentUserSettings = service.RetrieveMultiple(

                new QueryExpression("usersettings")

                {

                    ColumnSet = new ColumnSet("timezonecode"),

                    Criteria = new FilterExpression

                    {

                        Conditions =

                        {

                            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)

                        }

                    }

                }).Entities[0].ToEntity<Entity>();

            //return time zone code

            return (int?)currentUserSettings.Attributes["timezonecode"];

        }
        private static DateTime RetrieveLocalTimeFromUTCTime(DateTime utcTime, int? timeZoneCode, IOrganizationService svc)
        {

            if (!timeZoneCode.HasValue)

                return DateTime.Now;

            var request = new LocalTimeFromUtcTimeRequest

            {

                TimeZoneCode = timeZoneCode.Value,

                UtcTime = utcTime.ToUniversalTime()

            };

            var response = (LocalTimeFromUtcTimeResponse)svc.Execute(request);

            return response.LocalTime;

        }
        #endregion

        #region Client Extensions

        private static XmlObjectSerializer CreateDataContractJsonSerializer(Type type, IEnumerable<Type> knownTypes)
        {
            return new DataContractJsonSerializer(type, knownTypes, 2147483647, true, null, false);
        }

        private static object DataContractJsonDeserialize(string text, Type type, IEnumerable<Type> knownTypes, Func<Type, IEnumerable<Type>, XmlObjectSerializer> create)
        {
            object obj;
            using (MemoryStream memoryStream = new MemoryStream(Encoding.Unicode.GetBytes(text)))
            {
                obj = create(type, knownTypes).ReadObject(memoryStream);
            }
            return obj;
        }

        private static string DataContractJsonSerialize(object value, Type type, IEnumerable<Type> knownTypes, Func<Type, IEnumerable<Type>, XmlObjectSerializer> create)
        {
            string end;
            using (MemoryStream memoryStream = new MemoryStream())
            {
                create(type, knownTypes).WriteObject(memoryStream, value);
                memoryStream.Position = (long)0;
                using (StreamReader streamReader = new StreamReader(memoryStream))
                {
                    end = streamReader.ReadToEnd();
                }
            }
            return end;
        }

        public static object DeserializeByJson(this string text, Type type, IEnumerable<Type> knownTypes)
        {
            return DataContractJsonDeserialize(text, type, knownTypes, new Func<Type, IEnumerable<Type>, XmlObjectSerializer>(CreateDataContractJsonSerializer));
        }

        public static string SerializeByJson(this object obj, IEnumerable<Type> knownTypes)
        {
            return DataContractJsonSerialize(obj, obj.GetType(), knownTypes, new Func<Type, IEnumerable<Type>, XmlObjectSerializer>(CreateDataContractJsonSerializer));
        }
        #endregion

        #region Helper Methods


        /// <summary>
        /// This method will generate a random number and returns
        /// </summary>
        /// <returns></returns>
        public static string GenerateRandomOTP()
        {
            string _numbers = "0123456789";
            int _OTPLength = 6;
            StringBuilder randomNumber = new StringBuilder(_OTPLength);
            Random random = new Random();
            for (var i = 0; i < _OTPLength; i++)
            {
                char rNum = _numbers[random.Next(0, _numbers.Length)];
                if (randomNumber.Length == 0 && rNum == '0')
                    --i;
                else
                    randomNumber.Append(rNum);
            }

            return randomNumber.ToString();
        }

        #endregion Helper Methods

      
    }
}
