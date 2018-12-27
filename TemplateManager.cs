using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Yandex.Money.CRM.Common;

namespace Templator
{
    public class TemplateManager
    {

        private readonly string FIELD_PLUG_PATTERN = @"\%;.([^\%;])+;\%";

        private readonly IOrganizationService service;
        private readonly string templateName;
        private readonly string baseUrl;
        private readonly Guid? userTemplateId;
        private readonly Dictionary<string, OptionSetMetadata> OPTION_SET_VALUES;
        private Dictionary<string, QueryExpression> queriesForRelated;



        public TemplateManager(IOrganizationService _service, string _templateName, string _baseUrl)
        {
            service = _service;
            baseUrl = _baseUrl;
            templateName = _templateName;
            OPTION_SET_VALUES = new Dictionary<string, OptionSetMetadata>();
            queriesForRelated = new Dictionary<string, QueryExpression>();
        }


        public TemplateManager(IOrganizationService _service, Guid _userTemplateId, string _baseUrl)
        {
            service = _service;
            baseUrl = _baseUrl;
            userTemplateId = _userTemplateId;
            OPTION_SET_VALUES = new Dictionary<string, OptionSetMetadata>();
            queriesForRelated = new Dictionary<string, QueryExpression>();
        }

        private Entity GetTemplate(string templateName)
        {
            var query = new QueryExpression("template") { ColumnSet = new ColumnSet("subject", "body") };
            query.Criteria.AddCondition("title", ConditionOperator.Equal, templateName);
            return service.RetrieveMultiple(query).Entities.Single();
        }

        public void FillEmailByTemplate(EntityReference recordRef)
        {
            FillEmailByTemplate(recordRef.LogicalName, recordRef.Id);
        }

        public void FillEmailByTemplate(Entity record)
        {
            FillEmailByTemplate(record.LogicalName, record.Id);
        }


        public void FillEmailByTemplate(EntityReference recordRef, Dictionary<string, string> fillDictionary)
        {
            FillEmailByTemplate(recordRef.LogicalName, recordRef.Id, fillDictionary);
        }

        public void FillEmailByTemplate(Entity record, Dictionary<string, string> fillDictionary)
        {
            FillEmailByTemplate(record.LogicalName, record.Id, fillDictionary);
        }

        public Entity FillEmailByTemplate(string logicalName, Guid recordId, Dictionary<string, string> fillDictionary)
        {
            var email = FillEmailByTemplate(logicalName, recordId);
            var subject = email.GetAttributeValue<string>("subject");
            var description = email.GetAttributeValue<string>("description");
            email["subject"] = FillByDictionary(subject, fillDictionary);
            email["description"] = FillByDictionary(description, fillDictionary);
            return email;
        }

        private string FillByDictionary(string str, Dictionary<string, string> fillDictionary)
        {
            foreach (KeyValuePair<string, string> pair in fillDictionary)
            {
                str = str.Replace("&;" + pair.Key + ";&", pair.Value);
                str = str.Replace("&amp;;" + pair.Key + ";&amp;", pair.Value);
            }
            return str;
        }


        private Entity GetUserTemplate()
        {
            var userTemplate = service.Retrieve("ym_user_email_template", userTemplateId.Value, new ColumnSet("ym_clean_template", "ym_clean_subject"));
            return userTemplate;
        }


        private void GetSubjectAndBody(out string subject, out string body)
        {
            if (!string.IsNullOrWhiteSpace(templateName))
            {
                var template = GetTemplate(templateName);
                subject = GetTemplateElement(template, "subject");
                body = GetTemplateElement(template, "body");
            }
            else if (userTemplateId.HasValue)
            {
                var userTemplate = GetUserTemplate();
                subject = userTemplate.GetAttributeValue<string>("ym_clean_subject") ?? string.Empty;
                body = userTemplate.GetAttributeValue<string>("ym_clean_template") ?? string.Empty;
            }
            else
                throw new Exception("Can't find template");

            subject = ReplaceQuote(subject);
            body = ReplaceQuote(body);
        }

        public Entity FillEmailByTemplate()
        {
            string subject = string.Empty;
            string body = string.Empty;
            GetSubjectAndBody(out subject, out body);
            return new Entity("email") { ["subject"] = subject, ["description"] = body };
        }

        public Entity FillEmailByTemplate(string logicalName, Guid recordId)
        {
            var record = service.Retrieve(logicalName, recordId, new ColumnSet(true));
            string subject = string.Empty;
            string body = string.Empty;
            GetSubjectAndBody(out subject, out body);
            subject = ReplacePlugs(subject, record);
            body = ReplacePlugs(body, record);
            return new Entity("email") { ["subject"] = subject, ["description"] = body };
        }

        private string GetTemplateElement(Entity template, string fieldName)
        {
            var regExp = new Regex(@"<!\[CDATA\[(.|\n)*?\]\]>");
            var fieldValue = template.GetAttributeValue<string>(fieldName);
            var value = regExp.Match(fieldValue).Value;
            value = value.Replace("<![CDATA[", "").Replace("]]>", "");
            return value;
        }

        private List<string> GetFieldPlugs(string str)
        {
            var regExp = new Regex(FIELD_PLUG_PATTERN);
            var matches = regExp.Matches(str);
            var fieldPlugs = new List<string>();
            foreach (Match match in matches)
                fieldPlugs.Add(match.Value);
            return fieldPlugs;
        }



        public string ReplaceQuote(string str)
        {
            return str.Replace("&quot;", "'");
        }

        private string GetValueFromConnectedEntity(string field, Entity record)
        {
            var retrieveDetails = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.All,
                LogicalName = record.LogicalName
            };
            RetrieveEntityResponse retrieveEntityResponseObj = (RetrieveEntityResponse)service.Execute(retrieveDetails);
            Microsoft.Xrm.Sdk.Metadata.EntityMetadata metadata = retrieveEntityResponseObj.EntityMetadata;
            string value = null;
            if (field.Contains("."))
            {
                var split = field.Split('.');
                var currentField = split.First();
                var requiredField = split[1];
                field = string.Join(".", split.Skip(1));
                if (record.HasValue<EntityReference>(currentField))
                {
                    var connectedRecordRef = record.GetAttributeValue<EntityReference>(currentField);
                    return GetValueFromConnectedEntity(field, service.Retrieve(connectedRecordRef.LogicalName, connectedRecordRef.Id, new ColumnSet(requiredField)));
                }
            }
            else
                value = GetValueFromEntity(field, record, metadata);
            return value ?? string.Empty;
        }

        public string GetOptionsetText(string entityName, string attributeName, int optionSetValue, EntityMetadata metadata)
        {
            if (!OPTION_SET_VALUES.ContainsKey(attributeName))
            {

                Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata picklistMetadata = metadata.Attributes.FirstOrDefault(attribute => String.Equals
                    (attribute.LogicalName, attributeName, StringComparison.OrdinalIgnoreCase)) as Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata;
                OPTION_SET_VALUES.Add(attributeName, picklistMetadata.OptionSet);
            }
            Microsoft.Xrm.Sdk.Metadata.OptionSetMetadata options = OPTION_SET_VALUES[attributeName];
            IList<OptionMetadata> OptionsList = (from o in options.Options where o.Value.Value == optionSetValue select o).ToList();
            string optionsetLabel = (OptionsList.SingleOrDefault())?.Label.UserLocalizedLabel.Label;
            return optionsetLabel;
        }

        private string GetValueFromEntity(string field, Entity record, EntityMetadata metadata)
        {
            string value = null;
            if (record.Contains(field) && record[field] != null)
            {
                if (record[field] is OptionSetValue)
                {
                    value = GetOptionsetText(record.LogicalName, field, record.GetAttributeValue<OptionSetValue>(field).Value, metadata);
                }
                else if (record[field] is EntityReference)
                {
                    if (!string.IsNullOrWhiteSpace(record.GetAttributeValue<EntityReference>(field).Name))
                        value = record.GetAttributeValue<EntityReference>(field).Name;
                    else
                    {
                        var fieldToRetrieve = field.StartsWith("ym_") ? "ym_name" : "name";
                        value = service.Retrieve(record.GetAttributeValue<EntityReference>(field).LogicalName, record.GetAttributeValue<EntityReference>(field).Id, new ColumnSet(fieldToRetrieve)).GetAttributeValue<string>(fieldToRetrieve);
                    }
                }
                else
                {
                    var attribute = metadata.Attributes.FirstOrDefault(x => x.LogicalName == field);
                    if (attribute is StringAttributeMetadata)
                    {
                        var formatName = (attribute as StringAttributeMetadata)?.FormatName?.Value;
                        if (!string.IsNullOrWhiteSpace(formatName) && string.Equals(formatName, "url", StringComparison.InvariantCultureIgnoreCase))
                            value = string.IsNullOrWhiteSpace(record[field]?.ToString()) ? string.Empty : $"<a href='{record[field]?.ToString()}'>{record[field]?.ToString()}</a>";
                        else
                            value = record[field]?.ToString();
                    }

                    if (attribute is DateTimeAttributeMetadata)
                    {
                        var offset = TimeZoneInfo.Local.GetUtcOffset(DateTime.UtcNow).Hours;
                        if (record.Contains(field) && record[field] != null && DateTime.TryParse(record[field]?.ToString(), out DateTime resultDate))
                            value = resultDate.AddHours(offset).ToString("dd.MM.yyyy HH:mm");
                    }
                    else
                        value = record[field]?.ToString();
                }
            }
            return value ?? string.Empty;
        }

        private string ReplacePlugs(string str, Entity record)
        {
            var plugs = GetFieldPlugs(str);
            plugs.ForEach(plug =>
            {
                var clearedPlug = Regex.Replace(plug.Replace("%;", "").Replace(";%", ""), "<[^>]*>", string.Empty).Trim();
                if (clearedPlug.StartsWith("REF_RECORD_URL"))
                {
                    var split = clearedPlug.Split('|');
                    clearedPlug = split.First().Replace("REF_RECORD_URL_", "");
                    if (!string.IsNullOrWhiteSpace(clearedPlug) && record.HasValue<EntityReference>(clearedPlug))
                    {
                        var connectedRecordRef = record.GetAttributeValue<EntityReference>(clearedPlug);
                        var linkText = split.Last();
                        var url = CRMHelper.FormatUrlToRecord(baseUrl, connectedRecordRef.LogicalName, connectedRecordRef.Id);
                        str = str.Replace(plug, $" <a href=\"{url}\">{linkText}</a> ");
                    }
                }
                else if (clearedPlug.StartsWith("RECORD_URL"))
                {
                    var split = clearedPlug.Split('|');
                    var linkText = split.Last();
                    var url = CRMHelper.FormatUrlToRecord(baseUrl, record.LogicalName, record.Id);
                    str = str.Replace(plug, $" <a href=\"{url}\">{linkText}</a> ");
                }
                else if (clearedPlug.StartsWith("RELATED"))
                {
                    //One-To-Many
                    clearedPlug = clearedPlug.Replace("RELATED_", "");
                    var split = clearedPlug.Split('-');
                    var relatedEntityLogicalName = split[0];
                    var relatedEntityReferenceFieldLogicalName = split[1];
                    //var relatedEntityFieldLogicalName = split[2];
                    var relatedEntityFields = split.Skip(2).Where(x => x.StartsWith("ym_") || x.StartsWith("kc_")).ToArray();
                    QueryExpression query = queriesForRelated.ContainsKey(relatedEntityLogicalName + "-" + relatedEntityReferenceFieldLogicalName) ? queriesForRelated[relatedEntityLogicalName + "-" + relatedEntityReferenceFieldLogicalName] : null;
                    if (query != null) query.ColumnSet = new ColumnSet(relatedEntityFields);
                    var relatedRecords = query != null ? FindRelatedRecords(query) : FindRelatedRecords(relatedEntityLogicalName, relatedEntityReferenceFieldLogicalName, record.Id, relatedEntityFields);
                    var relatedRecordsStringBuilder = new StringBuilder();
                    var relatedRecordPlug = "%;" + string.Join(";%%;", split.Skip(2)) + ";%";
                    relatedRecordsStringBuilder.Append("<ul>");
                    relatedRecords.ForEach(relatedRecord =>
                    {
                        relatedRecordsStringBuilder.Append("<li>" + ReplacePlugs(relatedRecordPlug, relatedRecord) + "</li>");
                    });
                    relatedRecordsStringBuilder.Append("</ul>");
                    str = str.Replace(plug, relatedRecordsStringBuilder.ToString());
                }
                else
                    str = str.Replace(plug, GetValueFromConnectedEntity(clearedPlug, record));
            });
            return str;
        }

        public void AddFilter(string key, QueryExpression queryExpression)
        {
            if (queriesForRelated.ContainsKey(key)) queriesForRelated[key] = queryExpression;
            else queriesForRelated.Add(key, queryExpression);
        }

        private List<Entity> FindRelatedRecords(QueryExpression query)
        {
            return service.RetrieveMultiple(query).Entities.ToList();
        }

        private List<Entity> FindRelatedRecords(string relatedEntityLogicalName, string relatedEntityReferenceFieldLogicalName, Guid relatedEntityReferenceRecordId, string[] relatedEntityFields)
        {
            var query = new QueryExpression(relatedEntityLogicalName);
            query.ColumnSet = new ColumnSet(relatedEntityFields);
            query.Criteria.AddCondition(relatedEntityReferenceFieldLogicalName, ConditionOperator.Equal, relatedEntityReferenceRecordId);
            return service.RetrieveMultiple(query).Entities.ToList();
        }
    }
}

