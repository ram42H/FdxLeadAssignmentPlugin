using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAssignmentPlugin
{
    class CRMQueryExpression
    {
        string filterAttribute;
        ConditionOperator filterOperator;
        object filterValue;

        public CRMQueryExpression(string _attr, ConditionOperator _opr, object _val)
        {
            this.filterAttribute = _attr;
            this.filterOperator = _opr;
            this.filterValue = _val;
        }
        public static QueryExpression getQueryExpression(string _entityName, ColumnSet _columnSet, CRMQueryExpression[] _exp, LogicalOperator _filterOperator = LogicalOperator.And)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = _entityName;
            query.ColumnSet = _columnSet;
            query.Distinct = false;
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = _filterOperator;
            for (int i = 0; i < _exp.Length; i++)
            {
                query.Criteria.AddCondition(_exp[i].filterAttribute, _exp[i].filterOperator, _exp[i].filterValue);
            }

            return query;
        }

        public static QueryExpression getQueryWithChildExpression(string _entityName, ColumnSet _columnSet, CRMQueryExpression[] _exp, LogicalOperator _filterOperator = LogicalOperator.And, LogicalOperator _filterOperator2 = LogicalOperator.Or)
        {
            QueryExpression query = new QueryExpression();
            query.EntityName = _entityName;
            query.ColumnSet = _columnSet;
            query.Distinct = false;
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = _filterOperator;
            query.Criteria.AddCondition(_exp[0].filterAttribute, _exp[0].filterOperator, _exp[0].filterValue);
            FilterExpression childFilter = query.Criteria.AddFilter(_filterOperator2);
            for (int i = 1; i < _exp.Length; i++)
            {
                childFilter.AddCondition(_exp[i].filterAttribute, _exp[i].filterOperator, _exp[i].filterValue);
            }

            return query;
        }

        public static int GetOptionsetValue(IOrganizationService _service, string _optionsetName, string _optionsetSelectedText)
        {
            int optionsetValue = 0;
            try
            {
                RetrieveOptionSetRequest retrieveOptionSetRequest =
                    new RetrieveOptionSetRequest
                    {
                        Name = _optionsetName
                    };

                // Execute the request.
                RetrieveOptionSetResponse retrieveOptionSetResponse =
                    (RetrieveOptionSetResponse)_service.Execute(retrieveOptionSetRequest);

                // Access the retrieved OptionSetMetadata.
                OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;

                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();
                foreach (OptionMetadata optionMetadata in optionList)
                {
                    //If the value matches/....
                    if (optionMetadata.Label.UserLocalizedLabel.Label.ToString() == _optionsetSelectedText)
                    {
                        optionsetValue = (int)optionMetadata.Value;
                        break;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return optionsetValue;
        }
        public static string GetOptionsSetTextForValue(IOrganizationService service, string entityName, string attributeName, int selectedValue)
        {

            RetrieveAttributeRequest retrieveAttributeRequest = new
            RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };
            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            // Access the retrieved attribute.
            Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)
            retrieveAttributeResponse.AttributeMetadata;// Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
            string selectedOptionLabel = null;
            foreach (OptionMetadata oMD in optionList)
            {
                if (oMD.Value == selectedValue)
                {
                    selectedOptionLabel = oMD.Label.LocalizedLabels[0].Label.ToString();
                    break;
                }
            }
            return selectedOptionLabel;
        }

        public static EntityCollection GetLeadEntityCollection(LeadAssignment_Variables _variables, IOrganizationService _service, int _matchingCriteria)
        {
            EntityCollection entities = new EntityCollection ();
            QueryExpression queryExp = new QueryExpression();

            switch(_matchingCriteria)
            {
                case 1: //First name, last name and office phone on new lead matches First name, last name and office phone on existing lead
                    queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("firstname", ConditionOperator.Equal, _variables.firstName), new CRMQueryExpression("lastname", ConditionOperator.Equal, _variables.lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone) });
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 2: //First name, last name and office phone on new lead matches First name, last name and office phone on existing contact
                    queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("firstname", ConditionOperator.Equal, _variables.firstName), new CRMQueryExpression("lastname", ConditionOperator.Equal, _variables.lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone) });
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 3: //Office Phone on new lead matches office phone on existing lead
                    queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone) });
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 4: //Email on new lead matches email on existing lead
                    if (_variables.email != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, _variables.email) });
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 5: //Company name and zipcode on new lead matches company name and zipcode on an account
                    if (_variables.companyName != "" && _variables.zipcodetext != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, _variables.companyName), new CRMQueryExpression("address1_postalcode", ConditionOperator.Equal, _variables.zipcodetext) });
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 6: //Company name and zipcode on new lead matches company name and zipcode on a lead
                    if(_variables.companyName != "" && _variables.zipcodetext != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("companyname", ConditionOperator.Equal, _variables.companyName), new CRMQueryExpression("address1_postalcode", ConditionOperator.Equal, _variables.zipcodetext) });
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 7: //Email on new lead matches email on account
                    if (_variables.email != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, _variables.email) });
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 8: //Email on new lead matches email on existing contact
                    if (_variables.email != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, _variables.email) });
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 9: //Lastname and phone on new lead matches lastname and any phone on contact
                    queryExp = CRMQueryExpression.getQueryWithChildExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("lastname", ConditionOperator.Equal, _variables.lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.phone) });
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 10: //Lastname and phone on new lead matches lastname and any phone on lead
                    queryExp = CRMQueryExpression.getQueryWithChildExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("lastname", ConditionOperator.Equal, _variables.lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.phone) });
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 11: //Office Phone on new lead matches any phone on a lead
                    queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.phone) }, LogicalOperator.Or);
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 12: //Mobile Phone on new lead matches any phone on a lead
                    if (_variables.telephone1 != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.telephone1), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.telephone1), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.telephone1) }, LogicalOperator.Or);
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 13: //Office Phone on new lead matches any phone on account
                    queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.phone) }, LogicalOperator.Or);
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 14: //Mobile Phone on new lead matches any phone on account
                    if (_variables.telephone1 != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.telephone1), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.telephone1), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.telephone1) }, LogicalOperator.Or);
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 15: //Office Phone on new lead matches any phone on contact
                    queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.phone), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.phone) }, LogicalOperator.Or);
                    entities = _service.RetrieveMultiple(queryExp);
                    break;
                case 16: //Mobile Phone on new lead matches any phone on contact
                    if (_variables.telephone1 != "")
                    {
                        queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, _variables.telephone1), new CRMQueryExpression("telephone1", ConditionOperator.Equal, _variables.telephone1), new CRMQueryExpression("telephone3", ConditionOperator.Equal, _variables.telephone1) }, LogicalOperator.Or);
                        entities = _service.RetrieveMultiple(queryExp);
                    }
                    break;
                case 17: //Assiagn SAE based on zipcode
                    if (_variables.zipTerritory != Guid.Empty)
                    {
                        Entity territory = new Entity();
                        territory = _service.Retrieve("territory", _variables.zipTerritory, new ColumnSet("managerid"));

                        if (territory.Attributes.Contains("managerid"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("systemuser", new ColumnSet(true), new CRMQueryExpression[] { new CRMQueryExpression("systemuserid", ConditionOperator.Equal, ((EntityReference)territory.Attributes["managerid"]).Id) });

                            entities = _service.RetrieveMultiple(queryExp);
                        }
                    }
                    break;
            }

            return entities;
        }
    }
}
