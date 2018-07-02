using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FdxLeadAssignmentPlugin
{
    class LeadAssignment_Variables
    {
        public const string DEV_ENVIRONMENT_URL = "http://SMARTCRMSync.1800dentist.com/api";
        public const string STAGE_ENVIRONMENT_URL = "http://SMARTCRMSyncStage.1800dentist.com/api";
        public const string PROD_ENVIRONMENT_URL = "http://SMARTCRMSyncProd.1800dentist.com/api";
        public string smartCrmSyncWebServiceUrl;
        public bool leadAssigned;
        public bool acc_gmaccountno_exist;
        public bool isGroupPractice;
        public string zipcodetext;
        public Guid zip;
        public Guid zipTerritory;
        public string firstName;
        public string lastName;
        public string phone;
        public string telephone1;
        public string email;
        public string companyName;
        public string title;
        public string address1;
        public string address2;
        public string city;
        public string state;
        public Guid accountid;
        public DateTime tzTime = DateTime.Now;

        public LeadAssignment_Variables(Entity _leadEntity, IOrganizationService _service)
        {
            smartCrmSyncWebServiceUrl = DEV_ENVIRONMENT_URL;
            leadAssigned = true;
            acc_gmaccountno_exist = false;
            accountid = Guid.Empty;
            zip = Guid.Empty;
            zipTerritory = Guid.Empty;

            isGroupPractice = _leadEntity.Attributes.Contains("fdx_grppracactice") ? _leadEntity.GetAttributeValue<bool>("fdx_grppracactice") : false;

            if (_leadEntity.Attributes.Contains("fdx_zippostalcode"))
            {
                zip = ((EntityReference)_leadEntity.Attributes["fdx_zippostalcode"]).Id;
                Entity zipEntity = new Entity();
                zipEntity = _service.Retrieve("fdx_zipcode", zip, new ColumnSet("fdx_zipcode", "fdx_timezone", "fdx_territory"));
                zipcodetext = zipEntity.Attributes.Contains("fdx_zipcode") ? zipEntity.Attributes["fdx_zipcode"].ToString() : "";
                int timeZoneCode = Convert.ToInt32(zipEntity["fdx_timezone"]);
                QueryExpression tzDefinationQuery = CRMQueryExpression.getQueryExpression("timezonedefinition", new ColumnSet("standardname"), new CRMQueryExpression[] { new CRMQueryExpression("timezonecode", ConditionOperator.Equal, timeZoneCode) });

                Entity tzDefination = _service.RetrieveMultiple(tzDefinationQuery).Entities[0];
                DateTime timeUtc = DateTime.UtcNow;
                TimeZoneInfo tzInfo = TimeZoneInfo.FindSystemTimeZoneById(tzDefination.Attributes["standardname"].ToString());
                tzTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, tzInfo);

                if(zipEntity.Contains("fdx_territory"))
                {
                    zipTerritory = ((EntityReference)zipEntity.Attributes["fdx_territory"]).Id;
                }
            }
            
            firstName = _leadEntity.Attributes.Contains("firstname") ? _leadEntity.Attributes["firstname"].ToString() : "";

            lastName = _leadEntity.Attributes.Contains("lastname") ? _leadEntity.Attributes["lastname"].ToString() : "";

            phone = _leadEntity.Attributes.Contains("telephone2") ? Regex.Replace(_leadEntity.Attributes["telephone2"].ToString(), @"[^0-9]+", "") : "";

            telephone1 = _leadEntity.Attributes.Contains("telephone1") ? Regex.Replace(_leadEntity.Attributes["telephone1"].ToString(), @"[^0-9]+", "") : "";

            email = _leadEntity.Attributes.Contains("emailaddress1") ? _leadEntity.Attributes["emailaddress1"].ToString() : "";

            companyName = _leadEntity.Attributes.Contains("companyname") ? _leadEntity.Attributes["companyname"].ToString() : "";

            title = _leadEntity.Attributes.Contains("fdx_jobtitlerole") ? CRMQueryExpression.GetOptionsSetTextForValue(_service, "lead", "fdx_jobtitlerole", ((OptionSetValue)_leadEntity.Attributes["fdx_jobtitlerole"]).Value) : "";

            address1 = _leadEntity.Attributes.Contains("address1_line1") ? _leadEntity.Attributes["address1_line1"].ToString() : "";

            address2 = _leadEntity.Attributes.Contains("address1_line2") ? _leadEntity.Attributes["address1_line2"].ToString() : "";

            city = _leadEntity.Attributes.Contains("address1_city") ? _leadEntity.Attributes["address1_city"].ToString() : "";

            state = _leadEntity.Attributes.Contains("fdx_stateprovince") ? (_service.Retrieve("fdx_state", ((EntityReference)_leadEntity.Attributes["fdx_stateprovince"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString() : "";

            if ((((OptionSetValue)_leadEntity["leadsourcecode"]).Value == 2) || (((OptionSetValue)_leadEntity["leadsourcecode"]).Value == 3) || (((OptionSetValue)_leadEntity["leadsourcecode"]).Value == 5))
            {
                if (_leadEntity.Attributes.Contains("parentaccountid"))
                {
                    accountid = ((EntityReference)_leadEntity.Attributes["parentaccountid"]).Id;
                }
            }
        }
    }
}
