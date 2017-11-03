using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.Serialization.Json;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace FdxLeadAssignmentPlugin
{
    public class AssignLead_Create : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            int step = 0;
            bool leadAssigned = true;
            bool acc_gmaccountno_exist = false;

            //Call Input parameter collection to get all the data passes....
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    if (leadEntity.Attributes.Contains("telephone1"))
                        leadEntity.Attributes["telephone1"] = Regex.Replace(leadEntity.Attributes["telephone1"].ToString(), @"[^0-9]+", "");

                    if (leadEntity.Attributes.Contains("telephone2"))
                        leadEntity.Attributes["telephone2"] = Regex.Replace(leadEntity.Attributes["telephone2"].ToString(), @"[^0-9]+", "");

                    if (leadEntity.Attributes.Contains("telephone3"))
                        leadEntity.Attributes["telephone3"] = Regex.Replace(leadEntity.Attributes["telephone3"].ToString(), @"[^0-9]+", "");
                     
                    //Fetch data from lead entity....
                    string url = "";
                    step = 99;
                    bool isGroupPractice = false ;                    
                    if(leadEntity.Attributes.Contains("fdx_grppracactice"))
                        isGroupPractice = leadEntity.GetAttributeValue<bool>("fdx_grppracactice");
                    step = 98;
                    Guid zip = ((EntityReference)leadEntity.Attributes["fdx_zippostalcode"]).Id;
                    string zipcodetext = (service.Retrieve("fdx_zipcode", zip, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString();
                    step = 97;
                    string firstName = leadEntity.Attributes["firstname"].ToString();
                    step = 96;
                    string lastName = leadEntity.Attributes["lastname"].ToString();
                    step = 95;
                    string phone = Regex.Replace(leadEntity.Attributes["telephone2"].ToString(),@"[^0-9]+", "");
                    //string phone = leadEntity.Attributes["telephone2"].ToString();
                    step = 94;
                    string apiParm = string.Format("Zip={0}&Contact={1} {2}&Phone1={3}", zipcodetext, firstName, lastName, phone);
                    string email = "";
                    if(leadEntity.Attributes.Contains("emailaddress1"))
                        email = leadEntity.Attributes["emailaddress1"].ToString();
                    step = 93;
                    string companyName = "";
                    if (leadEntity.Attributes.Contains("companyname"))
                    {
                        companyName = leadEntity.Attributes["companyname"].ToString();
                        apiParm += string.Format("&Company={0}", companyName);
                    }
                    string title = "";
                    if(leadEntity.Attributes.Contains("fdx_jobtitlerole"))
                    {
                        title = CRMQueryExpression.GetOptionsSetTextForValue(service, "lead", "fdx_jobtitlerole", ((OptionSetValue)leadEntity.Attributes["fdx_jobtitlerole"]).Value);
                        apiParm += string.Format("&Title={0}", title);
                    }
                    string address1 = "";
                    if(leadEntity.Attributes.Contains("address1_line1"))
                    {
                        address1 = leadEntity.Attributes["address1_line1"].ToString();
                        apiParm += string.Format("&Address1={0}", address1);
                    }
                    string address2 = "";
                    if (leadEntity.Attributes.Contains("address1_line2"))
                    {
                        address2 = leadEntity.Attributes["address1_line2"].ToString();
                        apiParm += string.Format("&Address2={0}", address2);
                    }
                    string city = "";
                    if(leadEntity.Attributes.Contains("address1_city"))
                    {
                        city = leadEntity.Attributes["address1_city"].ToString();
                        apiParm += string.Format("&City={0}", city);
                    }
                    string state = "";
                    if(leadEntity.Attributes.Contains("fdx_stateprovince"))
                    {
                        state = (service.Retrieve("fdx_state", ((EntityReference)leadEntity.Attributes["fdx_stateprovince"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString();
                        apiParm += string.Format("&State={0}", state);
                    }

                    Guid accountid = Guid.Empty;

                    /*Raghava Chandra - Update 1 on 26-Apr-2017*/

                    //To Check if Existing Account is specified while creating a new Lead. This field will have value from context only if Lead is of type Email,PhoneCall or other generally
                    if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 2) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 3) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 5))
                    {
                        if (leadEntity.Attributes.Contains("parentaccountid"))
                        {
                            accountid = ((EntityReference)leadEntity.Attributes["parentaccountid"]).Id;
                        }
                    }

                    /*End of Update 1*/

                    //Set created on time based on Leads time zone....
                    step = 1;
                    Entity zipEntity = service.Retrieve("fdx_zipcode", zip, new ColumnSet("fdx_zipcode", "fdx_timezone"));
                    string zipCode = zipEntity.Attributes["fdx_zipcode"].ToString();
                    if (zipEntity.Attributes.Contains("fdx_timezone"))
                    {
                        int timeZoneCode = Convert.ToInt32(zipEntity["fdx_timezone"]);
                        QueryExpression tzDefinationQuery = CRMQueryExpression.getQueryExpression("timezonedefinition", new ColumnSet("standardname"), new CRMQueryExpression[] { new CRMQueryExpression("timezonecode", ConditionOperator.Equal, timeZoneCode) });

                        Entity tzDefination = service.RetrieveMultiple(tzDefinationQuery).Entities[0];
                        DateTime timeUtc = DateTime.UtcNow;
                        TimeZoneInfo tzInfo = TimeZoneInfo.FindSystemTimeZoneById(tzDefination.Attributes["standardname"].ToString());
                        DateTime tzTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, tzInfo);
                        leadEntity["fdx_createdontime"] = tzTime.Hour;
                    }

                    if(isGroupPractice)
                    {
                        step = 2;
                        #region Code commented to Hold DSO logic...
                        //Guid state = new Guid();
                        //Entity leadState = service.Retrieve("lead", leadEntity.Id, new ColumnSet("fdx_stateprovince"));
                        //if (leadState.Attributes.Contains("fdx_stateprovince"))
                        //    state = ((EntityReference)leadState.Attributes["fdx_stateprovince"]).Id;
                        //#region 1st check --> Get DSO rep based on territory assigned to State....                        
                        //Entity stateEntity = service.Retrieve("fdx_state", state, new ColumnSet("fdx_territory"));

                        //if(stateEntity.Attributes.Contains("fdx_territory"))
                        //{
                        //    step = 21;
                        //    Entity territory = service.Retrieve("territory", ((EntityReference)stateEntity.Attributes["fdx_territory"]).Id, new ColumnSet("managerid"));
                        //    if(territory.Attributes.Contains("managerid"))
                        //        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)territory["managerid"]).Id);
                        //}
                        //#endregion
                        #endregion
                    }
                    //else if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)) -- Code commented by Ram as part of SMART593
                    else
                    {
                        step = 3;

                        #region 1st check --> first name, last name and phone matches an existing Lead....
                        QueryExpression queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid","owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("firstname", ConditionOperator.Equal, firstName), new CRMQueryExpression("lastname", ConditionOperator.Equal, lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                        EntityCollection collection = service.RetrieveMultiple(queryExp);
                        if (collection.Entities.Count > 0)
                        {
                            step = 31;
                            Entity lead = new Entity();
                            lead = collection.Entities[0];
                            if (!leadEntity.Attributes.Contains("fdx_leadid"))
                                leadEntity["fdx_leadid"] = new EntityReference("lead", lead.Id);
                            if (lead.Attributes.Contains("contactid"))
                                leadEntity["contactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["contactid"]).Id);
                            if (lead.Attributes.Contains("parentcontactid"))
                                leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["parentcontactid"]).Id);
                            if (lead.Attributes.Contains("accountid"))
                                leadEntity["accountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["accountid"]).Id);
                            if (lead.Attributes.Contains("parentaccountid"))
                            {
                                leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["parentaccountid"]).Id);
                                accountid = ((EntityReference)lead.Attributes["parentaccountid"]).Id;
                            }

                            //Condition added by Ram as Part of SMART593....
                            if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                            {
                                if (lead.Attributes.Contains("owningteam"))
                                    leadEntity["ownerid"] = new EntityReference(lead["owningteam"] != null ? "team" : "systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                else
                                    leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                            }

                        }
                        #endregion
                                                   
                        #region 2nd check --> first name, last name and phone match an existing contact....
                        if (step == 3)
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("firstname", ConditionOperator.Equal, firstName), new CRMQueryExpression("lastname", ConditionOperator.Equal, lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 32;
                                Entity contact = new Entity();
                                contact = collection.Entities[0];
                                leadEntity["contactid"] = new EntityReference("contact", contact.Id);
                                leadEntity["parentcontactid"] = new EntityReference("contact", contact.Id);

                                //Condition added by Ram as Part of SMART593....
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (contact.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(contact["owningteam"] != null ? "team" : "systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);
                                }

                                //Check if the account exist for the contact....
                                if (contact.Attributes.Contains("parentcustomerid"))
                                {
                                    if (((EntityReference)contact.Attributes["parentcustomerid"]).LogicalName == "account")
                                    {
                                        leadEntity["accountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["parentcustomerid"]).Id);
                                        leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["parentcustomerid"]).Id);
                                        accountid = ((EntityReference)contact.Attributes["parentcustomerid"]).Id;
                                    }
                                }
                            }
                        }
                        #endregion

                        #region 8th check --> Office phone number maches an existing lead....
                        if (step == 3)
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 38;
                                Entity lead = new Entity();
                                lead = collection.Entities[0];
                                if (!leadEntity.Attributes.Contains("fdx_leadid"))
                                    leadEntity["fdx_leadid"] = new EntityReference("lead", lead.Id);
                                if (lead.Attributes.Contains("accountid"))
                                    leadEntity["accountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["accountid"]).Id);
                                if (lead.Attributes.Contains("parentaccountid"))
                                {
                                    leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["parentaccountid"]).Id);
                                    accountid = ((EntityReference)lead.Attributes["parentaccountid"]).Id;
                                }
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (lead.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(lead["owningteam"] != null ? "team" : "systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                }

                            }
                        }
                        #endregion

                        #region (Code commented) --> phone number in contact entity....
                        //if (step == 3)
                        //{
                        //    queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("accountid", "fullname", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                        //    if ((service.RetrieveMultiple(queryExp).Entities.Count) > 0)
                        //    {
                        //        step = 32;
                        //        Entity contact = new Entity();
                        //        contact = service.RetrieveMultiple(queryExp).Entities[0];
                        //        leadEntity["contactid"] = new EntityReference("contact", contact.Id);
                        //        leadEntity["parentcontactid"] = new EntityReference("contact", contact.Id);
                        //        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);

                        //        //Check if the account exist for the contact....
                        //        if (contact.Attributes.Contains("acountid"))
                        //        {                                    
                        //            leadEntity["accountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["account"]).Id);
                        //            leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["account"]).Id);
                        //        }
                        //    }
                        //}
                        #endregion

                        #region 9th check --> office phone matches any phone number in an existing account....
                        if (step == 3)
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone), new CRMQueryExpression("telephone1", ConditionOperator.Equal, phone), new CRMQueryExpression("telephone3", ConditionOperator.Equal, phone) }, LogicalOperator.Or);
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 39;
                                Entity account = new Entity();
                                account = collection.Entities[0];

                                Entity matchedLead = new Entity();
                                matchedLead = getAccountMatchedLead(account, service);
                                if (!(matchedLead.Id == Guid.Empty) && !(leadEntity.Attributes.Contains("fdx_leadid")))
                                    leadEntity["fdx_leadid"] = new EntityReference("lead", matchedLead.Id);

                                leadEntity["accountid"] = new EntityReference("account", account.Id);
                                leadEntity["parentaccountid"] = new EntityReference("account", account.Id);
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (account.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(account["owningteam"] != null ? "team" : "systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
                                }

                                accountid = account.Id;

                            }
                        }
                        #endregion
                        
                        #region 3rd check --> email matches an existing lead....
                        if (step == 3 && leadEntity.Attributes.Contains("emailaddress1"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, email) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 33;
                                Entity lead = new Entity();
                                lead = collection.Entities[0];
                                if (!leadEntity.Attributes.Contains("fdx_leadid"))
                                    leadEntity["fdx_leadid"] = new EntityReference("lead", lead.Id);
                                //Start :: Commented as part of SMART-627....
                                //if (lead.Attributes.Contains("contactid"))
                                //    leadEntity["contactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["contactid"]).Id);
                                //if (lead.Attributes.Contains("parentcontactid"))
                                //    leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["parentcontactid"]).Id);
                                //End :: Commented as part of SMART-627....
                                if (lead.Attributes.Contains("accountid"))
                                    leadEntity["accountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["accountid"]).Id);
                                if (lead.Attributes.Contains("parentaccountid"))
                                {
                                    leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["parentaccountid"]).Id);
                                    accountid=((EntityReference)lead.Attributes["parentaccountid"]).Id;
                                }

                                //Condition added by Ram as Part of SMART593....
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (lead.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(lead["owningteam"] != null ? "team" : "systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                }
                            }
                        }
                        #endregion

                        #region 10th check --> email matches an existing account....
                        if (step == 3)
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, email) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 310;
                                Entity account = new Entity();
                                account = collection.Entities[0];

                                Entity matchedLead = new Entity();
                                matchedLead = getAccountMatchedLead(account, service);
                                if (!(matchedLead.Id == Guid.Empty) && !(leadEntity.Attributes.Contains("fdx_leadid")))
                                    leadEntity["fdx_leadid"] = new EntityReference("lead", matchedLead.Id);

                                leadEntity["accountid"] = new EntityReference("account", account.Id);
                                leadEntity["parentaccountid"] = new EntityReference("account", account.Id);
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (account.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(account["owningteam"] != null ? "team" : "systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
                                }

                                accountid = account.Id;
                            }
                        }
                        #endregion

                        #region 4th check --> Email matches an existing contact....
                        if (step == 3 && leadEntity.Attributes.Contains("emailaddress1"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("parentcustomerid", "fullname", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, email) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 34;
                                Entity contact = new Entity();
                                contact = collection.Entities[0];
                                //Start :: Commented as part of SMART-627....
                                //leadEntity["contactid"] = new EntityReference("contact", contact.Id);
                                //leadEntity["parentcontactid"] = new EntityReference("contact", contact.Id);
                                //Start :: Commented as part of SMART-627....

                                //Condition added by Ram as Part of SMART593....
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (contact.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(contact["owningteam"] != null ? "team" : "systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);
                                }

                                //Check if the account exist for the contact....
                                if (contact.Attributes.Contains("parentcustomerid"))
                                {
                                    if (((EntityReference)contact.Attributes["parentcustomerid"]).LogicalName == "account")
                                    {
                                        leadEntity["accountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["parentcustomerid"]).Id);
                                        leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["parentcustomerid"]).Id);
                                        accountid = ((EntityReference)contact.Attributes["parentcustomerid"]).Id;
                                    }
                                }
                            }
                        }
                        #endregion                                                                      

                        #region 5th check --> company name and zipcode matches an existing account....
                        if (step == 3 && leadEntity.Attributes.Contains("companyname"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, companyName), new CRMQueryExpression("address1_postalcode", ConditionOperator.Equal, zipEntity.Attributes["fdx_zipcode"].ToString()) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 35;
                                Entity account = new Entity();
                                account = collection.Entities[0];

                                Entity matchedLead = new Entity();
                                matchedLead = getAccountMatchedLead(account, service);
                                if (!(matchedLead.Id == Guid.Empty) && !(leadEntity.Attributes.Contains("fdx_leadid")))
                                    leadEntity["fdx_leadid"] = new EntityReference("lead", matchedLead.Id);

                                leadEntity["accountid"] = new EntityReference("account", account.Id);
                                leadEntity["parentaccountid"] = new EntityReference("account", account.Id);

                                //Condition added by Ram as Part of SMART593....
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (account.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(account["owningteam"] != null ? "team" : "systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
                                }

                                accountid = account.Id;

                                //Check if the account exist for the contact....
                                //if (account.Attributes.Contains("primarycontactid"))
                                //{
                                //    leadEntity["contactd"] = new EntityReference("contact", ((EntityReference)account.Attributes["primarycontactid"]).Id);
                                //    leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)account.Attributes["primarycontactid"]).Id);
                                //}
                            }
                        }
                        #endregion

                        #region 6th check --> company name and zipcode matches an existing lead....
                        if (step == 3 && leadEntity.Attributes.Contains("companyname"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid","owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("companyname", ConditionOperator.Equal, companyName), new CRMQueryExpression("address1_postalcode", ConditionOperator.Equal, zipEntity.Attributes["fdx_zipcode"].ToString()) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 36;
                                Entity lead = new Entity();
                                lead = collection.Entities[0];
                                if (!leadEntity.Attributes.Contains("fdx_leadid"))
                                    leadEntity["fdx_leadid"] = new EntityReference("lead", lead.Id);
                                //if (lead.Attributes.Contains("contactid"))
                                //    leadEntity["contactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["contactid"]).Id);
                                //if (lead.Attributes.Contains("parentcontactid"))
                                //    leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["parentcontactid"]).Id);
                                if (lead.Attributes.Contains("accountid"))
                                    leadEntity["accountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["accountid"]).Id);
                                if (lead.Attributes.Contains("parentaccountid"))
                                {
                                    leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["parentaccountid"]).Id);
                                    accountid = ((EntityReference)lead.Attributes["parentaccountid"]).Id;
                                }

                                //Condition added by Ram as Part of SMART593....
                                if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                                {
                                    if (lead.Attributes.Contains("owningteam"))
                                        leadEntity["ownerid"] = new EntityReference(lead["owningteam"] != null ? "team" : "systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                    else
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                                }

                            }
                        }
                        #endregion                        

                        #region 7th check --> Assiagn SAE based on zipcode....
                        if (step == 3)
                        {
                            //Condition added by Ram as Part of SMART593....
                            if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                            {
                                Entity zipcode = service.Retrieve("fdx_zipcode", zip, new ColumnSet("fdx_territory", "fdx_zipcode"));
                                if (zipcode.Attributes.Contains("fdx_territory"))
                                {
                                    step = 37;
                                    Entity zipTerritory = service.Retrieve("fdx_zipcode", zipcode.Id, new ColumnSet("fdx_territory"));
                                    Entity territory = new Entity();
                                    if (zipTerritory.Attributes.Contains("fdx_territory"))
                                        territory = service.Retrieve("territory", ((EntityReference)zipcode.Attributes["fdx_territory"]).Id, new ColumnSet("managerid"));
                                    if (territory.Attributes.Contains("managerid"))
                                        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)territory.Attributes["managerid"]).Id);
                                }
                            }
                        }
                        #endregion                        
                    }

                    #region (Code Commented)8th check --> trigger next@bat....
                    //if (step == 3 && (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1))
                    //{
                    //    leadEntity["fdx_snb"] = true;
                    //    leadAssigned = false;
                    //}
                    #endregion

                    #region 8th check --> Set lead Assigned to False and trigger next@bat only for Web Leads (new leads only and not cloned)....
                    if (step == 3 && (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1 || ((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                    {
                        leadAssigned = false;
                        if (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1 && !leadEntity.Attributes.Contains("fdx_leadid"))
                            leadEntity["fdx_snb"] = true;                        
                    }
                    #endregion

                    #region (Code Commented)Set the address field as per the account if there is an existing account....
                    //step = 4;
                    //const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
                    //string url = "http://SMARTCRMSync.1800dentist.com/api/lead/createleadasync?" + apiParm;//Zip={0}";

                    //var uri = new Uri(string.Format(url,zipCode));
                    //var request = WebRequest.Create(uri);
                    //request.Method = WebRequestMethods.Http.Post;
                    //request.ContentType = "application/json";
                    //request.ContentLength = 0;
                    //request.Headers.Add("Authorization", token);
                    //step = 5;
                    //using (var getResponse = request.GetResponse())
                    //{
                    //    DataContractJsonSerializer serializer =
                    //                new DataContractJsonSerializer(typeof(Lead));

                    //    step = 6;
                    //    Lead leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());
                    //    step = 7;
                    //    leadEntity["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                    //    if(leadObj.goNoGo)
                    //    {
                    //        step = 71;
                    //        leadEntity["fdx_gonogo"] = new OptionSetValue(756480000);
                    //    }
                    //    else
                    //    {
                    //        step = 72;
                    //        leadEntity["fdx_gonogo"] = new OptionSetValue(756480001);
                    //        if (!leadAssigned && ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)))
                    //        {
                    //            step = 73;
                    //            leadEntity["fdx_snb"] = false;
                    //            QueryExpression teamQuery = CRMQueryExpression.getQueryExpression("team", new ColumnSet("name"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, "Lead Review Team") });
                    //            EntityCollection teamCollection = service.RetrieveMultiple(teamQuery);
                    //            step = 74;
                    //            if (teamCollection.Entities.Count > 0)
                    //            {
                    //                Entity team = teamCollection.Entities[0];
                    //                leadEntity["ownerid"] = new EntityReference("team",team.Id);
                    //            }                               
                    //        }
                    //    }
                    //}
                    #endregion

                    #region Set the address field as per the account if there is an existing account....
                    if(accountid != Guid.Empty)
                    {
                        QueryExpression accountQuery = CRMQueryExpression.getQueryExpression("account", new ColumnSet("name","fdx_goldmineaccountnumber", "fdx_gonogo","address1_line1","address1_line2","address1_city","fdx_stateprovinceid","fdx_zippostalcodeid","telephone1"), new CRMQueryExpression[] { new CRMQueryExpression("accountid", ConditionOperator.Equal, accountid) });
                        EntityCollection accountCollection = service.RetrieveMultiple(accountQuery);
                        if ((accountCollection.Entities.Count) > 0)
                        {
                            step = 135;
                            Entity account = new Entity();
                            account = accountCollection.Entities[0];
                            if (!account.Attributes.Contains("fdx_goldmineaccountnumber"))
                            {
                                step = 194;
                                apiParm = string.Format("Zip={0}&Phone1={1}", (service.Retrieve("fdx_zipcode", ((EntityReference)account.Attributes["fdx_zippostalcodeid"]).Id, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString(), Regex.Replace(account.Attributes["telephone1"].ToString(),@"[^0-9]+", ""));

                                //apiParm = string.Format("Zip={0}&Phone1={1}", (service.Retrieve("fdx_zipcode", ((EntityReference)account.Attributes["fdx_zippostalcodeid"]).Id, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString(), account.Attributes["telephone1"].ToString());
                                step = 195;
                                if (account.Attributes.Contains("name"))
                                    apiParm += string.Format("&Company={0}", account.Attributes["name"].ToString());

                                step = 196;
                                if (account.Attributes.Contains("address1_line1"))
                                    apiParm += string.Format("&Address1={0}", account.Attributes["address1_line1"].ToString());

                                step = 197;
                                if (account.Attributes.Contains("address1_line2"))
                                    apiParm += string.Format("&Address2={0}", account.Attributes["address1_line2"].ToString());

                                step = 198;
                                if (account.Attributes.Contains("address1_city"))
                                    apiParm += string.Format("&City={0}", account.Attributes["address1_city"].ToString());

                                step = 199;
                                if (account.Attributes.Contains("fdx_stateprovinceid"))
                                    apiParm += string.Format("&State={0}", (service.Retrieve("fdx_state", ((EntityReference)account.Attributes["fdx_stateprovinceid"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString());

                                //1. To point to Dev
                                //url = "http://SMARTCRMSync.1800dentist.com/api/lead/createlead?" + apiParm;
                                
                                //2. To point to Stage
                                //url = "http://smartcrmsyncstage.1800dentist.com/api/lead/createlead?" + apiParm;
                                
                                //3. To point to Production
                                url = "http://SMARTCRMSyncProd.1800dentist.com/api/lead/createlead?" + apiParm;
                                
                            }
                            else
                            {
                                acc_gmaccountno_exist = true;
                                leadEntity["fdx_goldmineaccountnumber"] = account.Attributes["fdx_goldmineaccountnumber"].ToString();
                                leadEntity["fdx_gonogo"] = account.Attributes["fdx_gonogo"];
                            }
                        }
                    }
                    else
                    {
                        //1. To point to Dev
                        //url = "http://SMARTCRMSync.1800dentist.com/api/lead/createlead?" + apiParm;

                        //2. To point to Stage
                        //url = "http://smartcrmsyncstage.1800dentist.com/api/lead/createlead?" + apiParm;

                        //3. To point to Production
                        url = "http://SMARTCRMSyncProd.1800dentist.com/api/lead/createlead?" + apiParm;
                                
                    }
                    #endregion

                    #region Call and update from API....
                    Lead leadObj = new Lead();
                    if (!acc_gmaccountno_exist)
                    {
                        step = 4;
                        const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
                        //This zipCode needs to be changed to that of Account
                        var uri = new Uri(url);
                        var request = WebRequest.Create(uri);
                        request.Method = WebRequestMethods.Http.Post;
                        request.ContentType = "application/json";
                        request.ContentLength = 0;
                        request.Headers.Add("Authorization", token);
                        step = 5;
                        using (var getResponse = request.GetResponse())
                        {
                            DataContractJsonSerializer serializer =
                                        new DataContractJsonSerializer(typeof(Lead));

                            step = 6;
                            leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());
                            step = 7;
                            leadEntity["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                            if (leadObj.goNoGo)
                            {
                                step = 71;
                                leadEntity["fdx_gonogo"] = new OptionSetValue(756480000);
                            }
                            else
                            {
                                step = 72;
                                leadEntity["fdx_gonogo"] = new OptionSetValue(756480001);                                
                            }

                            if (accountid != Guid.Empty)
                            {
                                step = 73;
                                Entity acc = new Entity("account")
                                {
                                    Id = accountid
                                };
                                acc.Attributes["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                                acc.Attributes["fdx_gonogo"] = leadObj.goNoGo ? new OptionSetValue(756480000) : new OptionSetValue(756480001);
                                service.Update(acc);
                            }
                        }
                    }
                    #endregion

                    #region Condition to assign Lead to Lead Review Team....
                    if (!leadAssigned && !leadObj.goNoGo && ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)))
                    {
                        step = 74;
                        leadEntity["fdx_snb"] = false;
                        QueryExpression teamQuery = CRMQueryExpression.getQueryExpression("team", new ColumnSet("name"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, "Lead Review Team") });
                        EntityCollection teamCollection = service.RetrieveMultiple(teamQuery);
                        step = 75;
                        if (teamCollection.Entities.Count > 0)
                        {
                            Entity team = teamCollection.Entities[0];
                            leadEntity["ownerid"] = new EntityReference("team", team.Id);
                        }
                    }
                    #endregion

                    #region Condition if Cloned Lead: override owner of cloned lead = originating lead's owner

                    if (leadEntity.Attributes.Contains("fdx_leadid"))
                    {
                        Entity OriginatingLead = service.Retrieve("lead", ((EntityReference)leadEntity.Attributes["fdx_leadid"]).Id, new ColumnSet("ownerid", "owningteam"));

                        if (OriginatingLead.Attributes.Contains("owningteam"))
                            leadEntity["ownerid"] = new EntityReference(OriginatingLead["owningteam"] != null ? "team" : "systemuser", ((EntityReference)OriginatingLead.Attributes["ownerid"]).Id);
                        else
                            leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)OriginatingLead.Attributes["ownerid"]).Id);
                    }
                    #endregion

                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the AssignLead_Create plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("AssignLead_Create: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }

        private Entity getAccountMatchedLead(Entity _account, IOrganizationService _service)
        {
            Entity lead = new Entity ();
            QueryExpression queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("accountid", ConditionOperator.Equal, _account.Id), new CRMQueryExpression("parentaccountid", ConditionOperator.Equal, _account.Id) }, LogicalOperator.Or);
            EntityCollection collection = _service.RetrieveMultiple(queryExp);

            if (collection.Entities.Count > 0)
                lead = collection.Entities[0];

            return lead;
        }
    }
}