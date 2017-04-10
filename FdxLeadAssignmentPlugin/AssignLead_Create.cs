﻿using Microsoft.Crm.Sdk.Messages;
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

            //Call Input parameter collection to get all the data passes....
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                //Entity object to update Lead....
                //Entity leadEntity = new Entity()
                //{
                //    LogicalName = "lead",
                //    Id = leadEntity.Id
                //};

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    //Fetch data from lead entity....
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
                    string phone = leadEntity.Attributes["telephone2"].ToString();
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

                    Guid accountid;

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
                    else if ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4))
                    {
                        step = 3;

                        #region 1st check --> first name, last name and phone matches an existing Lead....
                        QueryExpression queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("firstname", ConditionOperator.Equal, firstName), new CRMQueryExpression("lastname", ConditionOperator.Equal, lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                        EntityCollection collection = service.RetrieveMultiple(queryExp);
                        if (collection.Entities.Count > 0)
                        {
                            step = 31;
                            Entity lead = new Entity();
                            lead = collection.Entities[0];
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
                            leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);

                        }
                        #endregion

                        #region 2nd check --> first name, last name and phone match an existing contact....
                        if (step == 3)
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("accountid", "fullname", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("firstname", ConditionOperator.Equal, firstName), new CRMQueryExpression("lastname", ConditionOperator.Equal, lastName), new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 32;
                                Entity contact = new Entity();
                                contact = collection.Entities[0];
                                leadEntity["contactid"] = new EntityReference("contact", contact.Id);
                                leadEntity["parentcontactid"] = new EntityReference("contact", contact.Id);
                                leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);

                                //Check if the account exist for the contact....
                                if (contact.Attributes.Contains("acountid"))
                                {
                                    leadEntity["accountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["account"]).Id);
                                    leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["account"]).Id);
                                    accountid = ((EntityReference)contact.Attributes["account"]).Id;
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
                        
                        #region 3rd check --> email matches an existing lead....
                        if (step == 3 && leadEntity.Attributes.Contains("emailaddress1"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, email) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 33;
                                Entity lead = new Entity();
                                lead = collection.Entities[0];
                                if (lead.Attributes.Contains("contactid"))
                                    leadEntity["contactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["contactid"]).Id);
                                if (lead.Attributes.Contains("parentcontactid"))
                                    leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)lead.Attributes["parentcontactid"]).Id);
                                if (lead.Attributes.Contains("accountid"))
                                    leadEntity["accountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["accountid"]).Id);
                                if (lead.Attributes.Contains("parentaccountid"))
                                {
                                    leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)lead.Attributes["parentaccountid"]).Id);
                                    accountid=((EntityReference)lead.Attributes["parentaccountid"]).Id;
                                }
                                leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);
                            }
                        }
                        #endregion

                        #region 4th check check --> Email matches an existing contact....
                        if (step == 3 && leadEntity.Attributes.Contains("emailaddress1"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("contact", new ColumnSet("accountid", "fullname", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, email) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 34;
                                Entity contact = new Entity();
                                contact = collection.Entities[0];
                                leadEntity["contactid"] = new EntityReference("contact", contact.Id);
                                leadEntity["parentcontactid"] = new EntityReference("contact", contact.Id);
                                leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)contact.Attributes["ownerid"]).Id);

                                //Check if the account exist for the contact....
                                if (contact.Attributes.Contains("acountid"))
                                {
                                    leadEntity["accountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["account"]).Id);
                                    leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)contact.Attributes["account"]).Id);
                                    accountid = ((EntityReference)contact.Attributes["account"]).Id;
                                }
                            }
                        }
                        #endregion

                        #region (Code commented) --> phone in account entity....
                        //if (step == 3)
                        //{
                        //    queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("telephone2", ConditionOperator.Equal, phone) });
                        //    if ((service.RetrieveMultiple(queryExp).Entities.Count) > 0)
                        //    {
                        //        step = 34;
                        //        Entity account = new Entity();
                        //        account = service.RetrieveMultiple(queryExp).Entities[0];
                        //        leadEntity["accountid"] = new EntityReference("account", account.Id);
                        //        leadEntity["parentaccountid"] = new EntityReference("account", account.Id);
                        //        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);

                        //        //Check if the account exist for the contact....
                        //        if (account.Attributes.Contains("primarycontactid"))
                        //        {
                        //            leadEntity["contactd"] = new EntityReference("contact", ((EntityReference)account.Attributes["primarycontactid"]).Id);
                        //            leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)account.Attributes["primarycontactid"]).Id);
                        //        }
                        //    }
                        //}
                        #endregion

                        #region (Code commented) --> email in account entity....
                        //if (step == 3)
                        //{
                        //    queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("emailaddress1", ConditionOperator.Equal, email) });
                        //    if ((service.RetrieveMultiple(queryExp).Entities.Count) > 0)
                        //    {
                        //        step = 35;
                        //        Entity account = new Entity();
                        //        account = service.RetrieveMultiple(queryExp).Entities[0];
                        //        leadEntity["accountid"] = new EntityReference("account", account.Id);
                        //        leadEntity["parentaccountid"] = new EntityReference("account", account.Id);
                        //        leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);

                        //        //Check if the account exist for the contact....
                        //        if (account.Attributes.Contains("primarycontactid"))
                        //        {
                        //            leadEntity["contactd"] = new EntityReference("contact", ((EntityReference)account.Attributes["primarycontactid"]).Id);
                        //            leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)account.Attributes["primarycontactid"]).Id);
                        //        }
                        //    }
                        //}
                        #endregion

                        #region 5th check --> company name and zipcode matches an existing account....
                        if (step == 3 && leadEntity.Attributes.Contains("companyname"))
                        {
                            queryExp = CRMQueryExpression.getQueryExpression("account", new ColumnSet("accountid", "primarycontactid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, companyName), new CRMQueryExpression("address1_postalcode", ConditionOperator.Equal, zipEntity.Attributes["fdx_zipcode"].ToString()) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 35;
                                Entity account = new Entity();
                                account = collection.Entities[0];
                                leadEntity["accountid"] = new EntityReference("account", account.Id);
                                leadEntity["parentaccountid"] = new EntityReference("account", account.Id);
                                leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)account.Attributes["ownerid"]).Id);
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
                            queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid"), new CRMQueryExpression[] { new CRMQueryExpression("companyname", ConditionOperator.Equal, companyName), new CRMQueryExpression("address1_postalcode", ConditionOperator.Equal, zipEntity.Attributes["fdx_zipcode"].ToString()) });
                            collection = service.RetrieveMultiple(queryExp);
                            if (collection.Entities.Count > 0)
                            {
                                step = 36;
                                Entity lead = new Entity();
                                lead = collection.Entities[0];
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
                                leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)lead.Attributes["ownerid"]).Id);

                            }
                        }
                        #endregion

                        #region 7th check --> Assiagn SAE based on zipcode....
                        if (step == 3)
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
                        #endregion                        
                    }

                    //service.Update(leadEntity);

                    #region 8th check --> trigger next@bat....
                    if (step == 3 && (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1))
                    {
                        leadEntity["fdx_snb"] = true;
                        leadAssigned = false;
                    }
                    #endregion

                    #region Set the address field as per the account if there is an existing account....
                    step = 4;
                    const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
                    string url = "http://SMARTCRMSync.1800dentist.com/api/lead/createleadasync?" + apiParm;//Zip={0}";

                    var uri = new Uri(string.Format(url,zipCode));
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
                        Lead leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());
                        step = 7;
                        leadEntity["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                        if(leadObj.goNoGo)
                        {
                            step = 71;
                            leadEntity["fdx_gonogo"] = new OptionSetValue(756480000);
                        }
                        else
                        {
                            step = 72;
                            leadEntity["fdx_gonogo"] = new OptionSetValue(756480001);
                            if (!leadAssigned && ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)))
                            {
                                step = 73;
                                leadEntity["fdx_snb"] = false;
                                QueryExpression teamQuery = CRMQueryExpression.getQueryExpression("team", new ColumnSet("name"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, "Lead Review Team") });
                                EntityCollection teamCollection = service.RetrieveMultiple(teamQuery);
                                step = 74;
                                if (teamCollection.Entities.Count > 0)
                                {
                                    Entity team = teamCollection.Entities[0];
                                    leadEntity["ownerid"] = new EntityReference("team",team.Id);
                                }                               
                            }
                        }
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
    }
}