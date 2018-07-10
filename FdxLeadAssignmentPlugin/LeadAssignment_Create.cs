using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
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
    public class LeadAssignment_Create : IPlugin
    {
        LeadAssignment_Variables variables;
        IOrganizationService service;
        ITracingService tracingService;
        Entity leadEntity;
        public void Execute(IServiceProvider serviceProvider)
        {
            int step = 0;
            string url = "";
            string apiParm = "";

            //Extract the tracing service for use in debugging sandboxed plug-ins....
            step = 87;
            tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            step = 88;
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                step = 89;
                leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                try
                {
                    step = 90;
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Format phone numbers to remove special characters....
                    step = 91;
                    this.formatPhoneNumbers();

                    //Initialize lead variables....
                    step = 92;
                    variables = new LeadAssignment_Variables(leadEntity, service);

                    leadEntity["fdx_createdontime"] = variables.tzTime.Hour;

                    //Set api parameters....
                    step = 93;
                    apiParm = this.setApiParmFromLeadEntity();

                    //Set Connected lead, existing contact, existinng account and Lead owner based on criteria....
                    if(!variables.isGroupPractice)
                    {
                        EntityCollection entityCollection = new EntityCollection ();
                        if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 1)).Entities.Count > 0)
                        {
                            tracingService.Trace("First name, last name and office phone on new lead matches First name, last name and office phone on existing lead");
                            step = 1;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 2;
                            this.setExistingContact(entityCollection.Entities[0], "lead");
                            step = 3;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 4;
                            this.setLeadOwner(entityCollection.Entities[0]);
                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 2)).Entities.Count > 0)
                        {
                            tracingService.Trace("First name, last name and office phone on new lead matches First name, last name and office phone on existing contact");
                            step = 5;
                            this.setExistingContact(entityCollection.Entities[0], "contact");
                            step = 6;
                            this.setExistingAccount(entityCollection.Entities[0], "contact");
                            step = 7;
                            this.setLeadOwner(entityCollection.Entities[0]);
                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 3)).Entities.Count > 0)
                        {
                            tracingService.Trace("Office Phone on new lead matches office phone on existing lead");
                            step = 8;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 9;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 10;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 4)).Entities.Count > 0)
                        {
                            tracingService.Trace("Email on new lead matches email on existing lead");
                            step = 11;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 12;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 13;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 5)).Entities.Count > 0)
                        {
                            tracingService.Trace("Company name and zipcode on new lead matches company name and zipcode on an account");
                            step = 14;
                            Entity matchedLead = new Entity();
                            matchedLead = this.getAccountMatchedLead(entityCollection.Entities[0]);
                            step = 15;
                            if (matchedLead.Id != Guid.Empty)
                            {
                                step = 16;
                                this.setConnectedLead(matchedLead);
                            }

                            step = 17;
                            this.setExistingAccount(entityCollection.Entities[0], "account");
                            step = 18;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 6)).Entities.Count > 0)
                        {
                            tracingService.Trace("Company name and zipcode on new lead matches company name and zipcode on a lead");
                            step = 19;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 20;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 21;
                            this.setLeadOwner(entityCollection.Entities[0]);
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 7)).Entities.Count > 0)
                        {
                            tracingService.Trace("Email on new lead matches email on account");
                            Entity matchedLead = new Entity();
                            step = 22;
                            matchedLead = this.getAccountMatchedLead(entityCollection.Entities[0]);
                            step = 23;
                            if (matchedLead.Id != Guid.Empty)
                            {
                                step = 24;
                                this.setConnectedLead(matchedLead);
                            }
                            this.setExistingAccount(entityCollection.Entities[0], "account");
                            step = 25;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 8)).Entities.Count > 0)
                        {
                            tracingService.Trace("Email on new lead matches email on existing contact");
                            step = 26;
                            this.setExistingAccount(entityCollection.Entities[0], "contact");
                            step = 27;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 9)).Entities.Count > 0)
                        {
                            tracingService.Trace("Lastname and phone on new lead matches lastname and any phone on contact");
                            step = 28;
                            this.setExistingAccount(entityCollection.Entities[0], "contact");
                            step = 29;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 10)).Entities.Count > 0)
                        {
                            tracingService.Trace("Lastname and phone on new lead matches lastname and any phone on lead");
                            step = 30;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 31;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 32;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 11)).Entities.Count > 0)
                        {
                            tracingService.Trace("Office Phone on new lead matches any phone on a lead");
                            step = 33;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 34;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 35;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 12)).Entities.Count > 0)
                        {
                            tracingService.Trace("Mobile Phone on new lead matches any phone on a lead");
                            step = 36;
                            this.setConnectedLead(entityCollection.Entities[0]);
                            step = 37;
                            this.setExistingAccount(entityCollection.Entities[0], "lead");
                            step = 38;
                            this.setLeadOwner(entityCollection.Entities[0]);                            
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 13)).Entities.Count > 0)
                        {
                            tracingService.Trace("Office Phone on new lead matches any phone on account");
                            Entity matchedLead = new Entity();
                            step = 39;
                            matchedLead = this.getAccountMatchedLead(entityCollection.Entities[0]);
                            step = 40;
                            if (matchedLead.Id != Guid.Empty)
                            {
                                step = 41;
                                this.setConnectedLead(matchedLead);
                            }
                            step = 42;
                            this.setExistingAccount(entityCollection.Entities[0], "account");
                            step = 43;
                            this.setLeadOwner(entityCollection.Entities[0]);                            

                            //Code added as part of S893 by Meghana on Lead create for Incoming Phone call,web creation/import
                            //wf name - fdx_Lead update Existing Account
                            step = 44;
                            //this.updateExistingAccount(matchedLead.Id);
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 14)).Entities.Count > 0)
                        {
                            tracingService.Trace("Mobile Phone on new lead matches any phone on account");
                            Entity matchedLead = new Entity();
                            step = 45;
                            matchedLead = this.getAccountMatchedLead(entityCollection.Entities[0]);
                            step = 46;
                            if (matchedLead.Id != Guid.Empty)
                            {
                                step = 47;
                                this.setConnectedLead(matchedLead);
                            }
                            step = 48;
                            this.setExistingAccount(entityCollection.Entities[0], "account");
                            step = 49;
                            this.setLeadOwner(entityCollection.Entities[0]);
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 15)).Entities.Count > 0)
                        {
                            tracingService.Trace("Office Phone on new lead matches any phone on contact");
                            step = 50;
                            this.setExistingAccount(entityCollection.Entities[0], "contact");
                            step = 51;
                            this.setLeadOwner(entityCollection.Entities[0]);
                        }
                        else if ((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 16)).Entities.Count > 0)
                        {
                            tracingService.Trace("Mobile Phone on new lead matches any phone on contact");
                            step = 52;
                            this.setExistingAccount(entityCollection.Entities[0], "contact");
                            step = 53;
                            this.setLeadOwner(entityCollection.Entities[0]);
                        }
                        else if (((entityCollection = CRMQueryExpression.GetLeadEntityCollection(variables, service, 17)).Entities.Count > 0) && ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)))
                        {
                            tracingService.Trace("Assiagn SAE based on zipcode");
                            step = 54;
                            Entity user = entityCollection.Entities[0];
                            step = 55;
                            leadEntity["ownerid"] = new EntityReference("systemuser", user.Id);
                        }
                        else if (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1 || ((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)
                        {
                            tracingService.Trace("Set lead Assigned to False and trigger next@bat only for Web Leads (new leads only and not cloned)");
                            variables.leadAssigned = false;
                            step = 56;
                            if (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1 && !leadEntity.Attributes.Contains("fdx_leadid"))
                                leadEntity["fdx_snb"] = true;
                        }
                    }

                    #region Set the address field as per the account if there is an existing account....
                    if (variables.accountid != Guid.Empty)
                    {
                        step = 57;
                        EntityCollection accountCollection = this.getExisitingAccountDetails(variables.accountid);
                        if (accountCollection.Entities.Count > 0)
                        {
                            Entity account = new Entity();
                            account = accountCollection.Entities[0];
                            step = 58;
                            if (!account.Attributes.Contains("fdx_goldmineaccountnumber"))
                            {
                                step = 59;
                                apiParm = this.setApiParmFromAccountEntity(account);
                                url = variables.smartCrmSyncWebServiceUrl + "/lead/createlead?" + apiParm;
                            }
                            else
                            {
                                step = 60;
                                variables.acc_gmaccountno_exist = true;
                                leadEntity["fdx_goldmineaccountnumber"] = account.Attributes["fdx_goldmineaccountnumber"].ToString();
                                leadEntity["fdx_gonogo"] = account.Attributes["fdx_gonogo"];
                                step = 61;
                                ProspectData prospectData = this.GetProspectDataFromAccount(account);
                                step = 62;
                                CopyLeadProspectDataToSharedVariable(context.SharedVariables, prospectData);
                            }
                        }
                    }
                    else
                    {
                        step = 63;
                        url = variables.smartCrmSyncWebServiceUrl + "/lead/createlead?" + apiParm;
                    }
                    #endregion

                    #region Call and update from API....
                    tracingService.Trace(url);
                    Lead leadObj = new Lead();
                    if (!variables.acc_gmaccountno_exist)
                    {
                        step = 64;
                        const string token = "8b6asd7-0775-4278-9bcb-c0d48f800112";
                        var uri = new Uri(url);
                        step = 65;
                        var request = WebRequest.Create(uri);
                        request.Method = WebRequestMethods.Http.Post;
                        request.ContentType = "application/json";
                        request.ContentLength = 0;
                        request.Headers.Add("Authorization", token);

                        step = 66;
                        using (var getResponse = request.GetResponse())
                        {
                            step = 67;
                            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Lead));

                            step = 68;
                            leadObj = (Lead)serializer.ReadObject(getResponse.GetResponseStream());

                            step = 69;
                            EntityCollection priceLists = this.GetPriceListByName(leadObj.priceListName, service);
                            step = 70;
                            EntityCollection prospectGroups = this.GetProspectGroupByName(leadObj.prospectGroup, service);
                            step = 71;
                            ProspectData prospectData = this.GetProspectDataFromWebService(leadObj);
                            prospectData.PriceListName = leadObj.priceListName;

                            step = 72;
                            if (priceLists.Entities.Count == 1)
                                prospectData.PriceListId = priceLists.Entities[0].Id;

                            step = 73;
                            if (prospectGroups.Entities.Count == 1)
                                prospectData.ProspectGroupId = prospectGroups.Entities[0].Id;

                            step = 74;
                            leadEntity["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                            if (leadObj.goNoGo)
                            {
                                leadEntity["fdx_gonogo"] = new OptionSetValue(756480000);
                            }
                            else
                            {
                                leadEntity["fdx_gonogo"] = new OptionSetValue(756480001);
                            }

                            step = 75;
                            if (variables.accountid != Guid.Empty)
                            {
                                step = 76;
                                IOrganizationService impersonatedService = serviceFactory.CreateOrganizationService(null);

                                step = 77;
                                Entity acc = new Entity("account")
                                {
                                    Id = variables.accountid
                                };
                                acc.Attributes["fdx_goldmineaccountnumber"] = leadObj.goldMineId;
                                acc.Attributes["fdx_gonogo"] = leadObj.goNoGo ? new OptionSetValue(756480000) : new OptionSetValue(756480001);

                                step = 78;
                                this.UpdateProspectDataOnAccount(acc, prospectData);
                                step = 79;
                                impersonatedService.Update(acc);
                            }

                            step = 80;
                            CopyLeadProspectDataToSharedVariable(context.SharedVariables, prospectData);
                            step = 81;
                            tracingService.Trace(GetProspectDataString(prospectData));
                            tracingService.Trace("Prospect Data Updated in AssignLead_Create");
                        }
                    }
                    #endregion

                    #region Condition to assign Lead to Lead Review Team....
                    if (!variables.leadAssigned && !leadObj.goNoGo && ((((OptionSetValue)leadEntity["leadsourcecode"]).Value == 1) || (((OptionSetValue)leadEntity["leadsourcecode"]).Value == 4)))
                    {
                        step = 82;
                        leadEntity["fdx_snb"] = false;
                        QueryExpression teamQuery = CRMQueryExpression.getQueryExpression("team", new ColumnSet("name"), new CRMQueryExpression[] { new CRMQueryExpression("name", ConditionOperator.Equal, "Lead Review Team") });

                        step = 83;
                        EntityCollection teamCollection = service.RetrieveMultiple(teamQuery);

                        step = 84;
                        if (teamCollection.Entities.Count > 0)
                        {
                            Entity team = teamCollection.Entities[0];
                            leadEntity["ownerid"] = new EntityReference("team", team.Id);
                        }

                        tracingService.Trace("Condition to assign Lead to Lead Review Team");
                    }
                    #endregion

                    #region Condition if Cloned Lead: override owner of cloned lead = originating lead's owner

                    if (leadEntity.Attributes.Contains("fdx_leadid"))
                    {
                        step = 85;
                        Entity OriginatingLead = service.Retrieve("lead", ((EntityReference)leadEntity.Attributes["fdx_leadid"]).Id, new ColumnSet("ownerid", "owningteam"));

                        step = 86;
                        if (OriginatingLead.Attributes.Contains("owningteam"))
                            leadEntity["ownerid"] = new EntityReference(OriginatingLead["owningteam"] != null ? "team" : "systemuser", ((EntityReference)OriginatingLead.Attributes["ownerid"]).Id);
                        else
                            leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)OriginatingLead.Attributes["ownerid"]).Id);

                        tracingService.Trace("Condition if Cloned Lead: override owner of cloned lead = originating lead's owner");
                    }
                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException(string.Format("An error occurred in the LeadAssignment_Create plug-in at Step {0}.", step), ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("LeadAssignment_Create: step {0}, {1}", step, ex.ToString());
                    throw;
                }
            }
        }

        private EntityCollection getExisitingAccountDetails(Guid _accountId)
        {
            ColumnSet accountColumns = new ColumnSet("name", "fdx_goldmineaccountnumber", "fdx_gonogo", "address1_line1", "address1_line2", "address1_city", "fdx_stateprovinceid", "fdx_zippostalcodeid", "telephone1");
            accountColumns.AddColumns("fdx_prospectgroup", "defaultpricelevelid", "fdx_prospectpriority", "fdx_prospectscore", "fdx_prospectpercentile", "fdx_ratesource", "fdx_pprrate", "fdx_subrate", "fdx_prospectradius", "fdx_prospectdatalastupdated");
            QueryExpression accountQuery = CRMQueryExpression.getQueryExpression("account", accountColumns, new CRMQueryExpression[] { new CRMQueryExpression("accountid", ConditionOperator.Equal, _accountId) });
            EntityCollection accountCollection = service.RetrieveMultiple(accountQuery);

            return accountCollection;
        }

        private void formatPhoneNumbers()
        {
            if (leadEntity.Attributes.Contains("telephone1"))
                leadEntity.Attributes["telephone1"] = Regex.Replace(leadEntity.Attributes["telephone1"].ToString(), @"[^0-9]+", "");

            if (leadEntity.Attributes.Contains("telephone2"))
                leadEntity.Attributes["telephone2"] = Regex.Replace(leadEntity.Attributes["telephone2"].ToString(), @"[^0-9]+", "");

            if (leadEntity.Attributes.Contains("telephone3"))
                leadEntity.Attributes["telephone3"] = Regex.Replace(leadEntity.Attributes["telephone3"].ToString(), @"[^0-9]+", "");
        }

        private string setApiParmFromAccountEntity(Entity _account) 
        {
            string apiParm = "";

            apiParm = string.Format("Zip={0}&Phone1={1}", (service.Retrieve("fdx_zipcode", ((EntityReference)_account.Attributes["fdx_zippostalcodeid"]).Id, new ColumnSet("fdx_zipcode"))).Attributes["fdx_zipcode"].ToString(), Regex.Replace(_account.Attributes["telephone1"].ToString(), @"[^0-9]+", ""));

            if (_account.Attributes.Contains("name"))
                apiParm += string.Format("{1}Company={0}", _account.Attributes["name"].ToString(), apiParm != "" ? "&" : "");

            if (_account.Attributes.Contains("address1_line1"))
                apiParm += string.Format("{1}Address1={0}", _account.Attributes["address1_line1"].ToString(), apiParm != "" ? "&" : "");

            if (_account.Attributes.Contains("address1_line2"))
                apiParm += string.Format("{1}Address2={0}", _account.Attributes["address1_line2"].ToString(), apiParm != "" ? "&" : "");

            if (_account.Attributes.Contains("address1_city"))
                apiParm += string.Format("{1}City={0}", _account.Attributes["address1_city"].ToString(), apiParm != "" ? "&" : "");

            if (_account.Attributes.Contains("fdx_stateprovinceid"))
                apiParm += string.Format("{1}State={0}", (service.Retrieve("fdx_state", ((EntityReference)_account.Attributes["fdx_stateprovinceid"]).Id, new ColumnSet("fdx_statecode"))).Attributes["fdx_statecode"].ToString(), apiParm != "" ? "&" : "");

            return apiParm;
        }

        private string setApiParmFromLeadEntity()
        {
            string apiParm = "";

            if(variables.zipcodetext != "")
                apiParm = string.Format("Zip={0}", variables.zipcodetext);

            if (variables.firstName != "")
                apiParm += string.Format("{2}Contact={0} {1}", variables.firstName, variables.lastName, apiParm != "" ? "&" : "");

            if (variables.phone != "")
                apiParm += string.Format("{1}Phone1={0}", variables.phone, apiParm != "" ? "&" : "");

            if (variables.companyName != "")
                apiParm += string.Format("{1}Company={0}", variables.companyName, apiParm != "" ? "&" : "");

            if (variables.title != "")
                apiParm += string.Format("{1}Title={0}", variables.title, apiParm != "" ? "&" : "");

            if (variables.address1 != "")
                apiParm += string.Format("{1}Address1={0}", variables.address1, apiParm != "" ? "&" : "");

            if (variables.address2 != "")
                apiParm += string.Format("{1}Address2={0}", variables.address2, apiParm != "" ? "&" : "");

            if (variables.city != "")
                apiParm += string.Format("{1}City={0}", variables.city, apiParm != "" ? "&" : "");

            if (variables.state != "")
                apiParm += string.Format("{1}State={0}", variables.state, apiParm != "" ? "&" : "");

            return apiParm;
        }

        private void setConnectedLead(Entity _lead)
        {
            if (!leadEntity.Attributes.Contains("fdx_leadid"))
                leadEntity["fdx_leadid"] = new EntityReference("lead", _lead.Id);
        }

        private void setExistingContact(Entity _entity, string _entityType)
        {
            switch(_entityType)
            {
                case "lead":
                    if (_entity.Attributes.Contains("contactid"))
                        leadEntity["contactid"] = new EntityReference("contact", ((EntityReference)_entity.Attributes["contactid"]).Id);

                    if (_entity.Attributes.Contains("parentcontactid"))
                        leadEntity["parentcontactid"] = new EntityReference("contact", ((EntityReference)_entity.Attributes["parentcontactid"]).Id);
                    break;
                case "contact":
                    leadEntity["contactid"] = new EntityReference("contact", _entity.Id);
                    leadEntity["parentcontactid"] = new EntityReference("contact", _entity.Id);
                    break;
            }
            
        }

        private void setExistingAccount(Entity _entity, string _entityType)
        {
            switch (_entityType)
            {
                case "lead":
                    if (_entity.Attributes.Contains("accountid"))
                    {
                        leadEntity["accountid"] = new EntityReference("account", ((EntityReference)_entity.Attributes["accountid"]).Id);
                        variables.accountid = ((EntityReference)_entity.Attributes["accountid"]).Id;
                    }

                    if (_entity.Attributes.Contains("parentaccountid"))
                    {
                        leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)_entity.Attributes["parentaccountid"]).Id);
                        variables.accountid = ((EntityReference)_entity.Attributes["parentaccountid"]).Id;
                    }
                    break;
                case "contact":
                    if (_entity.Attributes.Contains("parentcustomerid"))
                    {
                        if (((EntityReference)_entity.Attributes["parentcustomerid"]).LogicalName == "account")
                        {
                            leadEntity["accountid"] = new EntityReference("account", ((EntityReference)_entity.Attributes["parentcustomerid"]).Id);
                            leadEntity["parentaccountid"] = new EntityReference("account", ((EntityReference)_entity.Attributes["parentcustomerid"]).Id);
                            variables.accountid = ((EntityReference)_entity.Attributes["parentcustomerid"]).Id;
                        }
                    }
                    break;
                case "account":
                    leadEntity["accountid"] = new EntityReference("account", _entity.Id);
                    leadEntity["parentaccountid"] = new EntityReference("account", _entity.Id);
                    variables.accountid = _entity.Id;
                    break;
            }
            
        }

        private void setLeadOwner(Entity _entity)
        {
            
            if (_entity.Attributes.Contains("owningteam"))
                leadEntity["ownerid"] = new EntityReference(_entity["owningteam"] != null ? "team" : "systemuser", ((EntityReference)_entity.Attributes["ownerid"]).Id);
            else
                leadEntity["ownerid"] = new EntityReference("systemuser", ((EntityReference)_entity.Attributes["ownerid"]).Id);
        }

        private Entity getAccountMatchedLead(Entity _account)
        {
            Entity lead = new Entity();
            QueryExpression queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("accountid", ConditionOperator.Equal, _account.Id), new CRMQueryExpression("parentaccountid", ConditionOperator.Equal, _account.Id) }, LogicalOperator.Or);
            EntityCollection collection = service.RetrieveMultiple(queryExp);

            if (collection.Entities.Count > 0)
                lead = collection.Entities[0];

            return lead;
        }

        private Entity getAccountMatchedLead(Entity _account, IOrganizationService _service)
        {
            Entity lead = new Entity();
            QueryExpression queryExp = CRMQueryExpression.getQueryExpression("lead", new ColumnSet("leadid", "contactid", "parentcontactid", "accountid", "parentaccountid", "ownerid", "owningteam"), new CRMQueryExpression[] { new CRMQueryExpression("accountid", ConditionOperator.Equal, _account.Id), new CRMQueryExpression("parentaccountid", ConditionOperator.Equal, _account.Id) }, LogicalOperator.Or);
            EntityCollection collection = _service.RetrieveMultiple(queryExp);

            if (collection.Entities.Count > 0)
                lead = collection.Entities[0];

            return lead;
        }

        private void updateExistingAccount(Guid _leadId)
        {
            //Code added as part of S893 by Meghana on Lead create for Incoming Phone call,web creation/import
            //wf name - fdx_Lead update Existing Account
            tracingService.Trace("OnDemand workflow started");
            string wfid = "8622B036-D051-489A-9C0F-B168554B1938";
            ExecuteWorkflowRequest request = new ExecuteWorkflowRequest
            {
                EntityId = _leadId,
                WorkflowId = new Guid(wfid)
            };
            tracingService.Trace("OnDemand workflow in loop");
            service.Execute(request);
            tracingService.Trace("OnDemand workflow req executed");
        }

        private ProspectData GetProspectDataFromStub()
        {
            ProspectData prospectData = new ProspectData();
            prospectData.ProspectGroupId = new Guid("9B3945FC-2728-E811-811D-3863BB34CB20");
            prospectData.PriceListId = new Guid("8A826A97-0B26-E811-811C-3863BB35EF70");
            prospectData.Priority = Convert.ToDecimal(1);
            prospectData.Score = Convert.ToDecimal(2);
            prospectData.Percentile = Convert.ToDecimal(3);
            prospectData.RateSource = "Stub";
            prospectData.PPRRate = Convert.ToDecimal(1);
            prospectData.SubRate = Convert.ToDecimal(2);
            prospectData.Radius = 2;
            return prospectData;
        }

        private ProspectData GetProspectDataFromWebService(Lead lead)
        {
            ProspectData prospectData = new ProspectData();
            prospectData.ProspectGroupName = lead.prospectGroup;
            prospectData.PriceListName = lead.priceListName;
            prospectData.Priority = lead.prospectPriority;
            prospectData.Score = lead.prspectScore;
            prospectData.Percentile = lead.prospectPercentile;
            prospectData.RateSource = lead.rateSource;
            prospectData.PPRRate = lead.pprRate;
            prospectData.SubRate = lead.subRate;
            prospectData.Radius = lead.prospectRadius;
            return prospectData;
        }

        private ProspectData GetProspectDataFromAccount(Entity account)
        {
            ProspectData prospectData = new ProspectData();
            if (account.Contains("fdx_prospectgroup"))
                prospectData.ProspectGroupId = ((EntityReference)account["fdx_prospectgroup"]).Id;
            if (account.Contains("defaultpricelevelid"))
            {
                EntityReference priceList = (EntityReference)account["defaultpricelevelid"];
                prospectData.PriceListId = priceList.Id;
            }
            if (account.Contains("fdx_prospectpriority"))
                prospectData.Priority = (decimal)account["fdx_prospectpriority"];
            if (account.Contains("fdx_prospectscore"))
                prospectData.Score = (decimal)account["fdx_prospectscore"];
            if (account.Contains("fdx_prospectpercentile"))
                prospectData.Percentile = (decimal)account["fdx_prospectpercentile"];
            if (account.Contains("fdx_ratesource"))
                prospectData.RateSource = (string)account["fdx_ratesource"];
            if (account.Contains("fdx_pprrate"))
                prospectData.PPRRate = ((Money)account["fdx_pprrate"]).Value;
            if (account.Contains("fdx_subrate"))
                prospectData.SubRate = ((Money)account["fdx_subrate"]).Value;
            if (account.Contains("fdx_prospectradius"))
                prospectData.Radius = (int)account["fdx_prospectradius"];
            if (account.Contains("fdx_prospectdatalastupdated"))
                prospectData.LastUpdated = (DateTime)account["fdx_prospectdatalastupdated"];
            return prospectData;
        }

        private string GetProspectDataString(ProspectData prospectData)
        {
            string traceString = "ProspectGroupName=" + prospectData.ProspectGroupName + Environment.NewLine;
            traceString += "PriceListName=" + prospectData.PriceListName + Environment.NewLine;
            traceString += "Priority=" + Convert.ToString(prospectData.Priority) + Environment.NewLine;
            traceString += "Score=" + Convert.ToString(prospectData.Score) + Environment.NewLine;
            traceString += "Percentile=" + Convert.ToString(prospectData.Percentile) + Environment.NewLine;
            traceString += "RateSource=" + prospectData.RateSource + Environment.NewLine;
            traceString += "PPRRate=" + Convert.ToString(prospectData.PPRRate) + Environment.NewLine;
            traceString += "SubRate=" + Convert.ToString(prospectData.SubRate) + Environment.NewLine;
            traceString += "Radius=" + Convert.ToString(prospectData.Radius) + Environment.NewLine;
            return traceString;
        }

        private void CopyLeadProspectDataToSharedVariable(ParameterCollection contextSharedVariable, ProspectData prospectData)
        {
            contextSharedVariable.Add("ProspectData", true);
            if (prospectData.ProspectGroupId.HasValue && !prospectData.ProspectGroupId.Equals(Guid.Empty))
                contextSharedVariable.Add("fdx_prospectgroup", prospectData.ProspectGroupId.Value);
            if (prospectData.PriceListId.HasValue && !prospectData.PriceListId.Equals(Guid.Empty))
                contextSharedVariable.Add("fdx_pricelist", prospectData.PriceListId.Value);
            if (prospectData.Priority.HasValue)
                contextSharedVariable.Add("fdx_prospectpriority", prospectData.Priority);
            if (prospectData.Score.HasValue)
                contextSharedVariable.Add("fdx_prospectscore", prospectData.Score);
            if (prospectData.Percentile.HasValue)
                contextSharedVariable.Add("fdx_prospectpercentile", prospectData.Percentile);
            if (!string.IsNullOrEmpty(prospectData.RateSource))
                contextSharedVariable.Add("fdx_ratesource", prospectData.RateSource);
            if (prospectData.PPRRate.HasValue)
                contextSharedVariable.Add("fdx_pprrate", prospectData.PPRRate);
            if (prospectData.SubRate.HasValue)
                contextSharedVariable.Add("fdx_subrate", prospectData.SubRate);
            if (prospectData.Radius.HasValue)
                contextSharedVariable.Add("fdx_prospectradius", prospectData.Radius);
            if (prospectData.LastUpdated.HasValue)
                contextSharedVariable.Add("fdx_prospectdatalastupdated", prospectData.LastUpdated.Value);
        }

        private void UpdateProspectDataOnAccount(Entity accountRecord, ProspectData prospectData)
        {
            if (prospectData.ProspectGroupId.HasValue && !prospectData.ProspectGroupId.Equals(Guid.Empty))
                accountRecord["fdx_prospectgroup"] = new EntityReference("fdx_prospectgroup", prospectData.ProspectGroupId.Value);
            if (prospectData.PriceListId.HasValue && !prospectData.PriceListId.Equals(Guid.Empty))
                accountRecord["defaultpricelevelid"] = new EntityReference("pricelevel", prospectData.PriceListId.Value);
            if (prospectData.Priority.HasValue)
                accountRecord["fdx_prospectpriority"] = prospectData.Priority;
            if (prospectData.Score.HasValue)
                accountRecord["fdx_prospectscore"] = prospectData.Score;
            if (prospectData.Percentile.HasValue)
                accountRecord["fdx_prospectpercentile"] = prospectData.Percentile;
            if (!string.IsNullOrEmpty(prospectData.RateSource))
                accountRecord["fdx_ratesource"] = prospectData.RateSource;
            if (prospectData.PPRRate.HasValue)
                accountRecord["fdx_pprrate"] = new Money(prospectData.PPRRate.Value);
            if (prospectData.SubRate.HasValue)
                accountRecord["fdx_subrate"] = new Money(prospectData.SubRate.Value);
            if (prospectData.Radius.HasValue)
                accountRecord["fdx_prospectradius"] = prospectData.Radius;
            accountRecord["fdx_prospectdatalastupdated"] = DateTime.UtcNow;
        }

        private EntityCollection GetPriceListByName(string priceListName, IOrganizationService crmService)
        {
            QueryByAttribute queryByPriceList = new QueryByAttribute("pricelevel");
            queryByPriceList.ColumnSet = new ColumnSet("pricelevelid");
            queryByPriceList.AddAttributeValue("name", priceListName);
            return crmService.RetrieveMultiple(queryByPriceList);
        }

        private EntityCollection GetProspectGroupByName(string prospectGroupName, IOrganizationService crmService)
        {
            QueryByAttribute queryByProspectGroup = new QueryByAttribute("fdx_prospectgroup");
            queryByProspectGroup.ColumnSet = new ColumnSet("fdx_prospectgroupid");
            queryByProspectGroup.AddAttributeValue("fdx_name", prospectGroupName);
            return crmService.RetrieveMultiple(queryByProspectGroup);
        }
    }
}
