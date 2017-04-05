using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAssignmentPlugin
{
    public class AssignLead_Update : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            //--> This class will not be executed as per the new DSO lead assignment....
            //Extract the tracing service for use in debugging sandboxed plug-ins....
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            //Obtain execution contest from the service provider....
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            //Call Input parameter collection to get all the data passes....
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity leadEntity = (Entity)context.InputParameters["Target"];

                if (leadEntity.LogicalName != "lead")
                    return;

                //Entity object to update Lead....
                Entity leadUpdate = new Entity()
                {
                    LogicalName = "lead",
                    Id = leadEntity.Id
                };

                try
                {
                    IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                    IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                    //Get current user information....
                    WhoAmIResponse response = (WhoAmIResponse)service.Execute(new WhoAmIRequest());

                    //Fetch data from lead entity....
                    bool isGroupPractice = leadEntity.GetAttributeValue<bool>("fdx_grppracactice");
                    int step = 0;


                    if (isGroupPractice)
                    {
                        Guid state = new Guid ();
                        Entity leadState = service.Retrieve("lead", leadEntity.Id, new ColumnSet("fdx_stateprovince"));
                        if(leadState.Attributes.Contains("fdx_stateprovince"))
                            state = ((EntityReference)leadState.Attributes["fdx_stateprovince"]).Id;
                        step = 1;
                        #region 1st check --> Get DSO rep based on territory assigned to State....
                        Entity stateEntity = service.Retrieve("fdx_state", state, new ColumnSet("fdx_territory"));

                        if (stateEntity.Attributes.Contains("fdx_territory"))
                        {
                            step = 21;
                            Entity territory = service.Retrieve("territory", ((EntityReference)stateEntity.Attributes["fdx_territory"]).Id, new ColumnSet("managerid"));
                            if (territory.Attributes.Contains("managerid"))
                                leadUpdate["ownerid"] = new EntityReference("systemuser", ((EntityReference)territory["managerid"]).Id);

                            service.Update(leadUpdate);
                        }
                        #endregion
                    }                                        
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the AssignLead_Update plug-in.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("AssignLead_Update: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
