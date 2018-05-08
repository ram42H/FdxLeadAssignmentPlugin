using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxLeadAssignmentPlugin
{
    public class UpdateProspectData : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            tracingService.Trace("Inside the UpdateProspectData");
            if (context.ParentContext != null && context.ParentContext.SharedVariables.ContainsKey("ProspectData"))
            {
                IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                Guid leadId = (Guid) context.OutputParameters["id"];
                tracingService.Trace("Lead Id" + leadId.ToString());
                Entity leadRecord = new Entity("lead", leadId);
                ProspectData prospectData = GetProspectDataFromContext(context.ParentContext.SharedVariables);
                UpdateProspectDataOnLead(leadRecord, prospectData);
                IOrganizationService impersonatedService = serviceFactory.CreateOrganizationService(null);
                impersonatedService.Update(leadRecord);
            }
        }

        private ProspectData GetProspectDataFromContext(ParameterCollection contextSharedVariable)
        {
            ProspectData prospectData = new ProspectData();
            if (contextSharedVariable.ContainsKey("fdx_prospectgroup"))
                prospectData.ProspectGroupId = (Guid)contextSharedVariable["fdx_prospectgroup"];
            if (contextSharedVariable.ContainsKey("fdx_pricelist"))
                prospectData.PriceListId = (Guid)contextSharedVariable["fdx_pricelist"];
            if (contextSharedVariable.ContainsKey("fdx_prospectpriority"))
                prospectData.Priority = (decimal)contextSharedVariable["fdx_prospectpriority"];
            if (contextSharedVariable.ContainsKey("fdx_prospectscore"))
                prospectData.Score = (decimal)contextSharedVariable["fdx_prospectscore"];
            if (contextSharedVariable.ContainsKey("fdx_prospectpercentile"))
                prospectData.Percentile = (decimal)contextSharedVariable["fdx_prospectpercentile"];
            if (contextSharedVariable.ContainsKey("fdx_ratesource"))
                prospectData.RateSource = (string)contextSharedVariable["fdx_ratesource"];
            if (contextSharedVariable.ContainsKey("fdx_pprrate"))
                prospectData.PPRRate = (decimal)contextSharedVariable["fdx_pprrate"];
            if (contextSharedVariable.ContainsKey("fdx_subrate"))
                prospectData.SubRate = (decimal)contextSharedVariable["fdx_subrate"];
            if (contextSharedVariable.ContainsKey("fdx_prospectradius"))
                prospectData.Radius = (int)contextSharedVariable["fdx_prospectradius"];
            if (contextSharedVariable.ContainsKey("fdx_prospectdatalastupdated"))
                prospectData.LastUpdated = (DateTime)contextSharedVariable["fdx_prospectdatalastupdated"];
            return prospectData;
        }

        private void UpdateProspectDataOnLead(Entity leadRecord, ProspectData prospectData)
        {
            if (prospectData.ProspectGroupId.HasValue && !prospectData.ProspectGroupId.Equals(Guid.Empty))
                leadRecord["fdx_prospectgroup"] = new EntityReference("fdx_prospectgroup", prospectData.ProspectGroupId.Value);
            if (prospectData.PriceListId.HasValue && !prospectData.PriceListId.Equals(Guid.Empty))
                leadRecord["fdx_pricelist"] = new EntityReference("pricelevel", prospectData.PriceListId.Value);
            if (prospectData.Priority.HasValue)
                leadRecord["fdx_prospectpriority"] = prospectData.Priority;
            if (prospectData.Score.HasValue)
                leadRecord["fdx_prospectscore"] = prospectData.Score;
            if (prospectData.Percentile.HasValue)
                leadRecord["fdx_prospectpercentile"] = prospectData.Percentile;
            if (!string.IsNullOrEmpty(prospectData.RateSource))
                leadRecord["fdx_ratesource"] = prospectData.RateSource;
            if (prospectData.PPRRate.HasValue)
                leadRecord["fdx_pprrate"] = new Money(prospectData.PPRRate.Value);
            if (prospectData.SubRate.HasValue)
                leadRecord["fdx_subrate"] = new Money(prospectData.SubRate.Value);
            if (prospectData.Radius.HasValue)
                leadRecord["fdx_prospectradius"] = prospectData.Radius;
            if (prospectData.LastUpdated.HasValue)
            {
                leadRecord["fdx_prospectdatalastupdated"] = prospectData.LastUpdated.Value;
            }
            else
            {
                leadRecord["fdx_prospectdatalastupdated"] = DateTime.UtcNow;
            }
        }
    }
}
