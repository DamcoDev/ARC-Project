using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace DuplicateCasePlugin
{
    public class Case: IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationService service = ((IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory))).
            CreateOrganizationService(new Guid?(context.UserId));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if ((context.InputParameters.Contains("Target")) && (context.InputParameters["Target"] is Entity))
            {
                Entity entity = (Entity)context.InputParameters["Target"];
                if (entity.LogicalName.ToLower() != "incident") return;
                if (context.Depth > 1)
                {
                    return;
                }
                try
                {
                    if (entity.Attributes.Contains("title"))
                    {
                        string email = entity.GetAttributeValue<string>("title").ToString();
                        //handle the above line of code accordingly based on dataType(String, EntityReference, Datetime etc.)
                        QueryExpression contactQuery = new QueryExpression("incident");
                        contactQuery.ColumnSet = new ColumnSet("title");
                        contactQuery.Criteria.AddCondition("title", ConditionOperator.Equal, email);
                        EntityCollection contactColl = service.RetrieveMultiple(contactQuery);
                        if (contactColl.Entities.Count > 0)
                        {
                            throw new InvalidPluginExecutionException("Duplicates found !!!");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw (ex);
                }
            }
        }
    }
}
