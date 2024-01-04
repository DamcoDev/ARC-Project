using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;

namespace CompetitorValidation
{
    public class Competitor:CodeActivity
    {
        [Input("Opportunity")]
        [RequiredArgument]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1)
            //{
            //    return;
            //}
            Guid OpportunityGuid = Opportunity.Get<EntityReference>(context).Id;
            executeBusinessLogic(OpportunityGuid, service, context);
        }

        private void executeBusinessLogic(Guid opportunityGuid, IOrganizationService service, CodeActivityContext context)
        {
            try
            {
                QueryExpression query = new QueryExpression("competitor");
                query.ColumnSet = new ColumnSet(new string[] { "competitorid", "name" });

                LinkEntity linkEntity1 = new LinkEntity("competitor", "opportunitycompetitors", "competitorid", "competitorid", JoinOperator.Inner);
                LinkEntity linkEntity2 = new LinkEntity("opportunitycompetitors", "opportunity", "opportunityid", "opportunityid", JoinOperator.Inner);

                linkEntity1.LinkEntities.Add(linkEntity2);
                query.LinkEntities.Add(linkEntity1);

                linkEntity2.LinkCriteria = new FilterExpression();
                linkEntity2.LinkCriteria.AddCondition(new ConditionExpression("opportunityid", ConditionOperator.Equal, opportunityGuid));

                EntityCollection centreCollection = service.RetrieveMultiple(query);
                if (centreCollection.Entities.Count > 0)
                {
                 
                }
                else
                {
                    throw new InvalidPluginExecutionException("You must add atleast one Competitor to continue!");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }
}
