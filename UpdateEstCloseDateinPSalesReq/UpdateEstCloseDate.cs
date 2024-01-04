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


namespace UpdateEstCloseDateinPSalesReq
{
    public class UpdateEstCloseDate: CodeActivity
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
            //Guid ClientManagerGuid = ClientManager.Get<EntityReference>(context).Id;
            //Guid ServiceDeliveryManagerGuid = ServiceDeliveryHead.Get<EntityReference>(context).Id;
            //decimal EstGP = ESTGP.Get<decimal>(context);
            Guid OpportunityGuid = Opportunity.Get<EntityReference>(context).Id;
            executeBusinessLogic(OpportunityGuid, service);

        }

        private void executeBusinessLogic(Guid opportunityGuid, IOrganizationService service)
        {
            try
            {
                Entity Opportunity = service.Retrieve("opportunity", opportunityGuid, new ColumnSet("estimatedclosedate"));
                if(Opportunity.Attributes.Contains("estimatedclosedate"))
                {
                    DateTime dt = Convert.ToDateTime(Opportunity.Attributes["estimatedclosedate"].ToString());

                    QueryExpression q1 = new QueryExpression();
                    q1.ColumnSet = new ColumnSet("regardingobjectid");
                    FilterExpression fe = new FilterExpression(LogicalOperator.And);
                    fe.AddCondition(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, opportunityGuid));
                    q1.Criteria = fe;
                    q1.EntityName = "its_presalesrequest";
                    EntityCollection ec = service.RetrieveMultiple(q1);
                    if (ec.Entities.Count > 0)
                    {
                        foreach (Entity c in ec.Entities)
                        {
                            Guid PreSalesRequestGuid = new Guid(c.Attributes["activityid"].ToString());

                            Entity PreSalesRequest = new Entity("its_presalesrequest");
                            PreSalesRequest["its_opportunityestclosedate"] = dt;
                            PreSalesRequest.Id = PreSalesRequestGuid;
                            service.Update(PreSalesRequest);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }
}
