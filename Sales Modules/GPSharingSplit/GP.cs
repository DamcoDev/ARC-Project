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

namespace GPSharingSplit
{
    public class GP:CodeActivity
    {
        [Input("Opportunity")]
        [RequiredArgument]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }

        [Input("Client Manager")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> ClientManager { get; set; }

        [Input("Service Delivery Head")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> ServiceDeliveryHead { get; set; }

        [Input("EST GP")]
        [RequiredArgument]
        public InArgument<decimal> ESTGP { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1)
            //{
            //    return;
            //}
            Guid ClientManagerGuid = ClientManager.Get<EntityReference>(context).Id;
            Guid ServiceDeliveryManagerGuid = ServiceDeliveryHead.Get<EntityReference>(context).Id;
            decimal EstGP = ESTGP.Get<decimal>(context);
            Guid OpportunityGuid = Opportunity.Get<EntityReference>(context).Id;
            executeBusinessLogic(ClientManagerGuid, ServiceDeliveryManagerGuid, EstGP, OpportunityGuid, service);

        }

        private void executeBusinessLogic(Guid clientManagerGuid,Guid ServiceDeliveryManagerGuid, decimal estGP, Guid opportunityGuid, IOrganizationService service)
        {
            try
            {
                Entity Opportunity = service.Retrieve("opportunity", opportunityGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet("parentaccountid"));

                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet("arc_opportunity", "arc_gpsharing", "arc_estgp");
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("arc_opportunity", ConditionOperator.Equal, opportunityGuid));
                q1.Criteria = fe;
                q1.EntityName = "arc_opportunitysharing";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {
                    foreach(Entity c in ec.Entities)
                    {
                        Guid GPSharingGuid = new Guid(c.Attributes["arc_opportunitysharingid"].ToString());
                        decimal GPSharing = Convert.ToDecimal(c.Attributes["arc_gpsharing"]);
                        //decimal EstGP =((Money)c.Attributes["arc_estgp"]).Value;
                        decimal GP1Value = (estGP * GPSharing) / 100;

                        Entity GP = new Entity("arc_opportunitysharing");
                        GP["arc_estgp"] = new Money(estGP);
                        GP["arc_gpvalue"] = new Money(GP1Value);
                        GP.Id = GPSharingGuid;
                        service.Update(GP);
                    }
                }
                else
                {
                    Guid AccountGuid = new Guid();
                    if (Opportunity.Attributes.Contains("parentaccountid"))
                    {
                        AccountGuid = ((EntityReference)Opportunity["parentaccountid"]).Id;
                    }

                    Entity GP1 = new Entity("arc_opportunitysharing");
                    GP1["arc_clientmanager"] = new EntityReference("systemuser", clientManagerGuid);
                    GP1["arc_estgp"] = new Money(estGP);
                    GP1["arc_opportunity"] = new EntityReference("opportunity", opportunityGuid);
                    GP1["arc_gpsharing"] = Convert.ToDouble(75);
                    decimal GP1Value = (estGP * 75) / 100;
                    GP1["arc_gpvalue"] = new Money(GP1Value);
                    if (AccountGuid != Guid.Empty)
                    {
                        GP1["arc_account"] = new EntityReference("account", AccountGuid);
                    }
                    Guid GP1Guid = service.Create(GP1);



                    Entity GP2 = new Entity("arc_opportunitysharing");
                    GP2["arc_clientmanager"] = new EntityReference("systemuser", ServiceDeliveryManagerGuid);
                    GP2["arc_estgp"] = new Money(estGP);
                    GP2["arc_opportunity"] = new EntityReference("opportunity", opportunityGuid);
                    GP2["arc_gpsharing"] = Convert.ToDouble(25);
                    decimal GP2Value = (estGP * 25) / 100;
                    GP2["arc_gpvalue"] = new Money(GP2Value);
                    GP2["arc_remaininggp"] = Convert.ToDouble(0);
                    if (AccountGuid != Guid.Empty)
                    {
                        GP2["arc_account"] = new EntityReference("account", AccountGuid);
                    }
                    Guid GP2Guid = service.Create(GP2);
                }
            }
            catch (InvalidPluginExecutionException ex)
            {
                if (ex != null)
                {

                    throw new InvalidPluginExecutionException("1) Unable to complete the Operation." + ex.Message + "Contact Your Administrator");
                }
                else
                    throw new InvalidPluginExecutionException("2) Unable to complete the Operation. Contact Administrator");
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("3) Unable to complete the Operation.", ex);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("4) Unable to complete the Operation.", ex);
            }
        }
    }
}
