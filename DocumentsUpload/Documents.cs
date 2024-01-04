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

namespace DocumentsUpload
{
    public class Documents:CodeActivity
    {
        [Input("Opportunity")]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }

        [Input("Project Handover")]
        [ReferenceTarget("itspsa_projecthandover")]
        public InArgument<EntityReference> ProjectHandover { get; set; }

        [Input("Project")]
        [ReferenceTarget("msdyn_project")]
        public InArgument<EntityReference> Project { get; set; }

        [Input("Flag")]
        [RequiredArgument]
        public InArgument<int> FlagValue { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1) { return; }
            int Flag = FlagValue.Get<int>(context);           
            if (Flag == 1) //1-Project Handover
            {
                Guid OpportunityGUID = Opportunity.Get<EntityReference>(context).Id;
                Guid ProjectHandoverGUID = ProjectHandover.Get<EntityReference>(context).Id;
                executeProjectHandoverBusinessLogic(OpportunityGUID,ProjectHandoverGUID, service);
            }
            else if (Flag == 2) //2-Project
            {
                Guid ProjectHandoverGUID = ProjectHandover.Get<EntityReference>(context).Id;
                Guid ProjectGUID = Project.Get<EntityReference>(context).Id;
                executeProjectBusinessLogic(ProjectGUID, ProjectHandoverGUID, service);
            }

        }

        private void executeProjectHandoverBusinessLogic(Guid opportunityGUID, Guid projectHandoverGUID, IOrganizationService service)
        {
            try
            {
                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet("its_opportunityid");
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("its_opportunityid", ConditionOperator.Equal, opportunityGUID));
                fe.AddCondition(new ConditionExpression("its_opportunityid", ConditionOperator.NotNull));
                q1.Criteria = fe;
                q1.EntityName = "itspsa_projectdocuments";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if(ec.Entities.Count>0)
                {
                    foreach(Entity c in ec.Entities)
                    {
                        Guid ProjectDocumentGuid = new Guid(c.Attributes["itspsa_projectdocumentsid"].ToString());
                        Entity PD = new Entity("itspsa_projectdocuments");
                        PD["its_projecthandoverid"] = new EntityReference("itspsa_projecthandover", projectHandoverGUID);
                        PD.Id = ProjectDocumentGuid;
                        service.Update(PD);
                    }
                }
            }
            catch (InvalidPluginExecutionException ex)
            {
                if (ex != null)
                {

                    throw new InvalidPluginExecutionException("1) Unable to do the operation." + ex + "Contact Your Administrator");
                }
                else
                    throw new InvalidPluginExecutionException("2) Unable to do the operation." + ex);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("3) Unable to do the operation." + ex);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("4) Unable to do the operation." + ex);
            }
        }

        private void executeProjectBusinessLogic(Guid projectGUID, Guid projectHandoverGUID, IOrganizationService service)
        {
            try
            {
                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet("its_projecthandoverid");
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("its_projecthandoverid", ConditionOperator.Equal, projectHandoverGUID));
                fe.AddCondition(new ConditionExpression("its_projecthandoverid", ConditionOperator.NotNull));
                q1.Criteria = fe;
                q1.EntityName = "itspsa_projectdocuments";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {
                    foreach (Entity c in ec.Entities)
                    {
                        Guid ProjectDocumentGuid = new Guid(c.Attributes["itspsa_projectdocumentsid"].ToString());
                        Entity PD = new Entity("itspsa_projectdocuments");
                        PD["itspsa_project"] = new EntityReference("msdyn_project", projectGUID);
                        PD.Id = ProjectDocumentGuid;
                        service.Update(PD);
                    }
                }
            }
            catch (InvalidPluginExecutionException ex)
            {
                if (ex != null)
                {

                    throw new InvalidPluginExecutionException("1) Unable to do the operation." + ex + "Contact Your Administrator");
                }
                else
                    throw new InvalidPluginExecutionException("2) Unable to do the operation." + ex);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("3) Unable to do the operation." + ex);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("4) Unable to do the operation." + ex);
            }
        }
    }
}
