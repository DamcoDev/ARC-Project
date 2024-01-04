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

namespace ResourceAssignment
{
    public class RATask:CodeActivity
    {
        [Input("Project Task")]
        [RequiredArgument]
        [ReferenceTarget("msdyn_projecttask")]
        public InArgument<EntityReference> ProjectTask { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            if (workflowContext.Depth > 1)
            {
                return;
            }
            Guid ProjectTaskGUID = ProjectTask.Get<EntityReference>(context).Id;
            executeBusinessLogic(ProjectTaskGUID, service);
        }
        private void executeBusinessLogic(Guid ProjectTaskGUID, IOrganizationService service)
        {
            try
            {
                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet { AllColumns = true };
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("msdyn_taskid", ConditionOperator.Equal, ProjectTaskGUID));
                fe.AddCondition(new ConditionExpression("msdyn_taskid", ConditionOperator.NotNull));
                q1.Criteria = fe;
                q1.EntityName = "msdyn_resourceassignment";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {
                    foreach (Entity c in ec.Entities)
                    {
                        if (c.Attributes.Contains("msdyn_bookableresourceid"))
                        {
                            Guid BookableResourceId = ((EntityReference)c.Attributes["msdyn_bookableresourceid"]).Id;
                            string entityType = ((EntityReference)c.Attributes["msdyn_bookableresourceid"]).LogicalName;
                            Entity BookableResource = service.Retrieve("bookableresource", BookableResourceId, new ColumnSet("resourcetype", "userid"));

                            if (BookableResource.Attributes.Contains("resourcetype"))
                            {
                                int resourceType = ((OptionSetValue)BookableResource.Attributes["resourcetype"]).Value;
                                if (resourceType == 3)
                                {
                                    Guid SystemUserId = ((EntityReference)BookableResource.Attributes["userid"]).Id;
                                    string SystemUserFullName = ((EntityReference)BookableResource.Attributes["userid"]).Name;

                                    Entity ProjectTask = new Entity("msdyn_projecttask");
                                    ProjectTask["itspsa_resourceuser"] = new EntityReference("systemuser", SystemUserId);
                                    ProjectTask.Id = ProjectTaskGUID;
                                    service.Update(ProjectTask);
                                    return;
                                }

                            }
                        }
                    }
                }

            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException("Unable to complete the Operation.", ex.InnerException);
            }
        }
    }
}
