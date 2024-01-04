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

namespace UpdateTaskStatus
{
    public class TaskStatus:CodeActivity
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

        private void executeBusinessLogic(Guid projectTaskGUID, IOrganizationService service)
        {
            try
            {
                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet("msdyn_taskid", "msdyn_resourceassignmentid");
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("msdyn_taskid", ConditionOperator.Equal, projectTaskGUID));
                q1.Criteria = fe;
                q1.EntityName = "msdyn_resourceassignment";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {
                    foreach(Entity c in ec.Entities)
                    {
                        Guid RAGuid = new Guid(c.Attributes["msdyn_resourceassignmentid"].ToString());

                        Entity ProjectTask = service.Retrieve("msdyn_projecttask", projectTaskGUID, new ColumnSet("itspsa_taskstatus"));
                        int? TaskStatusValue = ((OptionSetValue)ProjectTask.Attributes["itspsa_taskstatus"]).Value;

                        Entity ResourceAssignment = new Entity("msdyn_resourceassignment");
                        if (TaskStatusValue == 100000002) //Completed
                        {
                            ResourceAssignment["its_taskstatus"] = new OptionSetValue(100000002);
                        }
                        else if(TaskStatusValue== 100000001 || TaskStatusValue== 100000000)//Pending or InProgress
                        {
                            ResourceAssignment["its_taskstatus"] = new OptionSetValue(960760000);
                        }
                        ResourceAssignment.Id = RAGuid;
                        service.Update(ResourceAssignment);
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
