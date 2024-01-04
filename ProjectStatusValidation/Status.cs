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

namespace ProjectStatusValidation
{
    public class Status:CodeActivity
    {
        [Input("Project")]
        [RequiredArgument]
        [ReferenceTarget("msdyn_project")]
        public InArgument<EntityReference> Project { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            if (workflowContext.Depth > 1)
            {
                return;
            }
            Guid ProjectGUID = Project.Get<EntityReference>(context).Id;
            executeBusinessLogic(ProjectGUID, service);

        }

        private void executeBusinessLogic(Guid projectGUID, IOrganizationService service)
        {
            try
            {
                string fetchXML = string.Empty;
                fetchXML = @"<fetch mapping='logical' distinct='false'>
     <entity name='msdyn_resourceassignment'>
        <attribute name = 'msdyn_resourceassignmentid'/>
              <filter type='and'>
                 <condition attribute='msdyn_projectid' operator='eq' uiname='Commvault BackupforAzure' uitype='msdyn_project' value='{0}'/>
                        </filter>
                       <link-entity name ='msdyn_projecttask' from='msdyn_projecttaskid' to='msdyn_taskid' link-type='inner' alias='ab'>
                                      <filter type='and'>
                                         <condition attribute='itspsa_pmstatus' operator='ne' value='110920000'/>
                                           </filter>
                                         </link-entity>
                                       </entity>
                                     </fetch>";
                fetchXML = String.Format(fetchXML, projectGUID);
                EntityCollection ec = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (ec.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Please clomplete all the Project Tasks to finish this Project.");
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }
}
