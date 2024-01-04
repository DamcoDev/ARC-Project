using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using System;

namespace ARC.CustomActivity.ChangeProgress
{
    public class ChangeProgress :CodeActivity
    {

        [Input("Project")]
        [ReferenceTarget("msdyn_project")]
        public InArgument<EntityReference> Project { get; set; }

        [Output("Progress")]
        public OutArgument<Decimal> ProgressValue { get; set; }
        

        protected override void Execute(CodeActivityContext context)
        {
            var projectRef = Project.Get<EntityReference>(context);
            var projectId = projectRef.Id;
            
            

            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.UserId);
            Entity entity = service.Retrieve(projectRef.LogicalName, projectRef.Id, new ColumnSet(true));

            decimal progressFinal = 0;


            var fetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                     <entity name='msdyn_projecttask'>
                       <attribute name='msdyn_projecttaskid' />
                       <filter type='and'>
                         <condition attribute='msdyn_project' operator='eq' uitype='msdyn_project' value='{projectId}' />
                       </filter>
                     </entity>
                   </fetch>";

            var projectTasks = service.RetrieveMultiple(new FetchExpression(fetch));

            var totalProjectTasks = projectTasks.Entities.Count;


            if (totalProjectTasks == 0)
            {
                return;
            }

            var completedTasks = 0;
            var progress = 0;

            foreach (var projectTask in projectTasks.Entities)
            {
                var timeEntryFetch = $@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                   <entity name='msdyn_timeentry'>
                                     <attribute name='msdyn_timeentryid'/>
                                     <attribute name='msdyn_entrystatus'/>
                                     <filter type='and'>
                                       <condition attribute='msdyn_projecttask' operator='eq' uitype='msdyn_projecttask' value='{projectTask.Id}'/>
                                     </filter>
                                   </entity>
                                 </fetch>";

                var timeEntries = service.RetrieveMultiple(new FetchExpression(timeEntryFetch));

                int totalTasks = timeEntries.Entities.Count;

                if (totalTasks != 0)
                {

                    foreach (var timeEntry in timeEntries.Entities)
                    {
                        if (((OptionSetValue)timeEntry["msdyn_entrystatus"]).Value == 192350002 || ((OptionSetValue)timeEntry["msdyn_entrystatus"]).Value == 192350003)
                        {
                            completedTasks++;
                        }
                    
                    
                    }
                
                    progress += (int)((completedTasks / (decimal)totalTasks) * 100);
                    progressFinal += progress / totalProjectTasks;
                    
                }
                completedTasks = 0;
                progress = 0;
            }


            this.ProgressValue.Set(context, progressFinal);


        }

    }
}