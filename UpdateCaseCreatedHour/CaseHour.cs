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

namespace UpdateCaseCreatedHour
{
    public class CaseHour:CodeActivity
    {
        [Input("Case")]
        [RequiredArgument]
        [ReferenceTarget("incident")]
        public InArgument<EntityReference> Case { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1)
            //{
            //    return;
            //}
            Guid CaseGuid = Case.Get<EntityReference>(context).Id;
            executeBusinessLogic(CaseGuid, service, context);
        }
        private void executeBusinessLogic(Guid CaseGuid, IOrganizationService service, CodeActivityContext context)
        {
            try
            {
                Entity Incident = service.Retrieve("incident", CaseGuid, new ColumnSet("createdon", "itscs_escalatedtoengineeron"));

                DateTime CreatedOnDate = Convert.ToDateTime(Incident.Attributes["createdon"]).ToLocalTime().AddHours(4);

                string CaseCreatedHour = CreatedOnDate.ToString("HH");

                Entity Case = new Entity("incident");
                Case["its_casecreatedhour"] = CaseCreatedHour;


                if (Incident.Attributes.Contains("itscs_escalatedtoengineeron"))
                {
                    DateTime escalatedtoengineeron = Convert.ToDateTime(Incident.Attributes["itscs_escalatedtoengineeron"]).ToLocalTime().AddHours(4);

                    string escalatedtoengineeronHour = escalatedtoengineeron.ToString("HH");

                    Case["its_escalatedtoengineeronhour"] = escalatedtoengineeronHour;
                }
                Case.Id = CaseGuid;
                service.Update(Case);
                return;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }
}
