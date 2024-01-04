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

namespace CustomerCaseCreation
{
    public class CustomerCase:CodeActivity
    {
        [Input("Case")]
        [RequiredArgument]
        [ReferenceTarget("incident")]
        public InArgument<EntityReference> Case { get; set; }

        [Output("Result")]
        public OutArgument<int> Result { get; set; }

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
            executeBusinessLogic(CaseGuid, service,context);
        }

        private void executeBusinessLogic(Guid CaseGuid, IOrganizationService service, CodeActivityContext context)
        {
            try
            {
                int CustomerResult = 0;
                Guid CustomerGuid = new Guid();
                Guid EngineerGuid = new Guid();
                Guid HelpDeskUserGuid = new Guid();

                Entity Case = service.Retrieve("incident", CaseGuid, new ColumnSet { AllColumns = true });

                Guid CaseCustomerGuid = ((EntityReference)Case.Attributes["customerid"]).Id;

                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet { AllColumns = true };
                FilterExpression fe = new FilterExpression();
                fe.AddCondition(new ConditionExpression("itscs_customer", ConditionOperator.Equal, CaseCustomerGuid));
                q1.Criteria = fe;
                q1.EntityName = "itscs_casecreationforspecificcustomers";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {
                    foreach (Entity c in ec.Entities)
                    {
                        if (c.Attributes.Contains("itscs_customer"))
                        {
                            CustomerGuid = ((EntityReference)c.Attributes["itscs_customer"]).Id;
                            EngineerGuid = ((EntityReference)c.Attributes["itscs_engineer"]).Id;
                            HelpDeskUserGuid = ((EntityReference)c.Attributes["itscs_helpdeskuser"]).Id;

                          
                            DateTime currentTime = DateTime.Now.ToLocalTime();
                            Entity CaseEntity = new Entity("incident");
                            CaseEntity["itscs_caseownerorengineer"] = new EntityReference("systemuser", EngineerGuid);
                            CaseEntity["itscs_helpdeskuser"] = new EntityReference("systemuser", HelpDeskUserGuid);
                            CaseEntity["itscs_timestampofassignmenttoengineer"] = currentTime.ToUniversalTime();                         
                            CaseEntity.Id = CaseGuid;
                            service.Update(CaseEntity);

                            AssignRequest assign = new AssignRequest
                            {
                                Assignee = new EntityReference("systemuser",
                               EngineerGuid),
                                Target = new EntityReference("incident",
                               CaseGuid)
                            };


                            // Execute the Request
                            service.Execute(assign);

                            CustomerResult = 1;
                            Result.Set(context, CustomerResult);
                            return;
                        }
                    }
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
