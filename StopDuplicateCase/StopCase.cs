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

namespace StopDuplicateCase
{
    public class StopCase:CodeActivity
    {
        [Input("Case Title")]
        [RequiredArgument]
        public InArgument<string> CaseTitle { get; set; }

        //[Input("Case Title")]
        //[RequiredArgument]
        //[ReferenceTarget("incident")]
        //public InArgument<EntityReference> CaseTitle { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            if (workflowContext.Depth > 1)
            {
                return;
            }
            //Guid CaseTitleText = CaseTitle.Get<EntityReference>(context).Id;
            string CaseTitleText = CaseTitle.Get<string>(context);
            executeBusinessLogic(CaseTitleText, service);
        }

        private void executeBusinessLogic(string CaseTitleText, IOrganizationService service)
        {
            try
            {
                ///Entity Case = service.Retrieve("incident", CaseTitleText, new ColumnSet("title"));
                //string cTitle = Case.Attributes["title"].ToString();

                //QueryExpression q1 = new QueryExpression();
                //q1.ColumnSet = new ColumnSet("title");
                //FilterExpression fe = new FilterExpression(LogicalOperator.And);
                //fe.AddCondition(new ConditionExpression("title", ConditionOperator.Equal, cTitle));
                //q1.Criteria = fe;
                //q1.EntityName = "incident";
                //EntityCollection ec = service.RetrieveMultiple(q1);
                //if (ec.Entities.Count > 0)
                //{
                //    throw new InvalidPluginExecutionException("Duplicate Case Found with same Case Title");
                //}

                Entity contactRecord = new Entity("incident");

                contactRecord.Attributes["title"] = CaseTitleText;
                contactRecord.Attributes["customerid"] = new EntityReference("account", new Guid("e1647f82-b3a1-e711-810e-5065f38aa961"));

                var request = new RetrieveDuplicatesRequest

                {
                    //Entity Object to be searched with the values filled for the attributes to check
                    BusinessEntity = contactRecord,

                    //Logical Name of the Entity to check Matching Entity
                    MatchingEntityName = contactRecord.LogicalName,


                };

                var response = (RetrieveDuplicatesResponse)service.Execute(request);
                if (response.DuplicateCollection.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Duplicate Case Found.");
                }
            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
    }
}
