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

namespace DocumentValidation
{
    public class DocValidate:CodeActivity
    {
        [Input("Opportunity")]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1) { return; }
            Guid OpportunityGUID = Opportunity.Get<EntityReference>(context).Id;
            executeBusinessLogic(OpportunityGUID, service);
        }

        private void executeBusinessLogic(Guid opportunityGUID, IOrganizationService service)
        {
            //try
            //{


            QueryExpression qe = new QueryExpression();
            qe.ColumnSet = new ColumnSet { AllColumns = true };
            qe.EntityName = "its_documentsconfiguration";
            EntityCollection ecqe = service.RetrieveMultiple(qe);
            int OpportunityClassicationValue=0;
            if(ecqe.Entities.Count>0)
            {
                foreach(Entity cqe in ecqe.Entities)
                {
                    Guid DocTypeGuid = ((EntityReference)cqe.Attributes["its_documenttype"]).Id;
                    string DocTypeName = ((EntityReference)cqe.Attributes["its_documenttype"]).Name;
                    string ClassificationValuesFromConfiguration = cqe.Attributes["its_requiredclassificationvalues"].ToString();
                    //New Code 23-May-2021
                    Entity Opportunity = service.Retrieve("opportunity", opportunityGUID, new ColumnSet("its_classificationnew"));
                    if(Opportunity.Attributes.Contains("its_classificationnew"))
                    {
                        OpportunityClassicationValue = ((OptionSetValue)Opportunity.Attributes["its_classificationnew"]).Value;
                    }
                    if (ClassificationValuesFromConfiguration.Contains(OpportunityClassicationValue.ToString()))
                    {
                        QueryExpression q1 = new QueryExpression();
                        q1.ColumnSet = new ColumnSet("its_opportunityid", "its_documentattached");
                        FilterExpression fe = new FilterExpression(LogicalOperator.And);
                        fe.AddCondition(new ConditionExpression("its_opportunityid", ConditionOperator.Equal, opportunityGUID));
                        fe.AddCondition(new ConditionExpression("itspsa_documenttype", ConditionOperator.Equal, DocTypeGuid));
                        fe.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));
                        q1.Criteria = fe;
                        q1.EntityName = "itspsa_projectdocuments";
                        EntityCollection ec = service.RetrieveMultiple(q1);
                        if (ec.Entities.Count > 0)
                        {
                            foreach (Entity c in ec.Entities)
                            {
                                Guid ProjectDocumentGuid = new Guid(c.Attributes["itspsa_projectdocumentsid"].ToString());
                                if (c.Attributes.Contains("its_documentattached"))
                                {
                                    bool Document = Convert.ToBoolean(c.Attributes["its_documentattached"]);
                                    if (!Document)
                                    {
                                        throw new InvalidPluginExecutionException("Please do upload/attach the corresponding documents that got created in Documents Tab.!");
                                    }
                                }
                            }
                        }
                        else
                        {
                            throw new InvalidPluginExecutionException("Please do upload/attach " + DocTypeName + " document in Documents Tab to proceed further!");
                        }
                    }
                    else
                    {
                        //End 
                    }
                }
            }

               
            //}
            //catch (InvalidPluginExecutionException ex)
            //{
            //    if (ex != null)
            //    {

            //        throw new InvalidPluginExecutionException("1) Unable to do the operation." + ex + "Contact Your Administrator");
            //    }
            //    else
            //        throw new InvalidPluginExecutionException("2) Unable to do the operation." + ex);
            //}
            //catch (FaultException<OrganizationServiceFault> ex)
            //{
            //    throw new InvalidPluginExecutionException("3) Unable to do the operation." + ex);
            //}
            //catch (Exception ex)
            //{
            //    throw new InvalidPluginExecutionException("4) Unable to do the operation." + ex);
            //}
        }
    }
}
