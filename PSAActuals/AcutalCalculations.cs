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

namespace PSAActuals
{
    public class AcutalCalculations:CodeActivity
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

            if(workflowContext.Depth>1)
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
                Entity Project = new Entity("msdyn_project");

                #region Procurement Details
                string totalcost_sum = string.Empty;
                totalcost_sum = @" 
                                <fetch distinct='false' mapping='logical' aggregate='true'> 
                                    <entity name='itspsa_procurementdetails'> 
                                       <attribute name='itspsa_totalcost' alias='itspsa_totalcost_sum' aggregate='sum' /> 
                                    <filter type='and'>
                                      <condition attribute='itspsa_projectid' operator='eq' value='{0}'/>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    </entity> 
                                </fetch>";

                totalcost_sum = String.Format(totalcost_sum, projectGUID);
                EntityCollection totalcost_sum_result = service.RetrieveMultiple(new FetchExpression(totalcost_sum));
                decimal TotalItemCost = 0;
                foreach (var c in totalcost_sum_result.Entities)
                {
                    AliasedValue CostingValue = (AliasedValue)c["itspsa_totalcost_sum"];
                    if (CostingValue.Value != null)
                    {
                        TotalItemCost = ((Money)((AliasedValue)c["itspsa_totalcost_sum"]).Value).Value;
                    }
                }
                #endregion

                #region Sub Contractor Details
                string totalcost_sum_Contractor = string.Empty;
                totalcost_sum_Contractor = @" 
                                <fetch distinct='false' mapping='logical' aggregate='true'> 
                                    <entity name='itspsa_subcontractors'> 
                                       <attribute name='itspsa_povalue' alias='itspsa_povalue_sum' aggregate='sum' /> 
                                    <filter type='and'>
                                      <condition attribute='itspsa_projectid' operator='eq' value='{0}'/>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    </entity> 
                                </fetch>";

                totalcost_sum_Contractor = String.Format(totalcost_sum_Contractor, projectGUID);
                EntityCollection totalcost_sum_Contractor_result = service.RetrieveMultiple(new FetchExpression(totalcost_sum_Contractor));
                decimal TotalItemCostContractor = 0;
                foreach (var c in totalcost_sum_Contractor_result.Entities)
                {
                    AliasedValue CostingValueContractor = (AliasedValue)c["itspsa_povalue_sum"];
                    if (CostingValueContractor.Value != null)
                    {
                        TotalItemCostContractor = ((Money)((AliasedValue)c["itspsa_povalue_sum"]).Value).Value;
                    }
                }
                #endregion

                #region Change Request Details
                string totalcost_sum_CR = string.Empty;
                totalcost_sum_CR = @" 
                                <fetch distinct='false' mapping='logical' aggregate='true'> 
                                    <entity name='itspsa_changerequest'> 
                                       <attribute name='itspsa_costimpact' alias='itspsa_costimpact_sum' aggregate='sum' /> 
                                    <filter type='and'>
                                      <condition attribute='itspsa_project' operator='eq' value='{0}'/>
                                        <condition attribute='statecode' operator='eq' value='0' />
                                    </filter>
                                    </entity> 
                                </fetch>";

                totalcost_sum_CR = String.Format(totalcost_sum_CR, projectGUID);
                EntityCollection totalcost_sum_CR_result = service.RetrieveMultiple(new FetchExpression(totalcost_sum_CR));
                decimal TotalItemCostCR = 0;
                foreach (var c in totalcost_sum_CR_result.Entities)
                {
                    AliasedValue CostingValueContractor = (AliasedValue)c["itspsa_costimpact_sum"];
                    if (CostingValueContractor.Value != null)
                    {
                        TotalItemCostCR = ((Money)((AliasedValue)c["itspsa_costimpact_sum"]).Value).Value;
                    }
                }
                #endregion

                decimal totalAdditionalCost = TotalItemCost + TotalItemCostContractor+ TotalItemCostCR;

                Project["itspsa_actualadditionalcost"] = new Money(totalAdditionalCost);
                Project.Id = projectGUID;
                service.Update(Project);

                decimal additionalCost = 0;
                decimal laborCost = 0;
                decimal expenseCost = 0;

                Entity ProjectEntity = service.Retrieve("msdyn_project", projectGUID, new Microsoft.Xrm.Sdk.Query.ColumnSet("itspsa_actualadditionalcost", "msdyn_actuallaborcost", "msdyn_actualexpensecost"));
                if (ProjectEntity.Attributes.Contains("msdyn_actuallaborcost"))
                {
                    laborCost = ((Money)ProjectEntity.Attributes["msdyn_actuallaborcost"]).Value;
                }
                if (ProjectEntity.Attributes.Contains("msdyn_actualexpensecost"))
                {
                    expenseCost = ((Money)ProjectEntity.Attributes["msdyn_actualexpensecost"]).Value;
                }
                if (ProjectEntity.Attributes.Contains("itspsa_actualadditionalcost"))
                {
                    additionalCost = ((Money)ProjectEntity.Attributes["itspsa_actualadditionalcost"]).Value;
                }

                decimal finalcost = laborCost + expenseCost + additionalCost;
                Project["msdyn_totalactualcost"] = new Money(finalcost);
                Project.Id = projectGUID;
                service.Update(Project);
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
