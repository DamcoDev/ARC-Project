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

namespace ProjectContractLine
{
    public class PCL:CodeActivity
    {
        [Input("Project Contract")]
        [RequiredArgument]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> ProjectContract { get; set; }

        [Input("Project")]
        [RequiredArgument]
        [ReferenceTarget("msdyn_project")]
        public InArgument<EntityReference> Project { get; set; }

        [Input("Amount")]
        [RequiredArgument]
        public InArgument<decimal> Amount { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1)
            //{
            //    return;
            //}
            Guid ProjectContractGUid = ProjectContract.Get<EntityReference>(context).Id;
            decimal ContractAmount = Amount.Get<decimal>(context);
            Guid ProjectGuid = Project.Get<EntityReference>(context).Id;
            executeBusinessLogic(ProjectContractGUid, ContractAmount, ProjectGuid, service);

        }

        private void executeBusinessLogic(Guid projectContractGUid, decimal contractAmount, Guid projectGuid, IOrganizationService service)
        {
            try
            {
                Entity ContractLine = new Entity("salesorderdetail");
                ContractLine["producttypecode"] = new OptionSetValue(5);
                ContractLine["salesorderid"] = new EntityReference("salesorder", projectContractGUid);
                ContractLine["productdescription"] = "Project Contract Line";
                ContractLine["msdyn_billingmethod"] = new OptionSetValue(192350000);
                ContractLine["msdyn_project"] = new EntityReference("msdyn_project", projectGuid);
                ContractLine["priceperunit"] = new Money(contractAmount);
                ContractLine["msdyn_budgetamount"] = new Money(contractAmount);
                ContractLine["msdyn_includetime"] = true;
                ContractLine["msdyn_includeexpense"] = true;
                ContractLine["msdyn_includefee"] = true;
                Guid ContractLineGuid = service.Create(ContractLine);
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
