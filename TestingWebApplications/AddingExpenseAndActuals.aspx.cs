using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class AddingExpenseAndActuals : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        ClientCredentials credentials = new ClientCredentials();
        credentials.UserName.UserName = "itdev@alrostamanigroup.onmicrosoft.com";
        credentials.UserName.Password = "ARC@2019";
        Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        Uri HomeRealUir = null;
        Guid AccountGuid = new Guid();
        using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
        {
            IOrganizationService service = (IOrganizationService)serviceProxy;
            serviceProxy.EnableProxyTypes();

            Guid ProjectGuid = new Guid("ba6170fd-8395-e911-a966-000d3ab5a84e");

            Entity Project = new Entity("msdyn_project");

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

            totalcost_sum = String.Format(totalcost_sum, ProjectGuid);
            EntityCollection totalcost_sum_result = service.RetrieveMultiple(new FetchExpression(totalcost_sum));
            decimal TotalItemCost = 0;
            foreach (var c in totalcost_sum_result.Entities)
            {
                AliasedValue CostingValue = (AliasedValue)c["itspsa_totalcost_sum"];
                if (CostingValue.Value != null)
                {
                    TotalItemCost = ((Money)((AliasedValue)c["itspsa_totalcost_sum"]).Value).Value;
                }

                //Project["itspsa_actualadditionalcost"] = new Money(TotalItemCost);
                //Project.Id = ProjectGuid;
                //service.Update(Project);
            }

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

            totalcost_sum_Contractor = String.Format(totalcost_sum_Contractor, ProjectGuid);
            EntityCollection totalcost_sum_Contractor_result = service.RetrieveMultiple(new FetchExpression(totalcost_sum_Contractor));
            decimal TotalItemCostContractor = 0;
            foreach (var c in totalcost_sum_Contractor_result.Entities)
            {
                AliasedValue CostingValueContractor = (AliasedValue)c["itspsa_povalue_sum"];
                if (CostingValueContractor.Value != null)
                {
                    TotalItemCostContractor = ((Money)((AliasedValue)c["itspsa_povalue_sum"]).Value).Value;
                }

                //Project["itspsa_actualadditionalcost"] = new Money(TotalItemCostContractor);
                //Project.Id = ProjectGuid;
                //service.Update(Project);
            }

            decimal totalAdditionalCost = TotalItemCost + TotalItemCostContractor;

            Project["itspsa_actualadditionalcost"] = new Money(totalAdditionalCost);
            Project.Id = ProjectGuid;
            service.Update(Project);

            decimal additionalCost = 0;
            decimal laborCost = 0;
            decimal expenseCost = 0;

            Entity ProjectEntity = service.Retrieve("msdyn_project", ProjectGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet("itspsa_actualadditionalcost", "msdyn_actuallaborcost", "msdyn_actualexpensecost"));
            if (ProjectEntity.Attributes.Contains("msdyn_actuallaborcost"))
            {
                laborCost =((Money)ProjectEntity.Attributes["msdyn_actuallaborcost"]).Value;
            }
            if(ProjectEntity.Attributes.Contains("msdyn_actualexpensecost"))
            {
                expenseCost = ((Money)ProjectEntity.Attributes["msdyn_actualexpensecost"]).Value;
            }
            if(ProjectEntity.Attributes.Contains("itspsa_actualadditionalcost"))
            {
                additionalCost = ((Money)ProjectEntity.Attributes["itspsa_actualadditionalcost"]).Value;
            }

           
            //Project["itspsa_actualadditionalcost"] = new Money(100);
            decimal finalcost = laborCost + expenseCost + additionalCost;
            Project["msdyn_totalactualcost"] = new Money(finalcost);
            Project.Id = ProjectGuid;
            service.Update(Project);                
        }
    }
}