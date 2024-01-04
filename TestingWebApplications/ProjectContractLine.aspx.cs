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

public partial class ProjectContractLine : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        ClientCredentials credentials = new ClientCredentials();
        credentials.UserName.UserName = "itdev@alrostamanigroup.onmicrosoft.com";
        credentials.UserName.Password = "3edc#EDC511";
        Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        Uri HomeRealUir = null;

        using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
        {
            IOrganizationService service = (IOrganizationService)serviceProxy;
            serviceProxy.EnableProxyTypes();

            Entity Opportunity = service.Retrieve("opportunity", new Guid("4c1d7708-69ae-e711-810a-5065f38b8251"), new ColumnSet("estimatedclosedate"));
            if (Opportunity.Attributes.Contains("estimatedclosedate"))
            {
                DateTime dt = Convert.ToDateTime(Opportunity.Attributes["estimatedclosedate"].ToString());

                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet("regardingobjectid");
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("regardingobjectid", ConditionOperator.Equal, new Guid("4c1d7708-69ae-e711-810a-5065f38b8251")));
                q1.Criteria = fe;
                q1.EntityName = "its_presalesrequest";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {
                    foreach (Entity c in ec.Entities)
                    {
                        Guid PreSalesRequestGuid = new Guid(c.Attributes["activityid"].ToString());

                        Entity PreSalesRequest = new Entity("its_presalesrequest");
                        PreSalesRequest["its_opportunityestclosedate"] = dt;
                        PreSalesRequest.Id = PreSalesRequestGuid;
                        service.Update(PreSalesRequest);
                    }
                }
            }


            //Guid ProjectContractGUid = new Guid("78d7a171-2df8-ea11-a815-000d3ab19dd4");
            //Guid ProjectGuid = new Guid("f0ffc3cc-18cf-ea11-a812-000d3ab82fca");
            //decimal ContractAmount = 100;

            //Entity ContractLine = new Entity("salesorderdetail");
            //ContractLine["producttypecode"] = new OptionSetValue(5);
            //ContractLine["salesorderid"] = new EntityReference("salesorder", ProjectContractGUid);
            //ContractLine["productdescription"] = "Project Contract Line";
            //ContractLine["msdyn_billingmethod"] = new OptionSetValue(192350000);
            //ContractLine["msdyn_project"] = new EntityReference("msdyn_project", ProjectGuid);
            //ContractLine["priceperunit"] = new Money(ContractAmount);
            //ContractLine["msdyn_budgetamount"] = new Money(ContractAmount);
            //ContractLine["msdyn_includetime"] = true;
            //ContractLine["msdyn_includeexpense"] = true;
            //ContractLine["msdyn_includefee"] = true;
            //Guid ContractLineGuid = service.Create(ContractLine);
        }
    }
}