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

public partial class AutomaticAssignmentForCustomers : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        ClientCredentials credentials = new ClientCredentials();
        credentials.UserName.UserName = "itdev@alrostamanigroup.onmicrosoft.com";
        credentials.UserName.Password = "ARC@2019";
        Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        Uri HomeRealUir = null;
        Guid CustomerGuid = new Guid();
        Guid EngineerGuid = new Guid();
        Guid HelpDeskUserGuid = new Guid();
        using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
        {
            IOrganizationService service = (IOrganizationService)serviceProxy;
            serviceProxy.EnableProxyTypes();

            Guid CaseCustomerGuid = new Guid("A36A7F82-B3A1-E711-810E-5065F38AA961");
            Guid CaseGuid = new Guid("6605BC48-A9A7-E911-A965-000D3AB5A57D");

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
                        int CustomerResult = 1;
                        CaseEntity["description"] = CustomerResult.ToString();
                        CaseEntity.Id = CaseGuid;
                        service.Update(CaseEntity);
                       // Result.Set(context, true);
                        return;
                    }
                }
            
        }
        }
    }
}