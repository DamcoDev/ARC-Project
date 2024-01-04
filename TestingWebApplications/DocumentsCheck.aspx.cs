using Microsoft.Crm.Sdk.Messages;
using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.ServiceModel.Description;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DocumentsCheck : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        ClientCredentials credentials = new ClientCredentials();
        //credentials.UserName.UserName = "itdev@alrostamanigroup.onmicrosoft.com";
        //credentials.UserName.Password = "3edc#EDC34";
        //Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        credentials.UserName.UserName = "admin@zaabeel.onmicrosoft.com";
        credentials.UserName.Password = "palace@123";
        Uri OrganizationUri = new Uri("https://zaabeeluat.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        Uri HomeRealUir = null;

        using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
        {
            IOrganizationService service = (IOrganizationService)serviceProxy;
            serviceProxy.EnableProxyTypes();

            Guid auditId = new Guid("544D118A-4F36-EC11-981F-00155DA72869");

            RetrieveAuditDetailsRequest request = new RetrieveAuditDetailsRequest();
            request.AuditId = auditId;

            RetrieveAuditDetailsResponse response =
                (RetrieveAuditDetailsResponse)service.Execute(request);
            service.Create(response.AuditDetail.AuditRecord);

            //Guid OpportunityGUID = new Guid("974be903-19be-e711-810d-5065f38b8251");
            //executeBusinessLogic(OpportunityGUID, service);
        }
     }
    private void executeBusinessLogic(Guid opportunityGUID, IOrganizationService service)
    {
        try
        {
            QueryExpression qe = new QueryExpression();
            qe.ColumnSet = new ColumnSet { AllColumns = true };
            qe.EntityName = "its_documentsconfiguration";
            EntityCollection ecqe = service.RetrieveMultiple(qe);
            int OpportunityClassicationValue = 0;
            if (ecqe.Entities.Count > 0)
            {
                foreach (Entity cqe in ecqe.Entities)
                {
                    Guid DocTypeGuid = ((EntityReference)cqe.Attributes["its_documenttype"]).Id;
                    string DocTypeName = ((EntityReference)cqe.Attributes["its_documenttype"]).Name;
                    string ClassificationValuesFromConfiguration = cqe.Attributes["its_requiredclassificationvalues"].ToString();
                    //New Code 23-May-2021
                    Entity Opportunity = service.Retrieve("opportunity", opportunityGUID, new ColumnSet("its_classificationnew"));
                    if (Opportunity.Attributes.Contains("its_classificationnew"))
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
        catch (Exception ex)
        {
            throw new InvalidPluginExecutionException("4) Unable to do the operation." + ex);
        }
    }
}