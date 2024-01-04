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

public partial class PSAResourceAssignment : System.Web.UI.Page
{
    public string positionName;
    public Guid projectTeamGuid;
    public DateTime fromDate;
    public DateTime toDate;
    public decimal hours;
    protected void Page_Load(object sender, EventArgs e)
    {


        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        ClientCredentials credentials = new ClientCredentials();
        credentials.UserName.UserName = "itdev@alrostamanigroup.onmicrosoft.com";
        credentials.UserName.Password = "3edc#EDC5";
        Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        Uri HomeRealUir = null;
        Guid AccountGuid = new Guid();
        using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
        {
            IOrganizationService service = (IOrganizationService)serviceProxy;
            serviceProxy.EnableProxyTypes();

            Guid ProjectTaskGUID = new Guid("5fadca1e-1021-ec11-b6e6-000d3ade6ac6");

            QueryExpression q1 = new QueryExpression();
            q1.ColumnSet = new ColumnSet { AllColumns = true };
            FilterExpression fe = new FilterExpression(LogicalOperator.And);
            fe.AddCondition(new ConditionExpression("msdyn_taskid", ConditionOperator.Equal, ProjectTaskGUID));
            fe.AddCondition(new ConditionExpression("msdyn_taskid", ConditionOperator.NotNull));
            q1.Criteria = fe;
            q1.EntityName = "msdyn_resourceassignment";
            EntityCollection ec = service.RetrieveMultiple(q1);
            if (ec.Entities.Count > 0)
            {
                foreach (Entity c in ec.Entities)
                {
                    if (c.Attributes.Contains("msdyn_bookableresourceid"))
                    {
                        Guid BookableResourceId = ((EntityReference)c.Attributes["msdyn_bookableresourceid"]).Id;
                        string entityType = ((EntityReference)c.Attributes["msdyn_bookableresourceid"]).LogicalName;
                        Entity BookableResource = service.Retrieve("bookableresource", BookableResourceId, new ColumnSet("resourcetype", "userid"));

                        if (BookableResource.Attributes.Contains("resourcetype"))
                        {
                            int resourceType = ((OptionSetValue)BookableResource.Attributes["resourcetype"]).Value;
                            if (resourceType == 3)
                            {
                                Guid SystemUserId = ((EntityReference)BookableResource.Attributes["userid"]).Id;
                                string SystemUserFullName = ((EntityReference)BookableResource.Attributes["userid"]).Name;

                                Entity ProjectTask = new Entity("msdyn_projecttask");
                                ProjectTask["itspsa_resourceuser"] = new EntityReference("systemuser", SystemUserId);
                                ProjectTask.Id = ProjectTaskGUID;
                                service.Update(ProjectTask);
                                return;
                            }

                        }
                    }
                }
            }


            //WBS(service);


            //Guid ResourceAssignmentGuid = new Guid("AA5F62E3-9992-E911-A966-000D3AB5A84E");
            //Entity ResourceAssignment = service.Retrieve("msdyn_resourceassignment", ResourceAssignmentGuid, new ColumnSet { AllColumns = true });
            //if (ResourceAssignment.Attributes.Contains("msdyn_projectteamid"))
            //{
            //    positionName = ((EntityReference)ResourceAssignment.Attributes["msdyn_projectteamid"]).Name;
            //    projectTeamGuid= ((EntityReference)ResourceAssignment.Attributes["msdyn_projectteamid"]).Id;
            //    fromDate = Convert.ToDateTime(ResourceAssignment.Attributes["msdyn_fromdate"]);
            //    toDate = Convert.ToDateTime(ResourceAssignment.Attributes["msdyn_todate"]);
            //    hours = Convert.ToDecimal(ResourceAssignment.Attributes["msdyn_hours"]);

            //    QueryExpression q1 = new QueryExpression();
            //    q1.ColumnSet = new ColumnSet { AllColumns = true };
            //    FilterExpression fe = new FilterExpression();
            //    fe.AddCondition(new ConditionExpression("resourcetype", ConditionOperator.Equal, 3)); // 3- User
            //    q1.Criteria = fe;
            //    q1.EntityName = "bookableresource";
            //    EntityCollection ec = service.RetrieveMultiple(q1);
            //    if (ec.Entities.Count > 0)
            //    {
            //        foreach (Entity c in ec.Entities)
            //        {
            //            string userName = ((EntityReference)c.Attributes["userid"]).Name;
            //            if (userName == positionName)
            //            {
            //                Guid userGuid = new Guid(c.Attributes["bookableresourceid"].ToString());

            //                updateProjectTeam(projectTeamGuid,fromDate,toDate,hours,userGuid, service);
            //                updateResourceAssignment(ResourceAssignmentGuid, userGuid, service);
            //            }
            //        }
            //    }
            //}
            //Guid ProjectGuid = new Guid("2dcdcca4-6992-e911-a96c-000d3ab5a6ae");
            //var projectTasks = getProjectTasks(ProjectGuid,service);
            //foreach(Entity taskList in projectTasks.Entities)
            //{
            //    string projectTaskId = taskList["msdyn_projecttaskid"].ToString();
            //    string subject = taskList["msdyn_subject"].ToString();
            //}
        }
    }

    private void WBS(IOrganizationService service)
    {
        QueryExpression q1 = new QueryExpression();
        q1.ColumnSet = new ColumnSet { AllColumns = true };
        q1.EntityName = "msdyn_resourcerequirement";
    }

    private void updateResourceAssignment(Guid resourceAssignmentGuid, Guid userGuid, IOrganizationService service)
    {
        Entity ResourceAssignment = new Entity("msdyn_resourceassignment");
        ResourceAssignment["msdyn_bookableresourceid"] = new EntityReference("bookableresource", userGuid);
        ResourceAssignment.Id = resourceAssignmentGuid;
        service.Update(ResourceAssignment);
    }

    private void updateProjectTeam(Guid projectTeamGuid, DateTime fromDate,DateTime toDate,decimal hours, Guid userGuid, IOrganizationService service)
    {
        Entity ProjectTeam = new Entity("msdyn_projectteam");
        ProjectTeam["msdyn_bookableresourceid"] = new EntityReference("bookableresource", userGuid);
        ProjectTeam["msdyn_from"] = fromDate;
        ProjectTeam["msdyn_to"] = toDate;
        ProjectTeam["msdyn_hardbookedhours"] = hours;
        ProjectTeam.Id = projectTeamGuid;
        service.Update(ProjectTeam);
    }

    private EntityCollection getProjectTasks(Guid projectGuid, IOrganizationService service)
    {
        QueryExpression q1 = new QueryExpression();
        q1.ColumnSet = new ColumnSet { AllColumns = true };
        FilterExpression fe = new FilterExpression();
        fe.AddCondition(new ConditionExpression("msdyn_project", ConditionOperator.Equal, projectGuid));
        q1.Criteria = fe;
        q1.EntityName = "msdyn_projecttask";
        EntityCollection ec = service.RetrieveMultiple(q1);
        return ec;
    }
}