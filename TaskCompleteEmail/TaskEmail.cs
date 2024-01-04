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
using Xrm;

namespace TaskCompleteEmail
{
    public class TaskEmail:CodeActivity
    {
        [Input("Project Task")]
        [RequiredArgument]
        [ReferenceTarget("msdyn_projecttask")]
        public InArgument<EntityReference> ProjectTask { get; set; }

        [Input("Project")]
        [RequiredArgument]
        [ReferenceTarget("msdyn_project")]
        public InArgument<EntityReference> Project { get; set; }

        [Input("Email Admin")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> EmailAdmin { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            if (workflowContext.Depth > 1)
            {
                return;
            }
            Guid ProjectTaskGUID = ProjectTask.Get<EntityReference>(context).Id;
            Guid adminGUID = EmailAdmin.Get<EntityReference>(context).Id;
            Guid ProjectGUID = Project.Get<EntityReference>(context).Id;
            executeBusinessLogic(ProjectTaskGUID, ProjectGUID, adminGUID,service);

        }

        private void executeBusinessLogic(Guid projectTaskGUID,Guid ProjectGUID,Guid adminGUID, IOrganizationService service)
        {
            try
            {
                Entity Project = service.Retrieve("msdyn_project", ProjectGUID, new ColumnSet("msdyn_subject", "msdyn_customer", "msdyn_projectmanager", "itspsa_teamlead", "itspsa_projectnumber"));
                string subjectName = "Project Task Completed -> Project - " + Project.Attributes["msdyn_subject"].ToString();
                string PMName = ((EntityReference)Project.Attributes["msdyn_projectmanager"]).Name;
                string TLName = ((EntityReference)Project.Attributes["itspsa_teamlead"]).Name;

                Guid PMGuid = ((EntityReference)Project.Attributes["msdyn_projectmanager"]).Id;
                Guid TLGuid = ((EntityReference)Project.Attributes["itspsa_teamlead"]).Id;

                Entity ProjectTask = service.Retrieve("msdyn_projecttask", projectTaskGUID, new ColumnSet("msdyn_subject"));
                string ProjectTaskName = ProjectTask.Attributes["msdyn_subject"].ToString();

                List<Guid> resources = getResourceGUID(projectTaskGUID, service);

                Email email = new Email();
                ActivityParty fromParty = new ActivityParty();
                fromParty.PartyId = new EntityReference(SystemUser.EntityLogicalName, adminGUID);

                List<ActivityParty> ToPartyList = new List<ActivityParty>();

                ActivityParty toPartyPM = new ActivityParty();
                toPartyPM.PartyId = new EntityReference(SystemUser.EntityLogicalName, PMGuid);
                ToPartyList.Add(toPartyPM);

                ActivityParty toPartyTL = new ActivityParty();
                toPartyTL.PartyId = new EntityReference(SystemUser.EntityLogicalName, TLGuid);
                ToPartyList.Add(toPartyTL);


                List<ActivityParty> CCPartyList = new List<ActivityParty>();
                List<string> resourceNames = new List<string>();
                foreach (var guid in resources)
                {
                    ActivityParty ccParty = new ActivityParty();
                    ccParty.PartyId = new EntityReference(SystemUser.EntityLogicalName, guid);
                    CCPartyList.Add(ccParty);
                    //Get Names
                    Entity Resource = service.Retrieve("systemuser", guid, new ColumnSet("fullname"));
                    resourceNames.Add(Resource.Attributes["fullname"].ToString());
                }
                string rNames = string.Join(";", resourceNames.ToArray());

                email.From = new[] { fromParty };
                email.To = ToPartyList.ToArray();
                email.Cc = CCPartyList.ToArray();
                email.RegardingObjectId = new EntityReference("msdyn_project", ProjectGUID);

                email.Subject = subjectName;
                string body = "";
                body = "<div align='left' style='width:110px; font:12px Arial, Helvetica, sans-serif'>";
                body = body + "<div style='padding:10px'>";
                body = body + "Dear <b>" + PMName + "</b>,<br /><br />";
                body = body + "The below Project Task has been marked as completed." + "<br /><br />";
                body = body + "<b> Project: </b> " + Project.Attributes["msdyn_subject"].ToString() + "<br /><br />";
                body = body + "<b> Project Number: </b> " + Project.Attributes["itspsa_projectnumber"].ToString() + "<br /><br />";
                body = body + "<b> Project Task: </b> " + ProjectTaskName + "<br /><br />";
                body = body + "<b> Resource/s: </b> " + rNames + "<br /><br />";
                body = body + "<b> Customer Name: </b> " + ((EntityReference)Project.Attributes["msdyn_customer"]).Name + " <br /><br />";
                body = body + "<b> Thank You.</b><br /><br /><br />";
                body = body + "Thanks & Regards,<br /><br />";
                body = body + "ARC Admin<br /><br /></div>";
                email.Description = body;

                Guid EmailId = service.Create(email);

                SendEmailRequest req = new SendEmailRequest();
                req.EmailId = EmailId;
                req.IssueSend = true;
                req.TrackingToken = "";
                SendEmailResponse res = (SendEmailResponse)service.Execute(req);

            }
            catch(Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
        private static List<Guid> getResourceGUID(Guid projectTaskGuid, IOrganizationService service)
        {
            List<Guid> resourceGuid = new List<Guid>();
            QueryExpression q1 = new QueryExpression();
            q1.ColumnSet = new ColumnSet { AllColumns = true };
            FilterExpression fe = new FilterExpression(LogicalOperator.And);
            fe.AddCondition(new ConditionExpression("msdyn_taskid", ConditionOperator.Equal, projectTaskGuid));
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
                                resourceGuid.Add(((EntityReference)BookableResource.Attributes["userid"]).Id);
                            }
                        }
                    }
                }
            }
            return resourceGuid;
        }
    }
}
