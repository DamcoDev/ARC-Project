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
using Xrm;
using Microsoft.Crm.Sdk.Messages;

namespace SendEmailToTeam
{
    public class TeamTemplateEmail : CodeActivity
    {
        [Input("Case")]
        [RequiredArgument]
        [ReferenceTarget("incident")]
        public InArgument<EntityReference> Case { get; set; }

        [Input("CRMAdmin")]
        [RequiredArgument]
        [ReferenceTarget("queue")]
        public InArgument<EntityReference> CRMAdmin { get; set; }

        [Input("Team")]
        [RequiredArgument]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> Team { get; set; }

        [Input("Template GUID")]
        public InArgument<string> TemplateGuid { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            if (workflowContext.Depth > 1)
            {
                return;
            }
            Guid caseGUID = Case.Get<EntityReference>(context).Id;
            Guid adminGUID = CRMAdmin.Get<EntityReference>(context).Id;
            Guid teamGUID = Team.Get<EntityReference>(context).Id;
            Guid templateGUID =new Guid(TemplateGuid.Get(context));
            executeBusinessLogic(caseGUID,adminGUID,teamGUID,templateGUID, service);

        }
        private void executeBusinessLogic(Guid caseGUID, Guid adminGUID, Guid teamGUID, Guid templateGUID, IOrganizationService service)
        {
            try
            {
                XrmServiceContext con = new XrmServiceContext(service);
                //Getting Team Members based on Condition.
                List<TeamMembership> teamMembers = (from t in con.TeamMembershipSet
                                                    where t.TeamId == teamGUID
                                                    select t).ToList();

                Email email = new Email();
                ActivityParty fromParty = new ActivityParty();
                fromParty.PartyId = new EntityReference(Queue.EntityLogicalName, adminGUID);

                List<ActivityParty> PartyList = new List<ActivityParty>();
                foreach (TeamMembership user in teamMembers)
                {
                    ActivityParty toParty = new ActivityParty();
                    toParty.PartyId = new EntityReference(SystemUser.EntityLogicalName, user.SystemUserId.Value);
                    PartyList.Add(toParty);
                }
                email.From = new[] { fromParty };
                email.To = PartyList.ToArray();
                email.RegardingObjectId = new EntityReference("incident", caseGUID);
                string subjectName = "Subject";
                string reqNumber = "Request";

                email.Subject = subjectName;
                email.DirectionCode = true;
                // Create the request
                SendEmailFromTemplateRequest emailUsingTemplateReq = new SendEmailFromTemplateRequest
                {
                    Target = email,

                    // Use a built-in Email Template of type "contact".
                    TemplateId = templateGUID,

                    // The regarding Id is required, and must be of the same type as the Email Template.
                    RegardingId =caseGUID,
                    RegardingType = Incident.EntityLogicalName
                };

                SendEmailFromTemplateResponse emailUsingTemplateResp = (SendEmailFromTemplateResponse)service.Execute(emailUsingTemplateReq);

                // Verify that the e-mail has been created
                Guid _emailId = emailUsingTemplateResp.Id;
                if (!_emailId.Equals(Guid.Empty))
                {
                    Console.WriteLine("Successfully sent an e-mail message using the template.");
                }
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
