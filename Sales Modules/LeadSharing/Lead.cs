using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Xrm;

namespace LeadSharing
{
    public class Lead:IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            Microsoft.Xrm.Sdk.IPluginExecutionContext context = (Microsoft.Xrm.Sdk.IPluginExecutionContext)
            serviceProvider.GetService(typeof(Microsoft.Xrm.Sdk.IPluginExecutionContext));

            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
            {

                // Entity Record To be shared
                EntityReference Lead = ((EntityReference)context.InputParameters["Target"]);

                // User or Team  for whom the record has been shared
                EntityReference sharedRecord = ((PrincipalAccess)context.InputParameters["PrincipalAccess"]).Principal;

                if (sharedRecord.LogicalName == "systemuser")
                {

                    // System User ID who has shared the record
                    Guid fromUserId = context.UserId;
                    // Send An EMail
                    EmailUser(service, sharedRecord, fromUserId, Lead);
                }

                //else if (sharedRecord.LogicalName == "team")
                //{

                //    //Creating XRM  Service Context
                //    XrmServiceContext datacontext = new XrmServiceContext(service);

                //    // Retrieving List of Users in the Team

                //    List teamMembers = (from t in datacontext.TeamMembershipSet where t.TeamId == sharedRecord.Id select t).ToList();

                //    TeamEmail(service, teamMembers, context.UserId, sharedRecord);

                //}
            }

        }

        private void EmailUser(IOrganizationService service, EntityReference sysUser, Guid fromUserId, EntityReference regarding)
        {
            Entity email = new Entity("email");
            email.Attributes.Add("subject", "Lead Has been Shared");
            string body = "";
            Entity User = service.Retrieve("systemuser", sysUser.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("fullname"));
            string Fullname = string.Empty;
            if(User.Attributes.Contains("fullname"))
            {
                Fullname = User.Attributes["fullname"].ToString();
            }
            body = "<div align='left' style='width:110px; font:12px Arial, Helvetica, sans-serif'>";
            body = body + "<div style='padding:10px'>";
            body = body + "Dear <b></b>Admin,<br /><br />";
            body = body + "Lead has been shared with User." + Fullname + "<br /><br />";
            body = body + "<b> Thank You.</b><br /><br /><br />";
            //body = body + "<a href=" + url + ">Click to open the Record</a>";
            body = body + "Thanks & Regards,<br /><br />";
            body = body + "CRM Admin<br /><br /></div>";

            email.Attributes.Add("description", body);
            email.Attributes.Add("regardingobjectid", regarding);

            Guid QueueId = new Guid("43786172-c2ac-ea11-a812-000d3ab19dd4");
            Guid AdminId = new Guid("B03FB559-25F5-E711-8111-5065F38BD371");
            EntityReference from = new EntityReference("queue", QueueId);
            EntityReference to = new EntityReference("systemuser", AdminId);

            Entity fromParty = new Entity("activityparty");
            fromParty.Attributes.Add("partyid", from);
            Entity toParty = new Entity("activityparty");
            toParty.Attributes.Add("partyid", to);

            EntityCollection frmPartyCln = new EntityCollection();
            frmPartyCln.EntityName = "queue";
            frmPartyCln.Entities.Add(fromParty);

            EntityCollection toPartyCln = new EntityCollection();
            toPartyCln.EntityName = "systemuser";
            toPartyCln.Entities.Add(toParty);

            email.Attributes.Add("from", frmPartyCln);
            email.Attributes.Add("to", toPartyCln);

            //Create an EMail Record
            Guid _emailId = service.Create(email);

            // Use the SendEmail message to send an e-mail message.
            SendEmailRequest sendEmailreq = new SendEmailRequest
            {
                EmailId = _emailId,
                TrackingToken = "",
                IssueSend = true
            };

            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailreq);
        }

//        private void TeamEmail(IOrganizationService service, List teamMembers, Guid fromUserId, EntityReference regarding)
//        {

//            Email email = new Email();

//            email.Subject = "New Account Has been Shared With You";

//            email.Description = "
//" + "Dear Team  " + "
//" + "Please Go through the Account shared with you " + "
//";

//            // Set the from user of the email to Logged in user
//            ActivityParty fromParty = new ActivityParty();
//            fromParty.PartyId = new EntityReference(SystemUser.EntityLogicalName, fromUserId);

//            List toPartyList = new List();

//            // Add the Activity party for each member in the team
//            foreach (TeamMembership user in teamMembers)
//            {
//                ActivityParty toParty = new ActivityParty();
//                toParty.PartyId = new EntityReference(SystemUser.EntityLogicalName, user.SystemUserId.Value);
//                toPartyList.Add(toParty);
//            }

//            // Add To and From users to email
//            email.To = toPartyList.ToArray();
//            //  email.To = toPartyList.ToArray();
//            email.From = new[] { fromParty };

//            // Set the Account regarding
//            email.RegardingObjectId = new EntityReference(Account.EntityLogicalName, regarding.Id);

//            // Email GUID
//            Guid emailId = service.Create(email);

//            // Use the SendEmail message to send an e-mail message.
//            SendEmailRequest sendEmailreq = new SendEmailRequest
//            {
//                EmailId = emailId,
//                TrackingToken = "",
//                IssueSend = true
//            };

//            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailreq);

//        }
    }
}
