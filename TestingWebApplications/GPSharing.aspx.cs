using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class GPSharing : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
        ClientCredentials credentials = new ClientCredentials();
        credentials.UserName.UserName = "itdev@alrostamanigroup.onmicrosoft.com";
        credentials.UserName.Password = "3edc#EDC2";
        Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

        Uri HomeRealUir = null;

        using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
        {
            IOrganizationService service = (IOrganizationService)serviceProxy;
            serviceProxy.EnableProxyTypes();

            Entity email = new Entity("email");
            email.Attributes.Add("subject", "New Lead Has been Shared With You");
            string body = "";
            //string Username = sysUser.Name;
            body = "<div align='left' style='width:110px; font:12px Arial, Helvetica, sans-serif'>";
            body = body + "<div style='padding:10px'>";
            body = body + "Dear <b></b>Admin,<br /><br />";
            //body = body + "Lead has been shared with User." + Username + "<br /><br />";
            body = body + "<b> Thank You.</b><br /><br /><br />";
            //body = body + "<a href=" + url + ">Click to open the Record</a>";
            body = body + "Thanks & Regards,<br /><br />";
            body = body + "CRM Admin<br /><br /></div>";

            email.Attributes.Add("description", body);
            //email.Attributes.Add("regardingobjectid", regarding);

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

            Guid ClientManagerGuid = new Guid("dddb52fb-a15c-e711-80fb-5065f38bd371");
            decimal EstGP = 15000;
            Guid OpportunityGuid = new Guid("8e9c2311-eaab-e711-8106-5065f38bb391");

            Entity Opportunity = service.Retrieve("opportunity", OpportunityGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet("parentaccountid"));
            Guid AccountGuid = new Guid();
            if (Opportunity.Attributes.Contains("parentaccountid"))
            {
                AccountGuid = ((EntityReference)Opportunity["parentaccountid"]).Id;
            }

            Entity GP1 = new Entity("arc_opportunitysharing");
            GP1["arc_clientmanager"] = new EntityReference("systemuser", ClientManagerGuid);
            GP1["arc_estgp"] = new Money(EstGP);
            GP1["arc_opportunity"] = new EntityReference("opportunity", OpportunityGuid);
            GP1["arc_gpsharing"] =Convert.ToDouble(75);
            decimal GP1Value = (EstGP * 75) / 100;
            GP1["arc_gpvalue"] = new Money(GP1Value);
            if(AccountGuid!=Guid.Empty)
            {
                GP1["arc_account"] = new EntityReference("account", AccountGuid);
            }
            Guid GP1Guid = service.Create(GP1);

            

            Entity GP2 = new Entity("arc_opportunitysharing");
            GP2["arc_clientmanager"] = new EntityReference("systemuser", ClientManagerGuid);
            GP2["arc_estgp"] = new Money(EstGP);
            GP2["arc_opportunity"] = new EntityReference("opportunity", OpportunityGuid);
            GP2["arc_gpsharing"] = Convert.ToDouble(25);
            decimal GP2Value = (EstGP * 25) / 100;
            GP2["arc_gpvalue"] = new Money(GP2Value);
            GP2["arc_remaininggp"] =Convert.ToDouble(0);
            if (AccountGuid != Guid.Empty)
            {
                GP2["arc_account"] = new EntityReference("account", AccountGuid);
            }
            Guid GP2Guid = service.Create(GP2);

        }
        }
    }