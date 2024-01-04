using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using Xrm;
using Microsoft.Crm.Sdk.Messages;

namespace TeamEmailTest
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;
            ClientCredentials credentials = new ClientCredentials();
            credentials.UserName.UserName = "admin@alrostamanigroup.onmicrosoft.com";
            credentials.UserName.Password = "FsUt#qex4o9p";
            Uri OrganizationUri = new Uri("https://arcdev.api.crm4.dynamics.com/XRMServices/2011/Organization.svc");

            Uri HomeRealUir = null;
            Guid AccountGuid = new Guid();
            using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealUir, credentials, null))
            {
                IOrganizationService service = (IOrganizationService)serviceProxy;
                //serviceProxy.EnableProxyTypes();
                Guid caseGUID = new Guid("B4260446-0ADE-46E6-BCDA-ADF5C5AA2232");
                Guid adminGUID = new Guid("43786172-C2AC-EA11-A812-000D3AB19DD4");
                Guid teamGUID = new Guid("D8FC38F0-4980-E911-A964-000D3AB5A84E");
                Guid templateGUID = new Guid("A0FF989A-FDA5-E911-A966-000D3AB5A84E");
                executeBusinessLogic(caseGUID, adminGUID, teamGUID, templateGUID, service);
            }
        }
        static void executeBusinessLogic(Guid caseGUID, Guid adminGUID, Guid teamGUID, Guid templateGUID, IOrganizationService service)
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
                    RegardingId = caseGUID,
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
            //catch (FaultException<OrganizationServiceFault> ex)
            //{
            //    throw new InvalidPluginExecutionException("3) Unable to complete the Operation.", ex);
            //}
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("4) Unable to complete the Operation.", ex);
            }
        }
    }
}
