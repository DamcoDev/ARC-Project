using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System.Xml.Linq;
using ExportToExcel.Models;

namespace ExportToExcel
{
    public class SendExcelFileInEmail : IPlugin
    {
        public StringBuilder Trace { get; set; }

        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            //IOrganizationService crmService = serviceFactory.CreateOrganizationService(context.UserId);
            IOrganizationService crmService = serviceFactory.CreateOrganizationService(null);

            Trace = new StringBuilder();
            try
            {
                EntityReference entRef = null;
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    //Trace.AppendLine("Target found....");
                    entRef = (EntityReference)context.InputParameters["Target"];
                }

                if (entRef != null && entRef.Id != Guid.Empty)
                {
                    string customerId = entRef.Id.ToString();//context.InputParameters["CustomerId"].ToString();
                    List<Entity> customerContacts = null;
                    customerContacts = GetCustomerContacts(customerId, crmService);

                    string emailTemplateId = "37318889-1b50-ec11-8c62-000d3aba9b1b";
                    EmailTemplate emailTemplate = GetEmailTemplate(emailTemplateId, crmService);
                    if (emailTemplate != null)
                    {
                        bool areEmailsSent = SendEmailToCustomers(customerId, customerContacts, context.UserId, emailTemplate, crmService);
                        Trace.AppendLine("areEmailsSent = " + areEmailsSent.ToString());
                        Trace.AppendLine("emailTemplate != null");
                        context.OutputParameters["Message"] = "Success";
                    }
                    else
                    {
                        Trace.AppendLine("Email template not found");
                        context.OutputParameters["Message"] = "Failure";
                    }
                }

                context.OutputParameters["PluginTrace"] = Trace.ToString();
            }
            catch (Exception ex)
            {
                //throw ex.InnerException;
                context.OutputParameters["PluginTrace"] = Trace.ToString() + Environment.NewLine + "Exception Message = " + ex.Message + Environment.NewLine + "Exception Message = " + ex.InnerException;
                context.OutputParameters["Message"] = "Error";
            }
        }

        public EmailTemplate GetEmailTemplate(string emailTemplateId, IOrganizationService crmService)
        {
            //Outage Alert Email Notification Template
            EmailTemplate emailTemplate = null;
            string fetchXmlString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                          <entity name='template'>
                                            <attribute name='subject' />
                                            <attribute name='body' />
                                            <filter type='and'>
                                              <condition attribute='templateid' operator='eq' uitype='template' value='{" + emailTemplateId + @"}' />
                                            </filter>
                                          </entity>
                                        </fetch>";
            EntityCollection entColl = crmService.RetrieveMultiple(new FetchExpression(fetchXmlString));
            List<Entity> entitiesList = GetEntitiesList(entColl);
            if (entitiesList != null && entitiesList.Count > 0)
            {
                emailTemplate = new EmailTemplate();
                emailTemplate.EmailTemplateId = entitiesList[0].Id;
                emailTemplate.Subject = GetDataFromXml(entitiesList[0]["subject"].ToString());
                emailTemplate.Body = GetDataFromXml(entitiesList[0]["body"].ToString());
            }

            return emailTemplate;
        }

        private string GetDataFromXml(string value)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }
            XDocument document = XDocument.Parse(value);
            XNode node = document.Descendants().Elements().Last().LastNode;
            string nodeText = node.ToString().Replace("<![CDATA[", string.Empty).Replace("]]>", string.Empty);
            return nodeText;
        }

        public List<Entity> GetCustomerContacts(string customerId, IOrganizationService crmService)
        {
            List<Entity> customerContacts = new List<Entity>();
            string fetchXmlString = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                              <entity name='contact'>
                                                <attribute name='fullname' />
                                                <filter type='and'>
                                                  <condition attribute='parentcustomerid' operator='eq' uitype='account' value='{" + customerId + @"}' />
                                                </filter>
                                              </entity>
                                            </fetch>";

            EntityCollection entColl = crmService.RetrieveMultiple(new FetchExpression(fetchXmlString));
            List<Entity> entitiesList = GetEntitiesList(entColl);
            if (entitiesList.Count > 0)
            {
                customerContacts.AddRange(entitiesList);
                Trace.AppendLine("2.GetCustomerContacts..." + entitiesList.Count.ToString() + "...");
            }
            return customerContacts;
        }

        public List<Entity> GetEntitiesList(EntityCollection entColl)
        {
            List<Entity> entities = new List<Entity>();
            if (entColl != null && entColl.Entities != null && entColl.Entities.Count > 0)
            {
                entities.AddRange(entColl.Entities);
            }
            return entities;
        }

        public Entity[] GetToListOfCustomerContacts(List<Entity> customerContacts)
        {
            EntityCollection to = new EntityCollection();
            for (int j = 0; j < customerContacts.Count; j++)
            {
                Entity contact = customerContacts[j];
                if (contact.Id != Guid.Empty)
                {
                    //EntityReference toEntRefParty = new EntityReference("contact", authorizedContact.Id);
                    Entity toParty = new Entity("activityparty");
                    toParty["partyid"] = new EntityReference("contact", contact.Id);
                    to.Entities.Add(toParty);
                }
            }

            Entity[] arrayTo = to.Entities.ToArray();

            return arrayTo;
        }

        public bool SendEmailToCustomers(string customerId, List<Entity> customerContacts, Guid loggedInUserId, EmailTemplate emailTemplate, IOrganizationService crmService)
        {
            bool areEmailsSent = false;

            if (emailTemplate == null || emailTemplate.EmailTemplateId == Guid.Empty || string.IsNullOrEmpty(emailTemplate.Body))
            {
                Trace.AppendLine("4.EmailTemplate is missing in CRM or body text is not there in email template.");
            }
            else
            {
                EntityReference entRefCustomer = new EntityReference("account", new Guid(customerId));
                Trace.AppendLine("4.EmailTemplate found.");
                if (customerContacts.Count > 0)
                {

                    Trace.AppendLine("4.customers[i].OutageAlert is found");
                    //--Get customer's accounts in byte array for email attachment
                    //byte[] fileStream = GetCustomersAccountsStream(crmService, loggedInUserId, customers[i]);
                    //Trace.AppendLine("5.fileStream..." + fileStream.Length.ToString() + "...");

                    Entity fromPartyGuid = new Entity("activityparty");
                    fromPartyGuid["partyid"] = new EntityReference("queue", new Guid("43786172-c2ac-ea11-a812-000d3ab19dd4"));

                    Entity[] fromParty = new[] { fromPartyGuid };
                    Entity[] to = GetToListOfCustomerContacts(customerContacts);

                    if (to.Length > 0)
                    {
                        Entity entEmail = new Entity("email");
                        entEmail["subject"] = string.Format(emailTemplate.Subject);
                        var localDate = DateTime.UtcNow.ToLocalTime();
                        var localDateTime = DateTime.SpecifyKind(localDate, DateTimeKind.Unspecified);
                        string fromDate = ConvertCrmDateToUserDate(crmService, loggedInUserId, localDateTime).ToString();
                        string toDate = ConvertCrmDateToUserDate(crmService, loggedInUserId, localDateTime.AddDays(2)).ToString();
                        entEmail["description"] = emailTemplate.Body;
                        entEmail["to"] = to;
                        if (fromParty != null)
                        {
                            entEmail["from"] = fromParty;
                            
                        }
                        entEmail["regardingobjectid"] = entRefCustomer;
                        Guid emailId = crmService.Create(entEmail);
                        Trace.AppendLine("emailId = " + emailId.ToString() + "...");


                        //---Create attachment of email
                        Customer customer = new Customer();
                        customer.CompanyName = "Company 123";
                        CustomerExcel customerExcel = new CustomerExcel(loggedInUserId, crmService, customer, Trace);
                        //customerExcel.CreateExcelDoc_Test();
                        byte[] fileStream = customerExcel.CreateExcelDoc();
                        Trace.AppendLine("1122----fileStream = " + fileStream.Length.ToString());
                        CreateExcelFileAttachmentForEmail(customerExcel.FileName, fileStream, emailId, crmService);

                        //--- Commenting temporarily this line of code as email sending is not enabled on current crm org. Due to this it is throwing error
                        //SendEmailResponse response = SendEmail(emailId, crmService);
                        areEmailsSent = true;

                    }
                    else
                    {
                        Trace.AppendLine("4.to.Length is ZEROOOOOOOOO");
                    }

                }
                else
                {
                    Trace.AppendLine("4.While sending email, either object OutageAlert is null or OutageAlertId is empty GUID.");
                }


            }
            return areEmailsSent;
        }

        public void CreateExcelFileAttachmentForEmail(string fileName, byte[] fileStream, Guid emailId, IOrganizationService crmService)
        {
            //---Create attachment of email
            Entity attachment = new Entity("activitymimeattachment");
            attachment["subject"] = "Customer Accounts";
            attachment["filename"] = fileName;
            //byte[] fileStream = new byte[] { }; //Set file stream bytes
            attachment["body"] = Convert.ToBase64String(fileStream);
            attachment["mimetype"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";//"text/csv";
            attachment["attachmentnumber"] = 1;
            attachment["objectid"] = new EntityReference("email", emailId);
            attachment["objecttypecode"] = "email";
            crmService.Create(attachment);
        }

        public SendEmailResponse SendEmail(Guid emailId, IOrganizationService crmService)
        {
            SendEmailResponse response = null;
            SendEmailRequest sendRequest = new SendEmailRequest();
            sendRequest.EmailId = emailId;
            sendRequest.TrackingToken = "";
            sendRequest.IssueSend = true;
            try
            {
                response = (SendEmailResponse)crmService.Execute(sendRequest);
            }
            catch (Exception ex)
            {
                ;
            }
            return response;
        }

        public DateTime ConvertCrmDateToUserDate(IOrganizationService service, Guid userGuid, DateTime inputDate)
        {
            //replace userid with id of user
            Entity userSettings = service.Retrieve("usersettings", userGuid, new ColumnSet("timezonecode"));
            //timezonecode for UTC is 85
            int timeZoneCode = 85;

            //retrieving timezonecode from usersetting
            if ((userSettings != null) && (userSettings["timezonecode"] != null))
            {
                timeZoneCode = (int)userSettings["timezonecode"];
            }
            //retrieving standard name
            var qe = new QueryExpression("timezonedefinition");
            qe.ColumnSet = new ColumnSet("standardname");
            qe.Criteria.AddCondition("timezonecode", ConditionOperator.Equal, timeZoneCode);
            EntityCollection TimeZoneDef = service.RetrieveMultiple(qe);

            TimeZoneInfo userTimeZone = null;
            if (TimeZoneDef.Entities.Count == 1)
            {
                String timezonename = TimeZoneDef.Entities[0]["standardname"].ToString();
                userTimeZone = TimeZoneInfo.FindSystemTimeZoneById(timezonename);
            }
            //converting date from UTC to user time zone
            DateTime cstTime = TimeZoneInfo.ConvertTimeFromUtc(inputDate, userTimeZone);

            return cstTime;
        }
    }
}
