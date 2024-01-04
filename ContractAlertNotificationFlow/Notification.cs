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

namespace ContractAlertNotificationFlow
{
    public class Notification:CodeActivity
    {
        [Input("Contract")]
        [RequiredArgument]
        [ReferenceTarget("entitlement")]
        public InArgument<EntityReference> Contract { get; set; }

        [Input("Email Admin")]
        [RequiredArgument]
        [ReferenceTarget("queue")]
        public InArgument<EntityReference> FromEmail { get; set; }

        [Input("To User 1")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> ToUser1 { get; set; }

        [Input("To User 2")]
        [RequiredArgument]
        [ReferenceTarget("systemuser")]
        public InArgument<EntityReference> ToUser2 { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1)
            //{
            //    return;
            //}
            Guid ContractGuid = Contract.Get<EntityReference>(context).Id;
            Guid FromEmailGuid = FromEmail.Get<EntityReference>(context).Id;
            Guid ToUser1Guid = ToUser1.Get<EntityReference>(context).Id;
            Guid ToUser2Guid = ToUser2.Get<EntityReference>(context).Id;
            executeBusinessLogic(ContractGuid, FromEmailGuid, ToUser1Guid, ToUser2Guid,service);
        }

        private void executeBusinessLogic(Guid contractGuid, Guid fromEmailGuid, Guid toUser1Guid, Guid toUser2Guid, IOrganizationService service)
        {
            try
            {
                string caseTitle = "";
                DateTime startDate = new DateTime();
                DateTime endDate = new DateTime();
                int startDateDay = 0;
                int endDateDay = 0;
                string ContractID = "";
                string CustomerName = "";

                bool janEmailSent = false;
                bool febEmailSent = false;
                bool marEmailSent = false;
                bool aprEmailSent = false;
                bool mayEmailSent = false;
                bool junEmailSent = false;
                bool julEmailSent = false;
                bool augEmailSent = false;
                bool sepEmailSent = false;
                bool octEmailSent = false;
                bool novEmailSent = false;
                bool decEmailSent = false;

                Guid its_contractscheduleemailnotificationGuid = new Guid();

                Entity PPMCaseSetup = service.Retrieve("entitlement", contractGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet { AllColumns = true });

                if (PPMCaseSetup.Attributes.Contains("its_scheduletype"))
                {
                    int scheduleTypeValue = ((OptionSetValue)PPMCaseSetup.Attributes["its_scheduletype"]).Value;
                    #region Schedule Values
                    //1- One Time Schedule
                    //2- Monthly Schedule
                    //3- Daily Schedule
                    //4- Yearly Schedule
                    #endregion
                    if (PPMCaseSetup.Attributes.Contains("itscs_contractid"))
                    {
                        ContractID = PPMCaseSetup.Attributes["itscs_contractid"].ToString();
                    }
                    if (PPMCaseSetup.Attributes.Contains("customerid"))
                    {
                        CustomerName = ((EntityReference)PPMCaseSetup.Attributes["customerid"]).Name;
                    }
                    if (PPMCaseSetup.Attributes.Contains("startdate"))
                    {
                        startDate = Convert.ToDateTime(PPMCaseSetup.Attributes["startdate"]).ToLocalTime();
                        startDateDay = startDate.Day;
                    }
                    if (PPMCaseSetup.Attributes.Contains("enddate"))
                    {
                        endDate = Convert.ToDateTime(PPMCaseSetup.Attributes["enddate"]).ToLocalTime();
                        endDateDay = endDate.Day;
                    }
                    if (PPMCaseSetup.Attributes.Contains("its_casetitle"))
                    {
                        caseTitle = PPMCaseSetup.Attributes["its_casetitle"].ToString();
                    }

                    if (scheduleTypeValue == 2)
                    {
                        bool january = Convert.ToBoolean(PPMCaseSetup.Attributes["its_january"]);
                        bool february = Convert.ToBoolean(PPMCaseSetup.Attributes["its_february"]);
                        bool march = Convert.ToBoolean(PPMCaseSetup.Attributes["its_march"]);
                        bool april = Convert.ToBoolean(PPMCaseSetup.Attributes["its_april"]);
                        bool may = Convert.ToBoolean(PPMCaseSetup.Attributes["its_may"]);
                        bool june = Convert.ToBoolean(PPMCaseSetup.Attributes["its_june"]);
                        bool jul = Convert.ToBoolean(PPMCaseSetup.Attributes["its_july"]);
                        bool august = Convert.ToBoolean(PPMCaseSetup.Attributes["its_august"]);
                        bool september = Convert.ToBoolean(PPMCaseSetup.Attributes["its_september"]);
                        bool october = Convert.ToBoolean(PPMCaseSetup.Attributes["its_october"]);
                        bool november = Convert.ToBoolean(PPMCaseSetup.Attributes["its_november"]);
                        bool december = Convert.ToBoolean(PPMCaseSetup.Attributes["its_december"]);

                        QueryExpression q1 = new QueryExpression();
                        q1.ColumnSet = new ColumnSet { AllColumns = true };
                        FilterExpression fe = new FilterExpression(LogicalOperator.And);
                        fe.AddCondition(new ConditionExpression("its_contractid", ConditionOperator.Equal, contractGuid));
                        q1.Criteria = fe;
                        q1.EntityName = "its_contractscheduleemailnotification";
                        EntityCollection ec = service.RetrieveMultiple(q1);
                        if (ec.Entities.Count > 0)
                        {
                            foreach (Entity c in ec.Entities)
                            {
                                its_contractscheduleemailnotificationGuid = new Guid(c.Attributes["its_contractscheduleemailnotificationid"].ToString());
                            }

                        }




                        if (PPMCaseSetup.Attributes.Contains("its_date"))
                        {
                            int dateValue = Convert.ToInt32(PPMCaseSetup.Attributes["its_date"]);
                            Entity E = new Entity("its_contractscheduleemailnotification");

                            var currentMonth = DateTime.Now.Month;
                            if (january && currentMonth == 1 && !janEmailSent)
                            {
                                DateTime dtJan = new DateTime(DateTime.Now.Year, 1, dateValue);
                                if (dtJan >= startDate && dtJan <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid,ContractID,CustomerName, service);
                                    E["its_janemailsent"] = true;
                                }

                            }
                            if (february && currentMonth == 2 && !febEmailSent)
                            {
                                DateTime dtFeb = new DateTime(DateTime.Now.Year, 2, dateValue);
                                if (dtFeb >= startDate && dtFeb <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_febemailsent"] = true;
                                }
                            }
                            if (march && currentMonth == 3 && !marEmailSent)
                            {
                                DateTime dtMarch = new DateTime(DateTime.Now.Year, 3, dateValue);
                                if (dtMarch >= startDate && dtMarch <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_maremailsent"] = true;
                                }
                            }
                            if (april && currentMonth == 4 && !aprEmailSent)
                            {
                                DateTime dtApril = new DateTime(DateTime.Now.Year, 4, dateValue);
                                if (dtApril >= startDate && dtApril <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_apremailsent"] = true;
                                }
                            }
                            if (may && currentMonth == 5 && !mayEmailSent)
                            {
                                DateTime dtMay = new DateTime(DateTime.Now.Year, 5, dateValue);
                                if (dtMay >= startDate && dtMay <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_mayemailsent"] = true;
                                }
                            }
                            if (june && currentMonth == 6 && !junEmailSent)
                            {
                                DateTime dtjune = new DateTime(DateTime.Now.Year, 6, dateValue);
                                if (dtjune >= startDate && dtjune <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_junemailsent"] = true;                                   
                                }
                            }
                            if (jul && currentMonth == 7 && !julEmailSent)
                            {
                                DateTime dtJul = new DateTime(DateTime.Now.Year, 7, dateValue);
                                if (dtJul >= startDate && dtJul <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_julemailsent"] = true;
                                }
                            }
                            if (august && currentMonth == 8 && !augEmailSent)
                            {
                                DateTime dtAugust = new DateTime(DateTime.Now.Year, 8, dateValue);
                                if (dtAugust >= startDate && dtAugust <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_augemailsent"] = true;
                                }
                            }
                            if (september && currentMonth == 9 && !sepEmailSent)
                            {
                                DateTime dtseptember = new DateTime(DateTime.Now.Year, 9, dateValue);
                                if (dtseptember >= startDate && dtseptember <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_sepemailsent"] = true;
                                }
                            }
                            if (october && currentMonth == 10 && !octEmailSent)
                            {
                                DateTime dtoctober = new DateTime(DateTime.Now.Year, 10, dateValue);
                                if (dtoctober >= startDate && dtoctober <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_octemailsent"] = true;
                                }
                            }
                            if (november && currentMonth == 11 && !novEmailSent)
                            {
                                DateTime dtnovember = new DateTime(DateTime.Now.Year, 11, dateValue);
                                if (dtnovember >= startDate && dtnovember <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_novemailsent"] = true;
                                }
                            }
                            if (december && currentMonth == 12 && !decEmailSent)
                            {
                                DateTime dtdecember = new DateTime(DateTime.Now.Year, 12, dateValue);
                                if (dtdecember >= startDate && dtdecember <= endDate)
                                {
                                    CreatePPMBookingDate(contractGuid, fromEmailGuid, toUser1Guid, toUser2Guid, ContractID, CustomerName, service);
                                    E["its_decemailsent"] = true;
                                }
                            }
                            E.Id = its_contractscheduleemailnotificationGuid;
                            service.Update(E);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }

        private static void CreatePPMBookingDate(Guid contractGuid,Guid fromEmailGuid,Guid toUser1Guid, Guid toUser2Guid,string ContractID,string CustomerName, IOrganizationService service)
        {
            Entity fromActivityParty = new Entity("activityparty");
            Entity toActivityParty = new Entity("activityparty");
            Entity toActivityParty1 = new Entity("activityparty");

            fromActivityParty["partyid"] = new EntityReference("queue", fromEmailGuid);
            toActivityParty["partyid"] = new EntityReference("systemuser", toUser1Guid);
            toActivityParty1["partyid"] = new EntityReference("systemuser", toUser2Guid);

            Entity email = new Entity("email");
            email["from"] = new Entity[] { fromActivityParty };
            email["to"] = new Entity[] { toActivityParty, toActivityParty1 };
            email["regardingobjectid"] = new EntityReference("entitlement", contractGuid);
            email["subject"] = "Contract Schedule Notification for ContractID " + ContractID;
            string body = "";
            body = "<div align='left' style='width:110px; font:12px Arial, Helvetica, sans-serif'>";
            body = body + "<div style='padding:10px'>";
            body = body + "Dear Team,<br /><br />";
            body = body + "The below Contract has a schedule and please find the below details." + "<br /><br />";
            body = body + "<b> ContractID: </b> " + ContractID + "<br /><br />";
            body = body + "<b> Customer Name: </b> " + CustomerName + "<br /><br />";
            //body = body + "<b> Project Task: </b> " + ProjectTaskName + "<br /><br />";
            //body = body + "<b> Resource/s: </b> " + rNames + "<br /><br />";
            //body = body + "<b> Customer Name: </b> " + ((EntityReference)Project.Attributes["msdyn_customer"]).Name + " <br /><br />";
            body = body + "<b> Thank You.</b><br /><br /><br />";
            body = body + "Thanks & Regards,<br /><br />";
            body = body + "ARC Admin<br /><br /></div>";

            email["description"] = body;
            email["directioncode"] = true;
            Guid emailId = service.Create(email);

            // Use the SendEmail message to send an e-mail message.
            SendEmailRequest sendEmailRequest = new SendEmailRequest
            {
                EmailId = emailId,
                TrackingToken = "",
                IssueSend = true
            };

            SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);

        }
    }
}
