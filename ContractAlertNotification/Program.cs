using Microsoft.Crm.Sdk.Messages;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.WebServiceClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using Xrm;

namespace ContractAlertNotification
{
    class Program
    {
        static void Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            string organizationUrl = "https://arcdev.crm4.dynamics.com";
            //string resourceURL = "https://org5f08e148.api.crm.dynamics.com/api/data/v9.2/";
            string clientId = "dc5aa586-a2ab-48a2-a4d7-0ee33e0b51b2"; // Client Id
            string appKey = "IdS8Q~Bmaryespa~iiQiGIaNqsK14WcHzTqaLb9e"; //Client Secret
            string tenantID = "15dfb30d-bd12-4473-b080-7916974c5ec1";

            //Create the Client credentials to pass for authentication
            ClientCredential clientcred = new ClientCredential(clientId, appKey);

            var credentials = new ClientCredential(clientId, appKey);
            var authContext = new AuthenticationContext(
                "https://login.microsoftonline.com/" + tenantID);
            var result = authContext.AcquireTokenAsync(organizationUrl, credentials);

            var token = result.Result.AccessToken;


            Uri serviceUrl = new Uri(organizationUrl + @"/xrmservices/2011/organization.svc/web?SdkClientVersion=9.2");
            OrganizationWebProxyClient sdkService;
            IOrganizationService service;
            using (sdkService = new OrganizationWebProxyClient(serviceUrl, false))
            {
                sdkService.HeaderToken = token.ToString();

                service = (IOrganizationService)sdkService != null ? (IOrganizationService)sdkService : null;

                Entity contactRecord = new Entity("incident");

                contactRecord.Attributes["title"] = "Case 22062022";
                contactRecord.Attributes["customerid"] = new EntityReference("account", new Guid("e1647f82-b3a1-e711-810e-5065f38aa961"));
                contactRecord.Attributes["description"] = "sss";
                var request = new RetrieveDuplicatesRequest

                {
                    //Entity Object to be searched with the values filled for the attributes to check
                    BusinessEntity = contactRecord,

                    //Logical Name of the Entity to check Matching Entity
                    MatchingEntityName = contactRecord.LogicalName,


                };

                var response = (RetrieveDuplicatesResponse)service.Execute(request);
                if (response.DuplicateCollection.Entities.Count > 0)
                {
                    throw new InvalidPluginExecutionException("Duplicate Case Found.");
                }

                executeBusinessLogic(new Guid("53cbee3b-56db-ec11-bb3d-0022489c0b3c"), service);
                //Guid PPMCaseSetupGuid = new Guid("03b552b1-76e8-ec11-bb3c-000d3adbd798");

                //string caseTitle = "";
                //DateTime startDate = new DateTime();
                //DateTime endDate = new DateTime();
                //int startDateDay = 0;
                //int endDateDay = 0;

                ////Entity PPMCaseSetup = service.Retrieve("itscs_ppmcasesetup", PPMCaseSetupGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet { AllColumns = true });
                //Entity PPMCaseSetup = service.Retrieve("entitlement", PPMCaseSetupGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet { AllColumns = true });
                //int ContractStatus = ((OptionSetValue)PPMCaseSetup.Attributes["statuscode"]).Value;

                //QueryExpression q1 = new QueryExpression();
                //q1.ColumnSet = new ColumnSet("itscs_ppmcasebookingdatesid", "its_contractid");
                //FilterExpression fe = new FilterExpression(LogicalOperator.And);
                //fe.AddCondition(new ConditionExpression("its_contractid", ConditionOperator.Equal, PPMCaseSetupGuid));
                //q1.Criteria = fe;
                //q1.EntityName = "itscs_ppmcasebookingdates";
                //EntityCollection ec = service.RetrieveMultiple(q1);
                //if (ec.Entities.Count > 0)
                //{

                //}
                //else
                //{
                //    //Entity PPMCaseSetup = service.Retrieve("entitlement", PPMCaseSetupGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet { AllColumns = true });

                //    if (PPMCaseSetup.Attributes.Contains("its_scheduletype"))
                //    {
                //        int scheduleTypeValue = ((OptionSetValue)PPMCaseSetup.Attributes["its_scheduletype"]).Value;
                //        #region Schedule Values
                //        //1- One Time Schedule
                //        //2- Monthly Schedule
                //        //3- Daily Schedule
                //        //4- Yearly Schedule
                //        #endregion
                //        if (PPMCaseSetup.Attributes.Contains("startdate"))
                //        {
                //            startDate = Convert.ToDateTime(PPMCaseSetup.Attributes["startdate"]).ToLocalTime();
                //            startDateDay = startDate.Day;
                //        }
                //        if (PPMCaseSetup.Attributes.Contains("enddate"))
                //        {
                //            endDate = Convert.ToDateTime(PPMCaseSetup.Attributes["enddate"]).ToLocalTime();
                //            endDateDay = endDate.Day;
                //        }
                //        if (PPMCaseSetup.Attributes.Contains("its_casetitle"))
                //        {
                //            caseTitle = PPMCaseSetup.Attributes["its_casetitle"].ToString();
                //        }


                //        if (scheduleTypeValue == 1)
                //        {
                //            if (PPMCaseSetup.Attributes.Contains("its_scheduledate"))
                //            {
                //                DateTime scheduleDate = Convert.ToDateTime(PPMCaseSetup.Attributes["its_scheduledate"]).ToLocalTime();

                //                if (scheduleDate >= startDate && scheduleDate <= endDate)
                //                {
                //                    Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
                //                    PPMBookingDate["itscs_name"] = caseTitle;
                //                    //PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
                //                    PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
                //                    PPMBookingDate["itscs_bookingdate"] = scheduleDate.ToUniversalTime();
                //                    Guid PPMBookingDateGuid = service.Create(PPMBookingDate);
                //                }
                //                else
                //                {
                //                    throw new InvalidPluginExecutionException("The entered date is not in a range");
                //                }
                //            }
                //        }
                //        else if (scheduleTypeValue == 2)
                //        {
                //            bool january = Convert.ToBoolean(PPMCaseSetup.Attributes["its_january"]);
                //            bool february = Convert.ToBoolean(PPMCaseSetup.Attributes["its_february"]);
                //            bool march = Convert.ToBoolean(PPMCaseSetup.Attributes["its_march"]);
                //            bool april = Convert.ToBoolean(PPMCaseSetup.Attributes["its_april"]);
                //            bool may = Convert.ToBoolean(PPMCaseSetup.Attributes["its_may"]);
                //            bool june = Convert.ToBoolean(PPMCaseSetup.Attributes["its_june"]);
                //            bool jul = Convert.ToBoolean(PPMCaseSetup.Attributes["its_july"]);
                //            bool august = Convert.ToBoolean(PPMCaseSetup.Attributes["its_august"]);
                //            bool september = Convert.ToBoolean(PPMCaseSetup.Attributes["its_september"]);
                //            bool october = Convert.ToBoolean(PPMCaseSetup.Attributes["its_october"]);
                //            bool november = Convert.ToBoolean(PPMCaseSetup.Attributes["its_november"]);
                //            bool december = Convert.ToBoolean(PPMCaseSetup.Attributes["its_december"]);

                //            bool janEmailSent = Convert.ToBoolean(PPMCaseSetup.Attributes["its_januaryemailsent"]);
                //            if (PPMCaseSetup.Attributes.Contains("its_date"))
                //            {
                //                int dateValue = Convert.ToInt32(PPMCaseSetup.Attributes["its_date"]);
                //                //if (dateValue >= startDateDay || dateValue <= endDateDay)
                //                //{
                //                var currentMonth = DateTime.Now.Month;
                //                if (january && currentMonth==1)
                //                {
                //                    DateTime dtJan = new DateTime(DateTime.Now.Year, 1, dateValue);
                //                    if (dtJan >= startDate && dtJan <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 1, dateValue, service);
                //                    }

                //                }
                //                if (february && currentMonth == 2)
                //                {
                //                    DateTime dtFeb = new DateTime(DateTime.Now.Year, 2, dateValue);
                //                    if (dtFeb >= startDate && dtFeb <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 2, dateValue, service);
                //                    }
                //                }
                //                if (march && currentMonth == 3)
                //                {
                //                    DateTime dtMarch = new DateTime(DateTime.Now.Year, 3, dateValue);
                //                    if (dtMarch >= startDate && dtMarch <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 3, dateValue, service);
                //                    }
                //                }
                //                if (april && currentMonth == 4)
                //                {
                //                    DateTime dtApril = new DateTime(DateTime.Now.Year, 4, dateValue);
                //                    if (dtApril >= startDate && dtApril <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 4, dateValue, service);
                //                    }
                //                }
                //                if (may && currentMonth == 5)
                //                {
                //                    DateTime dtMay = new DateTime(DateTime.Now.Year, 5, dateValue);
                //                    if (dtMay >= startDate && dtMay <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 5, dateValue, service);
                //                    }
                //                }
                //                if (june && currentMonth == 6 && !janEmailSent)
                //                {
                //                    DateTime dtjune = new DateTime(DateTime.Now.Year, 6, dateValue);
                //                    if (dtjune >= startDate && dtjune <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 6, dateValue, service);
                //                        Entity E = new Entity("entitlement");
                //                        E["its_januaryemailsent"] = true;
                //                        E.Id = PPMCaseSetupGuid;
                //                        service.Update(E);
                //                    }
                //                }
                //                if (jul && currentMonth == 7)
                //                {
                //                    DateTime dtJul = new DateTime(DateTime.Now.Year, 7, dateValue);
                //                    if (dtJul >= startDate && dtJul <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 7, dateValue, service);
                //                    }
                //                }
                //                if (august && currentMonth == 8)
                //                {
                //                    DateTime dtAugust = new DateTime(DateTime.Now.Year, 8, dateValue);
                //                    if (dtAugust >= startDate && dtAugust <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 8, dateValue, service);
                //                    }
                //                }
                //                if (september && currentMonth == 9)
                //                {
                //                    DateTime dtseptember = new DateTime(DateTime.Now.Year, 9, dateValue);
                //                    if (dtseptember >= startDate && dtseptember <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 9, dateValue, service);
                //                    }
                //                }
                //                if (october && currentMonth == 10)
                //                {
                //                    DateTime dtoctober = new DateTime(DateTime.Now.Year, 10, dateValue);
                //                    if (dtoctober >= startDate && dtoctober <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 10, dateValue, service);
                //                    }
                //                }
                //                if (november && currentMonth == 11)
                //                {
                //                    DateTime dtnovember = new DateTime(DateTime.Now.Year, 11, dateValue);
                //                    if (dtnovember >= startDate && dtnovember <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 11, dateValue, service);
                //                    }
                //                }
                //                if (december && currentMonth == 12)
                //                {
                //                    DateTime dtdecember = new DateTime(DateTime.Now.Year, 12, dateValue);
                //                    if (dtdecember >= startDate && dtdecember <= endDate)
                //                    {
                //                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 12, dateValue, service);
                //                    }
                //                }

                //                //}
                //                //else
                //                //{
                //                //    throw new InvalidPluginExecutionException("The entered date is not in a range");
                //                //}
                //            }
                //        }
                //        else if (scheduleTypeValue == 3)
                //        {

                //            for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                //            {
                //                //DateTime result = day;
                //                //Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
                //                //PPMBookingDate["itscs_name"] = caseTitle + day;
                //                ////PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
                //                //PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
                //                //PPMBookingDate["itscs_bookingdate"] = result.ToUniversalTime();
                //                //Guid PPMBookingDateGuid = service.Create(PPMBookingDate);
                //            }
                //        }
                //        //else if (scheduleTypeValue == 4)
                //        //{

                //        //}
                //    }
                //}
            }
        }

        //private static void CreatePPMBookingDate(string caseTitle, Guid PPMCaseSetupGuid, int month, int dateValue, IOrganizationService service)
        //{
        //    //Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
        //    //PPMBookingDate["itscs_name"] = caseTitle;
        //    ////PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
        //    //PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
        //    //PPMBookingDate["itscs_bookingdate"] = new DateTime(DateTime.Now.Year, month, dateValue);
        //    //Guid PPMBookingDateGuid = service.Create(PPMBookingDate);

        //    Guid queueGUID = new Guid("43786172-c2ac-ea11-a812-000d3ab19dd4");
        //    Guid ToUserGuid1 = new Guid("b03fb559-25f5-e711-8111-5065f38bd371");
        //    Guid ToUserGuid2 = new Guid("b03fb559-25f5-e711-8111-5065f38bd371");


        //    Entity fromActivityParty = new Entity("activityparty");
        //    Entity toActivityParty = new Entity("activityparty");
        //    Entity toActivityParty1 = new Entity("activityparty");
        //    ////List<ActivityParty> ToPartyList = new List<ActivityParty>();


        //    fromActivityParty["partyid"] = new EntityReference("queue", queueGUID);
        //    toActivityParty["partyid"] = new EntityReference("systemuser", ToUserGuid1);
        //    toActivityParty1["partyid"] = new EntityReference("systemuser", ToUserGuid1);

        //    Entity email = new Entity("email");
        //    email["from"] = new Entity[] { fromActivityParty };
        //    email["to"] = new Entity[] { toActivityParty,toActivityParty1 };
        //    email["regardingobjectid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
        //    email["subject"] = "This is the subject";
        //    string body = "";
        //    body = "<div align='left' style='width:110px; font:12px Arial, Helvetica, sans-serif'>";
        //    body = body + "<div style='padding:10px'>";
        //    body = body + "Dear Team,<br /><br />";
        //    body = body + "The below Project Task has been marked as completed." + "<br /><br />";
        //    //body = body + "<b> Project: </b> " + Project.Attributes["msdyn_subject"].ToString() + "<br /><br />";
        //    //body = body + "<b> Project Number: </b> " + Project.Attributes["itspsa_projectnumber"].ToString() + "<br /><br />";
        //    //body = body + "<b> Project Task: </b> " + ProjectTaskName + "<br /><br />";
        //    //body = body + "<b> Resource/s: </b> " + rNames + "<br /><br />";
        //    //body = body + "<b> Customer Name: </b> " + ((EntityReference)Project.Attributes["msdyn_customer"]).Name + " <br /><br />";
        //    body = body + "<b> Thank You.</b><br /><br /><br />";
        //    body = body + "Thanks & Regards,<br /><br />";
        //    body = body + "ARC Admin<br /><br /></div>";

        //    email["description"] = body;
        //    email["directioncode"] = true;
        //    Guid emailId = service.Create(email);

        //    // Use the SendEmail message to send an e-mail message.
        //    SendEmailRequest sendEmailRequest = new SendEmailRequest
        //    {
        //        EmailId = emailId,
        //        TrackingToken = "",
        //        IssueSend = true
        //    };

        //    SendEmailResponse sendEmailresp = (SendEmailResponse)service.Execute(sendEmailRequest);


        //    ////Email email = new Email();
        //    //Entity email = new Entity("email");

        //    ////ActivityParty fromParty = new ActivityParty();
        //    ////fromParty.PartyId = new EntityReference(Queue.EntityLogicalName, queueGUID);

        //    ////List<ActivityParty> ToPartyList = new List<ActivityParty>();

        //    ////ActivityParty toPartyPM = new ActivityParty();
        //    ////toPartyPM.PartyId = new EntityReference(SystemUser.EntityLogicalName, ToUserGuid1);
        //    ////ToPartyList.Add(toPartyPM);

        //    //Entity fromParty = new Entity("activityparty");
        //    ////ActivityParty fromParty = new ActivityParty();
        //    //fromParty.Attributes["partyId"] = new EntityReference(Queue.EntityLogicalName, queueGUID);

        //    //Entity to = new Entity("activityparty");
        //    //List<to> ToPartyList = new List<to>();

        //    //Entity toPartyPM = new Entity("activityparty");
        //    ////ActivityParty toPartyPM = new ActivityParty();
        //    //toPartyPM.Attributes["partyid"] = new EntityReference(SystemUser.EntityLogicalName, ToUserGuid1);
        //    //ToPartyList.Add(toPartyPM);

        //    ////ActivityParty toPartyTL = new ActivityParty();
        //    ////toPartyTL.PartyId = new EntityReference(SystemUser.EntityLogicalName, ToUserGuid2);
        //    ////ToPartyList.Add(toPartyTL);


        //    ////List<ActivityParty> CCPartyList = new List<ActivityParty>();
        //    ////List<string> resourceNames = new List<string>();
        //    ////foreach (var guid in resources)
        //    ////{
        //    ////    ActivityParty ccParty = new ActivityParty();
        //    ////    ccParty.PartyId = new EntityReference(SystemUser.EntityLogicalName, guid);
        //    ////    CCPartyList.Add(ccParty);
        //    ////    //Get Names
        //    ////    Entity Resource = service.Retrieve("systemuser", guid, new ColumnSet("fullname"));
        //    ////    resourceNames.Add(Resource.Attributes["fullname"].ToString());
        //    ////}
        //    ////string rNames = string.Join(";", resourceNames.ToArray());

        //    //email.Attributes["from"] = new[] { fromParty };
        //    //email.Attributes["to"] = ToPartyList.ToArray();
        //    ////email.Cc = CCPartyList.ToArray();
        //    //email.Attributes["regardingobjectid"] = new EntityReference("entitlement", PPMCaseSetupGuid);

        //    //email.Attributes["subject"] = "AMC Notification";
        //    //string body = "";
        //    //body = "<div align='left' style='width:110px; font:12px Arial, Helvetica, sans-serif'>";
        //    //body = body + "<div style='padding:10px'>";
        //    ////body = body + "Dear Team,<br /><br />";
        //    //body = body + "The below Project Task has been marked as completed." + "<br /><br />";
        //    ////body = body + "<b> Project: </b> " + Project.Attributes["msdyn_subject"].ToString() + "<br /><br />";
        //    ////body = body + "<b> Project Number: </b> " + Project.Attributes["itspsa_projectnumber"].ToString() + "<br /><br />";
        //    ////body = body + "<b> Project Task: </b> " + ProjectTaskName + "<br /><br />";
        //    ////body = body + "<b> Resource/s: </b> " + rNames + "<br /><br />";
        //    ////body = body + "<b> Customer Name: </b> " + ((EntityReference)Project.Attributes["msdyn_customer"]).Name + " <br /><br />";
        //    //body = body + "<b> Thank You.</b><br /><br /><br />";
        //    //body = body + "Thanks & Regards,<br /><br />";
        //    //body = body + "ARC Admin<br /><br /></div>";
        //    ////email.Description = body;

        //    //Guid EmailId = service.Create(email);

        //    //SendEmailRequest req = new SendEmailRequest();
        //    //req.EmailId = EmailId;
        //    //req.IssueSend = true;
        //    //req.TrackingToken = "";
        //    //SendEmailResponse res = (SendEmailResponse)service.Execute(req);
        //}

        private static void executeBusinessLogic(Guid PPMCaseSetupGuid, IOrganizationService service)
        {
            try
            {
                string caseTitle = "";
                DateTime startDate = new DateTime();
                DateTime endDate = new DateTime();
                int startDateDay = 0;
                int endDateDay = 0;
                string ContractID = "";
                //Entity PPMCaseSetup = service.Retrieve("itscs_ppmcasesetup", PPMCaseSetupGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet { AllColumns = true });
                Entity PPMCaseSetup = service.Retrieve("entitlement", PPMCaseSetupGuid, new Microsoft.Xrm.Sdk.Query.ColumnSet { AllColumns = true });
                if (PPMCaseSetup.Attributes.Contains("itscs_contractid"))
                {
                    ContractID = PPMCaseSetup.Attributes["itscs_contractid"].ToString();
                }

                //Booking Schedule
                QueryExpression q2 = new QueryExpression();
                q2.ColumnSet = new ColumnSet("its_contractid");
                FilterExpression fe2 = new FilterExpression(LogicalOperator.And);
                fe2.AddCondition(new ConditionExpression("its_contractid", ConditionOperator.Equal, PPMCaseSetupGuid));
                q2.Criteria = fe2;
                q2.EntityName = "its_contractscheduleemailnotification";
                EntityCollection ec2 = service.RetrieveMultiple(q2);
                if (ec2.Entities.Count > 0)
                {

                }
                else
                {
                    Entity CSEN = new Entity("its_contractscheduleemailnotification");
                    CSEN["its_name"] = ContractID;
                    CSEN["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
                    Guid CSEDGuid = service.Create(CSEN);
                }

                //Booking Dates
                QueryExpression q1 = new QueryExpression();
                q1.ColumnSet = new ColumnSet("itscs_ppmcasebookingdatesid", "its_contractid");
                FilterExpression fe = new FilterExpression(LogicalOperator.And);
                fe.AddCondition(new ConditionExpression("its_contractid", ConditionOperator.Equal, PPMCaseSetupGuid));
                q1.Criteria = fe;
                q1.EntityName = "itscs_ppmcasebookingdates";
                EntityCollection ec = service.RetrieveMultiple(q1);
                if (ec.Entities.Count > 0)
                {

                }
                else
                {


                    if (PPMCaseSetup.Attributes.Contains("its_scheduletype"))
                    {
                        int scheduleTypeValue = ((OptionSetValue)PPMCaseSetup.Attributes["its_scheduletype"]).Value;
                        #region Schedule Values
                        //1- One Time Schedule
                        //2- Monthly Schedule
                        //3- Daily Schedule
                        //4- Yearly Schedule
                        #endregion
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


                        if (scheduleTypeValue == 1)
                        {
                            if (PPMCaseSetup.Attributes.Contains("its_scheduledate"))
                            {
                                DateTime scheduleDate = Convert.ToDateTime(PPMCaseSetup.Attributes["its_scheduledate"]).ToLocalTime();

                                if (scheduleDate >= startDate && scheduleDate <= endDate)
                                {
                                    Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
                                    PPMBookingDate["itscs_name"] = caseTitle;
                                    //PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
                                    PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
                                    PPMBookingDate["itscs_bookingdate"] = scheduleDate.ToUniversalTime();

                                    Guid PPMBookingDateGuid = service.Create(PPMBookingDate);
                                }
                                else
                                {
                                    throw new InvalidPluginExecutionException("The entered date is not in a range");
                                }
                            }
                        }
                        else if (scheduleTypeValue == 2)
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
                            if (PPMCaseSetup.Attributes.Contains("its_date"))
                            {
                                int dateValue = Convert.ToInt32(PPMCaseSetup.Attributes["its_date"]);
                                //if (dateValue >= startDateDay || dateValue <= endDateDay)
                                //{
                                if (january)
                                {
                                    DateTime dtJan = new DateTime(DateTime.Now.Year, 1, dateValue);
                                    if (dtJan >= startDate && dtJan <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 1, dateValue, service);
                                    }

                                }
                                if (february && dateValue<=29)
                                {
                                    DateTime dtFeb = new DateTime(DateTime.Now.Year, 2, dateValue);
                                    if (dtFeb >= startDate && dtFeb <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 2, dateValue, service);
                                    }
                                }
                                if (march)
                                {
                                    DateTime dtMarch = new DateTime(DateTime.Now.Year, 3, dateValue);
                                    if (dtMarch >= startDate && dtMarch <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 3, dateValue, service);
                                    }
                                }
                                if (april)
                                {
                                    DateTime dtApril = new DateTime(DateTime.Now.Year, 4, dateValue);
                                    if (dtApril >= startDate && dtApril <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 4, dateValue, service);
                                    }
                                }
                                if (may)
                                {
                                    DateTime dtMay = new DateTime(DateTime.Now.Year, 5, dateValue);
                                    if (dtMay >= startDate && dtMay <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 5, dateValue, service);
                                    }
                                }
                                if (june)
                                {
                                    DateTime dtjune = new DateTime(DateTime.Now.Year, 6, dateValue);
                                    if (dtjune >= startDate && dtjune <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 6, dateValue, service);
                                    }
                                }
                                if (jul)
                                {
                                    DateTime dtJul = new DateTime(DateTime.Now.Year, 7, dateValue);
                                    if (dtJul >= startDate && dtJul <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 7, dateValue, service);
                                    }
                                }
                                if (august)
                                {
                                    DateTime dtAugust = new DateTime(DateTime.Now.Year, 8, dateValue);
                                    if (dtAugust >= startDate && dtAugust <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 8, dateValue, service);
                                    }
                                }
                                if (september)
                                {
                                    DateTime dtseptember = new DateTime(DateTime.Now.Year, 9, dateValue);
                                    if (dtseptember >= startDate && dtseptember <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 9, dateValue, service);
                                    }
                                }
                                if (october)
                                {
                                    DateTime dtoctober = new DateTime(DateTime.Now.Year, 10, dateValue);
                                    if (dtoctober >= startDate && dtoctober <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 10, dateValue, service);
                                    }
                                }
                                if (november)
                                {
                                    DateTime dtnovember = new DateTime(DateTime.Now.Year, 11, dateValue);
                                    if (dtnovember >= startDate && dtnovember <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 11, dateValue, service);
                                    }
                                }
                                if (december)
                                {
                                    DateTime dtdecember = new DateTime(DateTime.Now.Year, 12, dateValue);
                                    if (dtdecember >= startDate && dtdecember <= endDate)
                                    {
                                        CreatePPMBookingDate(caseTitle, PPMCaseSetupGuid, 12, dateValue, service);
                                    }
                                }

                                //}
                                //else
                                //{
                                //    throw new InvalidPluginExecutionException("The entered date is not in a range");
                                //}
                            }
                        }
                        else if (scheduleTypeValue == 3)
                        {
                            for (var day = startDate.Date; day.Date <= endDate.Date; day = day.AddDays(1))
                            {
                                DateTime result = day;
                                Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
                                PPMBookingDate["itscs_name"] = caseTitle + day;
                                //PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
                                PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
                                PPMBookingDate["itscs_bookingdate"] = result.ToUniversalTime();
                                Guid PPMBookingDateGuid = service.Create(PPMBookingDate);
                            }
                        }
                        //else if (scheduleTypeValue == 4)
                        //{

                        //}
                    }
                }
            }
            //catch (InvalidPluginExecutionException ex)
            //{
            //    if (ex != null)
            //    {

            //        throw new InvalidPluginExecutionException("1) Unable to complete the Operation." + ex.Message + "Contact Your Administrator");
            //    }
            //    else
            //        throw new InvalidPluginExecutionException("2) Unable to complete the Operation. Contact Administrator");
            //}
            //catch (FaultException<OrganizationServiceFault> ex)
            //{
            //    throw new InvalidPluginExecutionException("3) Unable to complete the Operation.", ex);
            //}
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.ToString());
            }
        }
        private static void CreatePPMBookingDate(string caseTitle, Guid PPMCaseSetupGuid, int month, int dateValue, IOrganizationService service)
        {
            Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
            PPMBookingDate["itscs_name"] = caseTitle;
            //PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
            PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
            PPMBookingDate["itscs_bookingdate"] = new DateTime(DateTime.Now.Year, month, dateValue);
            Guid PPMBookingDateGuid = service.Create(PPMBookingDate);
        }
    }
}
