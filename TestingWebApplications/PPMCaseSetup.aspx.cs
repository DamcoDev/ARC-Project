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
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class PPMCaseSetup : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
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

            executeBusinessLogic(new Guid("53cbee3b-56db-ec11-bb3d-0022489c0b3c"), service);

        }
    }
    private void executeBusinessLogic(Guid PPMCaseSetupGuid, IOrganizationService service)
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
                            if (february)
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
    private void CreatePPMBookingDate(string caseTitle, Guid PPMCaseSetupGuid, int month, int dateValue, IOrganizationService service)
    {
        Entity PPMBookingDate = new Entity("itscs_ppmcasebookingdates");
        PPMBookingDate["itscs_name"] = caseTitle;
        //PPMBookingDate["itscs_ppmcasesetupid"] = new EntityReference("itscs_ppmcasesetup", PPMCaseSetupGuid);
        PPMBookingDate["its_contractid"] = new EntityReference("entitlement", PPMCaseSetupGuid);
        PPMBookingDate["itscs_bookingdate"] = new DateTime(DateTime.Now.Year, month, dateValue);
        Guid PPMBookingDateGuid = service.Create(PPMBookingDate);
    }

}