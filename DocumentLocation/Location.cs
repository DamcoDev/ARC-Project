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
using System;
using Microsoft.SharePoint.Client;
using System.Security;

namespace DocumentLocation
{
    public class Location:CodeActivity
    {
        [Input("SP URL")]
        public InArgument<string> SPURL { get; set; }

        [Input("FolderName")]
        public InArgument<string> FolderName { get; set; }

        [Input("ProjectDocument")]
        [ReferenceTarget("itspsa_projectdocuments")]
        public InArgument<EntityReference> ProjectDocument { get; set; }
        // push & pull by github
        protected override void Execute(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = (IWorkflowContext)context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = (IOrganizationService)serviceFactory.CreateOrganizationService(workflowContext.UserId);

            //if (workflowContext.Depth > 1) { return; }
            string spurl = SPURL.Get<string>(context);
            string foldername = FolderName.Get<string>(context);
            Guid ProjectDocumentGUID = ProjectDocument.Get<EntityReference>(context).Id;

            executeBusinessLogic(spurl, foldername, ProjectDocumentGUID,service);
        }

        private void executeBusinessLogic(string spurl, string foldername, Guid projectDocumentGUID, IOrganizationService service)
        {
            try
            {
                CreateSPOContext(spurl);
                using (ClientContext clientContext = CreateSPOContext(spurl))
                {
                    List list = clientContext.Web.Lists.GetByTitle("Project Documents");

                    clientContext.Load(list);
                    clientContext.Load(list.RootFolder);
                    clientContext.Load(list.RootFolder.Folders);
                    clientContext.Load(list.RootFolder.Files);
                    clientContext.ExecuteQuery();
                    FolderCollection fcol = list.RootFolder.Folders;
                    //List listFile = new MList();
                    foreach (Folder f in fcol)
                    {
                        if (f.Name == foldername)
                        {
                            clientContext.Load(f.Files);
                            clientContext.ExecuteQuery();
                            //FileCollection fileCol = f.Files;
                            var fileCount = f.Files.Count;
                            if(fileCount>0)
                            {
                                Entity PD = new Entity("itspsa_projectdocuments");
                                PD["its_documentattached"] = true;
                                PD.Id = projectDocumentGUID;
                                service.Update(PD);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
        private static ClientContext CreateSPOContext(string web)
        {
            var securePassword = new SecureString();
            var password = "3edc#EDC2";

            foreach (char c in password.ToCharArray()) securePassword.AppendChar(c);

            var context = new ClientContext(web);
            context.Credentials = new SharePointOnlineCredentials("itdev@alrostamanigroup.onmicrosoft.com", securePassword);
            context.RequestTimeout = Int32.MaxValue;
            context.ExecuteQuery();
            return context;
        }
    }
}
