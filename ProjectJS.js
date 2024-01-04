function disableAllFields(formContext) {
    var context = formContext.getFormContext().data.entity;
    //var closeProjectVal=context.attributes.getByName('header_process_itspsa_closeproject');
    var closeProjectVal = Xrm.Page.getAttribute("itspsa_closeproject").getValue();
    if (closeProjectVal == 110920000) {
        Xrm.Page.data.entity.attributes.forEach(function (attribute, index) {

            var control = Xrm.Page.getControl(attribute.getName());

            if (control) {
                control.setDisabled(true)
            }
        });
    }
    else {
        return;
        /* Xrm.Page.data.entity.attributes.forEach(function (attribute, index) {
 
             var control = Xrm.Page.getControl(attribute.getName());
 
             if (control) {
                 control.setDisabled(false)
             }
         });*/
    }
}


function stagePhases() {
    var stageName = Xrm.Page.getAttribute("msdyn_stagename").getValue();
    if (stageName == "Monitoring & Controlling" || stageName == "Close") {
        Xrm.Page.ui.tabs.get("ProjectStatus").setVisible(true);
    }
    else {
        Xrm.Page.ui.tabs.get("ProjectStatus").setVisible(false);
    }
}

/*msdyn_projectid: "ebb03fd6-73fc-e811-a958-000d3ab5a84e",
PATCH [Organization URI]/api/data/v9.0/new_mycustombpfs(dc2ab599-306d-e811-80ff-00155d513100) HTTP/1.1   
Content-Type: application/json   
OData-MaxVersion: 4.0   
OData-Version: 4.0 
  
{ 
    "statecode" : "1", 
    "statuscode": "3" 
}

Similarly, to reactivate a process instance, replace the statecode and statuscode values in the above code with 0 and 1 respectively.

Finally, to set a process instance status as Finished, which is only possible at the last stage of a process instance, replace the statecode and statuscode values in the above code with 0 and 2 respectively.*/


/*function OnSave(executionContext) {
    
executionContext.getEventArgs();   
    var formType = Xrm.Page.ui.getFormType();
    if (formType != 1) {
      if
        (Xrm.Page.getAttribute("itspsa_closeproject").getValue()==110920000)
        {

        //var pocopyValue = Xrm.Page.getAttribute("itspsa_ispocopyuploaded").getValue();
        //var proposalValue = Xrm.Page.getAttribute("itspsa_isproposaluploaded").getValue();
        //if (pocopyValue == 1 || proposalValue == 1) {
            if (Xrm.Page.data.entity.getId() === null) return;

            var serverUrl = Xrm.Page.context.getClientUrl();
            var oDataSelect = serverUrl + "/XRMServices/2011/OrganizationData.svc/AnnotationSet?$filter=ObjectId/Id eq guid'" + Xrm.Page.data.entity.getId() + "'&$select=IsDocument";
            var allowSave = false;
            var req = new XMLHttpRequest();
            req.open("GET", oDataSelect, false);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json;charset=utf-8");
            req.onreadystatechange = function () {
                if (req.readyState === 4) {
                    if (req.status === 200) {
                        var retrieved = JSON.parse(req.responseText).d;
                        for (var i = 0; i < retrieved.results.length; i++) {
                            if (retrieved.results.length >= 1) {
                                allowSave = true;
                                break;
                            }
                            //if (retrieved.results[0].IsDocument) {
                            //    allowSave = true;
                            //    break;
                            //}
                        }
                    }
                }
            };
            req.send();

            if (!allowSave) {
                debugger;
                //alert("Please Upload Required Documents");
                

                //var GUID = Xrm.Page.data.entity.getId();
                //GUID = GUID.replace('{', '').replace('}', '');
                //var url = "https://arcdev.crm4.dynamics.com/notes/edit.aspx?hideDesc=1&pId=%7b" + GUID + "%7d&pType=10191";

                //var openUrlOptions = {
                //    height: 400,
                //    width: 800
                //};

                //Xrm.Navigation.openUrl(url, openUrlOptions);
                ////executionContext.getEventArgs();   
                executionContext.getEventArgs().preventDefault();

 
            }
       }
  
    }

}*/



function actualEnd() {
    //debugger;
    if (Xrm.Page.getAttribute("itspsa_closeproject").getValue() == 110920000) {
        var currentDate = new Date();

        currentDate.setHours(currentDate.getHours());
        //alert(currentDate );
        Xrm.Page.getAttribute("msdyn_actualend").setValue(currentDate);
    }
    else { Xrm.Page.getAttribute("msdyn_actualend").setValue(null); }
}




function OnSave(executionContext) {
    var formContext = executionContext.getFormContext();
    var id = Xrm.Page.data.entity.getId().slice(1, -1)
    Xrm.WebApi.retrieveMultipleRecords("annotation", "?$filter=_objectid_value eq " + id + "").then(success, error)
    function success(result) {
        //debugger;
        for (var i = 0; i < result.entities.length; i++)
            if (result.entities.length >= 1) {
                return;
            } else { executionContext.getEventArgs().preventDefault(); }
    }
    function error(result) {
        //debugger;
    }
    executionContext.getEventArgs().preventDefault();
    formContext.data.process.addOnPreStageChange(OnSave);
}

function preventSave(econtext) {
    //debugger;
    var eventArgs = econtext.getEventArgs();
    if (econtext != null && econtext.getEventArgs() != null) {
        if (eventArgs.getSaveMode() == 1) {
            eventArgs.preventDefault();
        }
    }
}

//onchange of Project Status
function OnCloseProjectCompleted() {

    if (Xrm.Page.getAttribute("itspsa_closeproject").getValue() == 110920000) {
        var id = Xrm.Page.data.entity.getId().slice(1, -1);

        var fetchXML = "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">"+
  "<entity name=\"msdyn_resourceassignment\">"+
    "<attribute name=\"msdyn_resourceassignmentid\" />"+
    "<filter type=\"and\">"+
      "<condition attribute=\"msdyn_projectid\" operator=\"eq\" uiname=\"Commvault BackupforAzure\" uitype=\"msdyn_project\" value=\"{39c56641-47af-ec11-9841-000d3ab86930}\" />" +
    "</filter>"+
    "<link-entity name=\"msdyn_projecttask\" from=\"msdyn_projecttaskid\" to=\"msdyn_taskid\" link-type=\"inner\" alias=\"ab\">"+
      "<filter type=\"and\">"+
        "<condition attribute=\"itspsa_pmstatus\" operator=\"eq\" value=\"110920000\" />"+
      "</filter>"+
    "</link-entity>"+
    "</entity>"+
"</fetch>";

        var encodedFetchXML = encodeURI(fetchXML);
        var queryPath = "https://arc.api.crm4.dynamics.com/api/data/v9.2/msdyn_resourceassignments?fetchXml=" + encodedFetchXML;
        var requestPath = queryPath;

        var req = new XMLHttpRequest();

        req.open("GET", requestPath, true);

        req.setRequestHeader("Accept", "application/json");

        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

        req.onreadystatechange = function () {
            if (this.readyState === 4) {
                this.onreadystatechange = null;

                if (this.status === 200) {
                    var returned = JSON.parse(this.responseText);

                    var results = returned.value;
                    alert(results);
                    alert(results.length);
                }
                else {

                    alert(this.statusText);

                }
            }
        };
        req.send();


        // Xrm.WebApi.retrieveMultipleRecords("msdyn_projecttask", "?$filter=_msdyn_project_value eq " + id + " and itspsa_pmstatus ne 110920000").then(successa, errora)
        //Xrm.WebApi.retrieveMultipleRecords("msdyn_resourceassignment", "?$filter=_msdyn_projectid_value eq " + id + "").then(successa, errora)
        //function successa(resulta) {
            
        //    var totTaskAvail = resulta.entities.length;
        //    if (totTaskAvail > 0) {
        //        Xrm.Navigation.openAlertDialog("Please clomplete all the Project Tasks to finish this Project.");
        //        Xrm.Page.getAttribute("itspsa_closeproject").setValue(null);
        //    }
            //debugger;

            //var id = Xrm.Page.data.entity.getId().slice(1, -1)
            //Xrm.WebApi.retrieveMultipleRecords("msdyn_projecttask", "?$filter=_msdyn_project_value eq " + id + " and itspsa_pmstatus eq 110920000").then(successb, errorb)
            //function successb(resultb) {
            //    //debugger;
            //    var TotTaskCom = resultb.entities.length
            //    if (totTaskAvail == TotTaskCom) {
            //        return;
            //    } else {
            //        Xrm.Navigation.openAlertDialog("Please clomplete all the Project Tasks to finish this Project.");
            //        //Xrm.Page.ui.setFormNotification("Please clomplete all the Project Tasks to finish this Project. ", "WARNING", "2")
            //        //setTimeout(function () { Xrm.Page.ui.clearFormNotification("2"); }, 5000);
            //        Xrm.Page.getAttribute("itspsa_closeproject").setValue(null);
            //    }
            //}
            //function errorb(resultb) {
            //    //debugger;
            //}
        //}
        //function errora(resulta) {
        //    //debugger;
        //}

    } else { return; }
}

//function OnCloseProjectCompleted() {

//    if (Xrm.Page.getAttribute("itspsa_closeproject").getValue() == 110920000) {
//        var id = Xrm.Page.data.entity.getId().slice(1, -1)
//        Xrm.WebApi.retrieveMultipleRecords("annotation", "?$filter=_objectid_value eq " + id + "").then(success, error)
//        function success(result) {
//            var annotationLength = result.entities.length
//            //debugger;
//            if (annotationLength == 0) {
//                Xrm.Page.getAttribute("itspsa_closeproject").setValue(null);
//                Xrm.Navigation.openAlertDialog("Please Upload the Required Documents. ");
//                //Xrm.Page.ui.setFormNotification("Please Upload the Required Documents. ", "WARNING","1")
//                // setTimeout( function () {Xrm.Page.ui.clearFormNotification("1");}, 5000 );
//                return;
//            }
//            var id = Xrm.Page.data.entity.getId().slice(1, -1)
//            Xrm.WebApi.retrieveMultipleRecords("msdyn_projecttask", "?$filter=_msdyn_project_value eq " + id + "").then(successa, errora)
//            function successa(resulta) {
//                var totTaskAvail = resulta.entities.length
//                //debugger;

//                var id = Xrm.Page.data.entity.getId().slice(1, -1)
//                Xrm.WebApi.retrieveMultipleRecords("msdyn_projecttask", "?$filter=_msdyn_project_value eq " + id + " and itspsa_pmstatus eq 110920000").then(successb, errorb)
//                function successb(resultb) {
//                    //debugger;
//                    var TotTaskCom = resultb.entities.length
//                    if ((annotationLength >= 1) && (totTaskAvail == TotTaskCom)) { return; } else {
//                        Xrm.Page.ui.setFormNotification("Please clomplete all the Project Tasks to finish this Project. ", "WARNING", "2")
//                        setTimeout(function () { Xrm.Page.ui.clearFormNotification("2"); }, 5000);
//                        Xrm.Page.getAttribute("itspsa_closeproject").setValue(null)
//                    }
//                }
//                function errorb(resultb) { debugger; }
//            }
//            function errora(resulta) { debugger; }
//        }
//        function error(result) { debugger; }

//    } else { return; }
//}

function changeRequestReqNullIfChangeReqNoRecord(executionContext) {
    //debugger;

    var formContext = executionContext.getFormContext();
    var ifYes = formContext.getAttribute('itspsa_changerequestisrequired').getValue();
    if (ifYes == 110920001) {
        var id = formContext.data.entity.getId().slice(1, -1);
        Xrm.WebApi.retrieveMultipleRecords("itspsa_changerequest", "?$filter=_itspsa_project_value eq (" + id + ")").then(successCallback, errorCallback);
        function successCallback(result) {
            //debugger;
            var lenRec = result.entities.length;
            if (lenRec == 0) {
                formContext.getAttribute('itspsa_changerequestisrequired').setValue(null);
                Xrm.Navigation.openAlertDialog("Please check 'Change Request'");
                //alert("Please check 'Change Request' ");
                //formContext.ui.setFormNotification("Please check 'Change Request Details' ", "WARNING", "1")
                //  setTimeout(function () { formContext.ui.clearFormNotification("1"); }, 5000);
            }
            else {
                return;
            }
        }
        function errorCallback(result) {
            // Handle error conditions
            Xrm.Utility.alertDialog(error.message, null);
        };
    }
    else {
        return;
    }
}


function projectProgress(executionContext) {
    var formContext = executionContext.getFormContext();
    var projectGUID = formContext.data.entity.getId();
    //Xrm.Utility.alertDialog(projectGUID);
    if (projectGUID != null) {
        //Xrm.WebApi.retrieveMultipleRecords("msdyn_projecttask", "?$filter=_msdyn_project_value eq (" + projectGUID + ")").then(successCallback, errorCallback);
        Xrm.WebApi.retrieveMultipleRecords("msdyn_resourceassignment", "?$filter=_msdyn_projectid_value eq (" + projectGUID + ")").then(successCallback, errorCallback);
        
    }
    else {
        //formContext.getAttribute('rak_finesamount').setValue(null);
    }

    function successCallback(result) {
        if (result.entities.length > 0) {
            var totalTasks = result.entities.length;
            Xrm.WebApi.retrieveMultipleRecords("msdyn_resourceassignment", "?$filter=_msdyn_projectid_value eq (" + projectGUID + ") and its_taskstatus eq 100000002").then(countSuccess, countError);
            function countSuccess(countResult) {
                var completedTasks = countResult.entities.length;
                var progress = (completedTasks / totalTasks) * 100;
                formContext.getAttribute("itspsa_progress").setValue(progress);
                formContext.data.entity.save();
                //formContext.getAttribute("itspsa_progress").setSubmitMode("always");
                //Xrm.Utility.alertDialog(progress);
            }
            function countError(countResult) {
                Xrm.Utility.alertDialog(error.message, null);
            }
        }
    }
    function errorCallback(result) {
        Xrm.Utility.alertDialog(error.message, null);
    };
}

function OnChangeUATDocument(executionContext) {
    //var formType = Xrm.Page.ui.getFormType();
    //if (formType != 1) {
    //    var value = Xrm.Page.getAttribute("itspsa_uploaduatdocument").getValue();
    //    if (value == 110920000) {

    //        var GUID = Xrm.Page.data.entity.getId();
    //        GUID = GUID.replace('{', '').replace('}', '');
    //        var url = "https://arcdev.crm4.dynamics.com/notes/edit.aspx?hideDesc=1&pId=%7b" + GUID + "%7d&pType=10132";

    //        var openUrlOptions = {
    //            height: 400,
    //            width: 800
    //        };

    //        Xrm.Navigation.openUrl(url, openUrlOptions);

    //    }
    //}

    var formContext = executionContext.getFormContext();
    var value = Xrm.Page.getAttribute("itspsa_uploaduatdocument").getValue();
    if (value == 110920000) {
        var id = formContext.data.entity.getId().slice(1, -1);
        var documentTypeId = "CAEACD31-C9DF-EB11-BACB-000D3ADDDDBD";
        Xrm.WebApi.retrieveMultipleRecords("itspsa_projectdocuments", "?$filter=_itspsa_project_value eq (" + id + ") and _itspsa_documenttype_value eq (" + documentTypeId + ")").then(successCallback, errorCallback);
        function successCallback(result) {
            var lenRec = result.entities.length;
            if (lenRec == 0) {
                formContext.getAttribute('itspsa_uploaduatdocument').setValue(null);
                Xrm.Navigation.openAlertDialog("Please do Upload UAT Document in Documents tab to proceed!");
            }
        }
        function errorCallback(result) {
            // Handle error conditions
            Xrm.Utility.alertDialog(error.message, null);
        }
    }
}

function OnSave(exectionObj) {
    var formType = Xrm.Page.ui.getFormType();
    if (formType != 1) {

        //var pocopyValue = Xrm.Page.getAttribute("itspsa_ispocopyuploaded").getValue();
        //var proposalValue = Xrm.Page.getAttribute("itspsa_isproposaluploaded").getValue();
        //if (pocopyValue == 1 || proposalValue == 1) {
        if (Xrm.Page.data.entity.getId() === null) return;

        var value = Xrm.Page.getAttribute("itspsa_uploaduatdocument").getValue();
        var projectStatus = Xrm.Page.getAttribute("itspsa_closeproject").getValue();
        if (value == 110920000 && (projectStatus == 110920002 || projectStatus == 110920001)) {

            var serverUrl = Xrm.Page.context.getClientUrl();
            var oDataSelect = serverUrl + "/XRMServices/2011/OrganizationData.svc/AnnotationSet?$filter=ObjectId/Id eq guid'" + Xrm.Page.data.entity.getId() + "'&$select=IsDocument";
            var allowSave = false;
            var req = new XMLHttpRequest();
            req.open("GET", oDataSelect, false);
            req.setRequestHeader("Accept", "application/json");
            req.setRequestHeader("Content-Type", "application/json;charset=utf-8");
            req.onreadystatechange = function () {
                if (req.readyState === 4) {
                    if (req.status === 200) {
                        var retrieved = JSON.parse(req.responseText).d;
                        for (var i = 0; i < retrieved.results.length; i++) {
                            if (retrieved.results.length >= 1) {
                                allowSave = true;
                                break;
                            }
                            //if (retrieved.results[0].IsDocument) {
                            //    allowSave = true;
                            //    break;
                            //}
                        }
                    }
                }
            };
            req.send();

            if (!allowSave) {

                alert("Please Upload the document to proceed!");
                //var GUID = Xrm.Page.data.entity.getId();
                //GUID = GUID.replace('{', '').replace('}', '');
                //var url = "https://arcdev.crm4.dynamics.com/notes/edit.aspx?hideDesc=1&pId=%7b" + GUID + "%7d&pType=10191";

                //var openUrlOptions = {
                //    height: 400,
                //    width: 800
                //};

                //Xrm.Navigation.openUrl(url, openUrlOptions);
                exectionObj.getEventArgs().preventDefault();
            }
            // }

        }
    }

}


function displayContract() {
    var value = Xrm.Page.getAttribute("itspsa_projectclassification").getValue();
    if (value == 110920002) {
        Xrm.Page.ui.tabs.get("Contracts").sections.get("Section").setVisible(true);
    }
    else {
        Xrm.Page.ui.tabs.get("Contracts").sections.get("Section").setVisible(false);
    }
}