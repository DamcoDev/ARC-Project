﻿function escalation() {
    var caseOwnerorEngineer = Xrm.Page.getAttribute("itscs_caseownerorengineer").getValue();
    var currentDateTime = new Date();
    var severity = Xrm.Page.getAttribute("severitycode").getValue();
    //Escalation Level-1
    if (caseOwnerorEngineer != null) {
        if (severity == 1) {
            var fourhour = currentDateTime.setHours(currentDateTime.getHours() + 4);
            Xrm.Page.getAttribute("itscs_timestampofassignmenttoengineer").setValue(fourhour);
        }
        else if (severity == 2) {
            var twelvehour = currentDateTime.setHours(currentDateTime.getHours() + 12);
            Xrm.Page.getAttribute("itscs_timestampofassignmenttoengineer").setValue(twelvehour);
        }
        else if (severity == 3) {
            var twentyhour = currentDateTime.setHours(currentDateTime.getHours() + 20);
            Xrm.Page.getAttribute("itscs_timestampofassignmenttoengineer").setValue(twentyhour);
        }
    }
    //Escalation Level-2
    if (caseOwnerorEngineer != null) {
        if (severity == 1) {
            var sixhour = currentDateTime.setHours(currentDateTime.getHours() + 6);
            Xrm.Page.getAttribute("itscs_timestampafter1stlevelescalation").setValue(sixhour);
        }
        else if (severity == 2) {
            var fourteenhour = currentDateTime.setHours(currentDateTime.getHours() + 14);
            Xrm.Page.getAttribute("itscs_timestampafter1stlevelescalation").setValue(fourteenhour);
        }
        else if (severity == 3) {
            var twentytwohour = currentDateTime.setHours(currentDateTime.getHours() + 22);
            Xrm.Page.getAttribute("itscs_timestampafter1stlevelescalation").setValue(twentytwohour);
        }
    }
    //Escalation Level-3
    if (caseOwnerorEngineer != null) {
        if (severity == 1) {
            var eighthour = currentDateTime.setHours(currentDateTime.getHours() + 8);
            Xrm.Page.getAttribute("itscs_timestampafter2ndlevelescalation").setValue(eighthour);
        }
        else if (severity == 2) {
            var sixteenhour = currentDateTime.setHours(currentDateTime.getHours() + 16);
            Xrm.Page.getAttribute("itscs_timestampafter2ndlevelescalation").setValue(sixteenhour);
        }
        else if (severity == 3) {
            var twentyfourhour = currentDateTime.setHours(currentDateTime.getHours() + 24);
            Xrm.Page.getAttribute("itscs_timestampafter2ndlevelescalation").setValue(twentyfourhour);
        }
    }

    //Escalation Level-4
    if (caseOwnerorEngineer != null) {
        if (severity == 1) {
            var fourteenhour = currentDateTime.setHours(currentDateTime.getHours() + 14);
            Xrm.Page.getAttribute("itscs_timestampafter3rdlevelescalation").setValue(fourteenhour);
        }
        else if (severity == 2) {
            var twentytwohour = currentDateTime.setHours(currentDateTime.getHours() + 22);
            Xrm.Page.getAttribute("itscs_timestampafter3rdlevelescalation").setValue(twentytwohour);
        }
        else if (severity == 3) {
            var thirtyhour = currentDateTime.setHours(currentDateTime.getHours() + 30);
            Xrm.Page.getAttribute("itscs_timestampafter3rdlevelescalation").setValue(thirtyhour);
        }
    }
}

function AssignDate() {
    var date = new Date();
    Xrm.Page.getAttribute("itscs_timestampofassignmenttoengineer").setValue(date);
}

function checkResolvingUser() {
    var status = Xrm.Page.getAttribute("statuscode").getValue();
    var caseOwner = Xrm.Page.getAttribute("ownerid").getValue();
    var caseOwnerId = caseOwner[0].id;
    caseOwnerId = caseOwnerId.slice(1, -1);

    var helpDeskUser = Xrm.Page.getAttribute("itscs_helpdeskuser").getValue();
    if (helpDeskUser != null) {
        var helpDeskUserId = helpDeskUser[0].id.slice(1, -1);
        if (status == 173020007 && caseOwnerId != helpDeskUserId) {
            alert("You are not authorized to resolve the case");
            Xrm.Page.getAttribute("statuscode").setValue(173020005);
        }
    }
}

function changeTaskStatusNullIfTimeEntriesNoRecord(executionContext) {
    //debugger;
    var formContext = executionContext.getFormContext();
    var ifCompleted = formContext.getAttribute('statuscode').getValue();
    if (ifCompleted == 173020005) {
        var id = formContext.data.entity.getId().slice(1, -1);
        Xrm.WebApi.retrieveMultipleRecords("msdyn_timeentry", "?$filter=_itspsa_case_value eq (" + id + ")").then(successCallback, errorCallback);
        function successCallback(result) {
            //debugger;
            var lenRec = result.entities.length;
            if (lenRec == 0) {
                //Xrm.Navigation.openAlertDialog("Please do enter Time Sheets");

                formContext.getAttribute('statuscode').setValue(1);
                alert("Please do enter Time Sheet to close the Ticket!");
            }
            else {
                return;
            }
        }

        function errorCallback(result) {
            // Handle error conditions
            Xrm.Utility.alertDialog(error.message, null);
        };
    } else {
        return
    }
}

function checkTaskStatusforClosingCase(executionContext) {
    //debugger;
    var formContext = executionContext.getFormContext();
    var ifCompleted = formContext.getAttribute('statuscode').getValue();
    if (ifCompleted == 173020007) {
        var id = formContext.data.entity.getId().slice(1, -1);
        Xrm.WebApi.retrieveMultipleRecords("msdyn_timeentry", "?$filter=_itspsa_case_value eq (" + id + ") and msdyn_entrystatus ne 192350002").then(successCallback, errorCallback);
        function successCallback(result) {
            //debugger;
            var lenRec = result.entities.length;
            if (lenRec >= 1) {
                //Xrm.Navigation.openAlertDialog("Please do enter Time Sheets");

                formContext.getAttribute('statuscode').setValue(1);
                alert("Associated Time Entries are not approved to close the Case!");
            }
            else {
                return;
            }
        }

        function errorCallback(result) {
            // Handle error conditions
            Xrm.Utility.alertDialog(error.message, null);
        };
    } else {
        return
    }
}