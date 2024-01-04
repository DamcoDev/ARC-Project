function setSendforApprovalReadOnly() {
    var optionValue = Xrm.Page.getAttribute("itspsa_sendforapproval").getValue();
    if (optionValue == 1) {
        Xrm.Page.getControl("itspsa_sendforapproval").setDisabled(true);
    }
}

function mandatoryFields() {
    var formType = Xrm.Page.ui.getFormType();
    if (formType == 1) {
        // Xrm.Page.getAttribute("itspsa_proposedprojectstartdate").setRequiredLevel("none");
        //Xrm.Page.getAttribute("itspsa_proposedprojectenddate").setRequiredLevel("none");
        Xrm.Page.getAttribute("itspsa_povalue").setRequiredLevel("none");
        //Xrm.Page.getAttribute("itspsa_amcvalue").setRequiredLevel("none");
        Xrm.Page.getAttribute("itspsa_proposedcostexternal").setRequiredLevel("none");
        Xrm.Page.getAttribute("itspsa_proposedcostinternal").setRequiredLevel("none");
        //Xrm.Page.getAttribute("itspsa_amcduration").setRequiredLevel("none");
    }
    else {
        var ProjectClassification = Xrm.Page.getAttribute("itspsa_projectclassification").getValue();
        if (ProjectClassification == 960760000)
        {
            Xrm.Page.getAttribute("itspsa_povalue").setRequiredLevel("required");
            //Xrm.Page.getAttribute("itspsa_amcvalue").setRequiredLevel("required");
            Xrm.Page.getAttribute("itspsa_proposedcostexternal").setRequiredLevel("required");
            Xrm.Page.getAttribute("itspsa_proposedcostinternal").setRequiredLevel("none");
        }
        else
        {
            //Xrm.Page.getAttribute("itspsa_proposedprojectstartdate").setRequiredLevel("required");
            // Xrm.Page.getAttribute("itspsa_proposedprojectenddate").setRequiredLevel("required");
            Xrm.Page.getAttribute("itspsa_povalue").setRequiredLevel("required");
            //Xrm.Page.getAttribute("itspsa_amcvalue").setRequiredLevel("required");
            Xrm.Page.getAttribute("itspsa_proposedcostexternal").setRequiredLevel("required");
            Xrm.Page.getAttribute("itspsa_proposedcostinternal").setRequiredLevel("required");
            //Xrm.Page.getAttribute("itspsa_amcduration").setRequiredLevel("required");
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
                        if (retrieved.results.length >= 2) {
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

            alert("Please Upload Proposal and PO Copy");
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

function OnChangeProposalCopy() {
    var formType = Xrm.Page.ui.getFormType();
    if (formType != 1) {
        var value = Xrm.Page.getAttribute("itspsa_isproposaluploaded").getValue();
        if (value == 1) {

            var GUID = Xrm.Page.data.entity.getId();
            GUID = GUID.replace('{', '').replace('}', '');
            var url = "https://arcdev.crm4.dynamics.com/notes/edit.aspx?hideDesc=1&pId=%7b" + GUID + "%7d&pType=10191";

            var openUrlOptions = {
                height: 400,
                width: 800
            };

            Xrm.Navigation.openUrl(url, openUrlOptions);

        }
    }
}

function OnChangePOCopy() {
    var formType = Xrm.Page.ui.getFormType();
    if (formType != 1) {
        var value = Xrm.Page.getAttribute("itspsa_ispocopyuploaded").getValue();
        if (value == 1) {

            var GUID = Xrm.Page.data.entity.getId();
            GUID = GUID.replace('{', '').replace('}', '');
            var url = "https://arcdev.crm4.dynamics.com/notes/edit.aspx?hideDesc=1&pId=%7b" + GUID + "%7d&pType=10191";

            var openUrlOptions = {
                height: 400,
                width: 800
            };

            Xrm.Navigation.openUrl(url, openUrlOptions);

        }
    }
}


// JavaScript source code

// JavaScript source code

function retrieveTotCurrency() {
    //debugger;
    var poVal = Xrm.Page.getAttribute('itspsa_povalue').getValue();
    if (poVal == null) {
        var idCheck = Xrm.Page.getAttribute('itspsa_opportunity').getValue();
        if (idCheck != null) {
            var id = Xrm.Page.getAttribute('itspsa_opportunity').getValue()[0].id.slice(1, -1);

            Xrm.WebApi.retrieveMultipleRecords("emitac_commercialproposal", "?$filter=_emitac_opportunityid_value eq (" + id + ")").then(successCallback, errorCallback);
        } else {
            Xrm.Page.getAttribute('itspsa_povalue').setValue(null);
            Xrm.Page.getAttribute("itspsa_povalue").setSubmitMode("always");
        }

        function successCallback(result) {
            //debugger;
            var tSellingValue = 0;
            for (var i = 0; i < result.entities.length; i++) {

                var totalSellingValue = result.entities[i].arc_totalsellingvalueincvat;
                {
                    tSellingValue = tSellingValue + totalSellingValue
                }
            }
            Xrm.Page.getAttribute('itspsa_povalue').setValue(tSellingValue);
            Xrm.Page.getAttribute("itspsa_povalue").setSubmitMode("always");
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

// JavaScript source code

function retrieveTotGP() {
    //debugger;
    var gpVal = Xrm.Page.getAttribute('itspsa_gp').getValue();
    if (gpVal == null) {
        var idCheck = Xrm.Page.getAttribute('itspsa_opportunity').getValue();
        if (idCheck != null) {
            var id = Xrm.Page.getAttribute('itspsa_opportunity').getValue()[0].id.slice(1, -1);

            Xrm.WebApi.retrieveMultipleRecords("emitac_commercialproposal", "?$filter=_emitac_opportunityid_value eq (" + id + ")").then(successCallback, errorCallback);
        } else {
            Xrm.Page.getAttribute('itspsa_gp').setValue(null);
            Xrm.Page.getAttribute("itspsa_gp").setSubmitMode("always");
        }

        function successCallback(result) {
            //debugger;
            var tGPValue = 0;
            for (var i = 0; i < result.entities.length; i++) {

                var totalGP = result.entities[i].emitac_total;
                {
                    tGPValue = tGPValue + totalGP
                }
            }
            Xrm.Page.getAttribute('itspsa_gp').setValue(tGPValue);
            Xrm.Page.getAttribute("itspsa_gp").setSubmitMode("always");
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

//Team Lead Lookup value populate

function callingTeamLeadLookup() {
    var idCheck = Xrm.Page.getAttribute('itspsa_departmentlineofbusiness').getValue();
    if (idCheck != null) {
        var id = Xrm.Page.getAttribute('itspsa_departmentlineofbusiness').getValue()[0].id.slice(1, -1);
        Xrm.WebApi.retrieveRecord("emitac_lineofbusinesses", id).then(success, error)
        var lineManagerId = "";
        function success(result) {
            //debugger;
            lineManagerId = result._itspsa_linemanager_value;
            {
                Xrm.WebApi.retrieveRecord("systemuser", lineManagerId).then(success1, error1)
                function success1(result1) {
                    //debugger;

                    if (result1.systemuserid != null) {
                        var teamLead = new Array();
                        teamLead[0] = new Object();
                        teamLead[0].id = result1.systemuserid;
                        teamLead[0].name = result1.fullname;
                        teamLead[0].entityType = "systemuser";
                        Xrm.Page.getAttribute("itspsa_teamlead").setValue(teamLead);
                        Xrm.Page.getAttribute("itspsa_teamlead").setSubmitMode("always");
                    }

                }
                function error1(result) {
                    console.log(error.message);
                    // handle error conditions
                }
            }
        }
        function error(result) {
            console.log(error.message);
            // handle error conditions
        }

    } else {
        Xrm.Page.getAttribute("itspsa_teamlead").setValue(null);
        Xrm.Page.getAttribute("itspsa_teamlead").setSubmitMode("always");
    }
}

var checkDept = "";
function randNum() {
    //debugger;
    //var value=Xrm.Page.data.entity.getId().slice(1, -29);
    var formType = Xrm.Page.ui.getFormType();
    if (formType == 1) {
        checkDept = Xrm.Page.getAttribute('itspsa_departmentlineofbusiness').getValue();
        var value = "ARC-PN" + Math.floor(Math.random() * 1000001);
        if (checkDept != null) {
            var deptID = checkDept[0].id;
            if (deptID == "{4E9C1273-F3A0-E711-810E-5065F38AA961}") {
                Xrm.Page.getAttribute('itspsa_projectnumber').setValue(value.toString());
                Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(true);
                Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("required");
            } else {
                //debugger;
                Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(false);
                Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("none");
                //Xrm.Page.getAttribute('itspsa_projectnumber').setValue(null);
            }
        } else {
            //debugger;
            Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(false);
            Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("none");
            //Xrm.Page.getAttribute('itspsa_projectnumber').setValue(null);
        }
    }
}

// Validate PO Value
function validatePOValueExt() {

    var POValue = Xrm.Page.getAttribute('itspsa_povalue').getValue();
    var ExtProposal = Xrm.Page.getAttribute('itspsa_proposedcostexternal').getValue();
    var IntProposal = Xrm.Page.getAttribute('itspsa_proposedcostinternal').getValue();
    var tot = ExtProposal + IntProposal - 0.01
    if (tot >= POValue) {
        alert("Sum of Proposal Values must be less than PO Value");
        // clean the field
        Xrm.Page.getAttribute("itspsa_proposedcostexternal").setValue(null);
    }
    else { return; }
}

function validatePOValueInt() {

    var POValue = Xrm.Page.getAttribute('itspsa_povalue').getValue();
    var ExtProposal = Xrm.Page.getAttribute('itspsa_proposedcostexternal').getValue();
    var IntProposal = Xrm.Page.getAttribute('itspsa_proposedcostinternal').getValue();
    var tot = ExtProposal + IntProposal - 0.01
    if (tot >= POValue) {
        alert("Sum of Proposal Values must be less than PO Value");
        // clean the field
        Xrm.Page.getAttribute("itspsa_proposedcostinternal").setValue(null);
    }
    else { return; }
}

function projectStartDate() {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var dateValue = Xrm.Page.getAttribute("itspsa_proposedprojectstartdate").getValue();
    if (dateValue != null) {
        dateValue.setHours(0, 0, 0, 0);
        if (dateValue < today) {
            alert("Proposed Project Start Date cannot be Past Date");
            // clean the field
            Xrm.Page.getAttribute("itspsa_proposedprojectstartdate").setValue(null);
        }
        else {
            return;
        }
    }
}

function projectEndDate() {
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var dateValue = Xrm.Page.getAttribute("itspsa_proposedprojectenddate").getValue();
    if (dateValue != null) {
        dateValue.setHours(0, 0, 0, 0);
        if (dateValue < today) {
            alert("Proposed Project End Date cannot be Past Date");
            // clean the field
            Xrm.Page.getAttribute("itspsa_proposedprojectenddate").setValue(null);
        }
        else {
            return;
        }
    }
}

//Approval status is Approved
function disableFormFields() {
    //debugger;
    var approvalCheck = Xrm.Page.getAttribute('itspsa_approvalstatus');
    if (approvalCheck != null) {
        var approvalVal = Xrm.Page.getAttribute('itspsa_approvalstatus').getValue();
        if (approvalVal == 2) {

            Xrm.Page.ui.controls.forEach(function (control, index) {

                var controlType = control.getControlType();

                if (controlType != "iframe" && controlType != "webresource" && controlType != "subgrid") {

                    control.setDisabled(true);

                }

            });

        }
        else { return; }
    }
}




// Requirement, Account_parentaccountid_value, Account Manager_emitac_clientmanager_value, LineOfBusiness _arc_primarylineofbusiness_value , ((PO and GP) Other code)

function onChangeOfOpportunity() {
    var idCheck = Xrm.Page.getAttribute('itspsa_opportunity').getValue();
    //For Requirement
    if (idCheck != null) {
        {
            var id = Xrm.Page.getAttribute('itspsa_opportunity').getValue()[0].id.slice(1, -1);
            Xrm.WebApi.retrieveRecord("opportunity", id).then(success1, error1)
            function success1(result1) {
                //debugger;
                Xrm.Page.getAttribute('itspsa_name').setValue(result1.name);

            }
            function error1(result1) {
                console.log(error.message);
                // handle error conditions
            }
        }

        //For account
        var idCheck = Xrm.Page.getAttribute('itspsa_opportunity').getValue();
        var id = Xrm.Page.getAttribute('itspsa_opportunity').getValue()[0].id.slice(1, -1);
        Xrm.WebApi.retrieveRecord("opportunity", id).then(success2, error2)
        function success2(result2) {
            //debugger;
            accountId = result2._parentaccountid_value;
            {
                Xrm.WebApi.retrieveRecord("account", accountId).then(success2a, error2a)
                function success2a(result2a) {
                    //debugger;

                    if (result2a.accountid != null) {
                        var teamLead = new Array();
                        teamLead[0] = new Object();
                        teamLead[0].id = result2a.accountid;
                        teamLead[0].name = result2a.name;
                        teamLead[0].entityType = "account";
                        Xrm.Page.getAttribute("itspsa_customer").setValue(teamLead);
                    }

                }
                function error2a(result2a) {
                    console.log(error.message);
                    // handle error conditions
                }
            }
        }
        function error2(result2) {
            console.log(error.message);
            // handle error conditions
        }


        //For account Manager


        var idCheck = Xrm.Page.getAttribute('itspsa_opportunity').getValue();
        var id = Xrm.Page.getAttribute('itspsa_opportunity').getValue()[0].id.slice(1, -1);
        Xrm.WebApi.retrieveRecord("opportunity", id).then(success3, error3)
        function success3(result3) {
            //debugger;
            accountManId = result3._emitac_clientmanager_value;
            {
                Xrm.WebApi.retrieveRecord("systemuser", accountManId).then(success3a, error3a)
                function success3a(result3a) {
                    //debugger;


                    var teamLead = new Array();
                    teamLead[0] = new Object();
                    teamLead[0].id = result3a.systemuserid;
                    teamLead[0].name = result3a.fullname;
                    teamLead[0].entityType = "systemuser";
                    Xrm.Page.getAttribute("itspsa_accountmanager").setValue(teamLead);


                }
                function error3a(result3a) {
                    console.log(error.message);
                    // handle error conditions
                }
            }
        }
        function error3(result3) {
            console.log(error.message);
            // handle error conditions
        }


        // for Line of Bussiness


        var id = Xrm.Page.getAttribute('itspsa_opportunity').getValue()[0].id.slice(1, -1);
        Xrm.WebApi.retrieveRecord("opportunity", id).then(success4, error4)

        function success4(result4) {
            //debugger;
            var lineManagerId = result4._arc_primarylineofbusiness_value;
            {
                Xrm.WebApi.retrieveRecord("emitac_lineofbusiness", lineManagerId).then(success4a, error4a)
                function success4a(result4a) {
                    //debugger;

                    if (result4._arc_primarylineofbusiness_value != null) {
                        var teamLead = new Array();
                        teamLead[0] = new Object();
                        teamLead[0].id = result4._arc_primarylineofbusiness_value;
                        teamLead[0].name = result4a.emitac_name;
                        teamLead[0].entityType = "emitac_lineofbusiness";
                        Xrm.Page.getAttribute("itspsa_departmentlineofbusiness").setValue(teamLead);
                        checkDept = Xrm.Page.getAttribute('itspsa_departmentlineofbusiness').getValue();
                        var value = "ARC-PN" + Math.floor(Math.random() * 1000001);
                        if (checkDept != null) {
                            var deptID = checkDept[0].id.slice(1, -1).toLowerCase();
                            if (deptID == "4e9c1273-f3a0-e711-810e-5065f38aa961") {
                                Xrm.Page.getAttribute('itspsa_projectnumber').setValue(value.toString());
                                Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(true);
                                Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("required");
                            } else {
                                //debugger;
                                Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(false);
                                Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("none");
                                //Xrm.Page.getAttribute('itspsa_projectnumber').setValue(null);
                            }
                        } else {
                            //debugger;
                            Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(false);
                            Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("none");
                            //Xrm.Page.getAttribute('itspsa_projectnumber').setValue(null);
                        }
                    }

                }
                function error4a(result4a) {
                    console.log(error.message);
                    // handle error conditions
                }
            }
        }

        function error4(result4) {
            console.log(error.message);
            // handle error conditions

        }
    }
        //end of if statements
    else {
        Xrm.Page.getAttribute('itspsa_name').setValue(null);
        Xrm.Page.getAttribute("itspsa_customer").setValue(null);
        Xrm.Page.getAttribute("itspsa_accountmanager").setValue(null);
        Xrm.Page.getAttribute("itspsa_departmentlineofbusiness").setValue(null);
        Xrm.Page.getControl('itspsa_additionalsonumbers').setVisible(false);
        Xrm.Page.getAttribute("itspsa_additionalsonumbers").setRequiredLevel("none");
        //Xrm.Page.getAttribute('itspsa_projectnumber').setValue(null);

    }
}

function gpZero() {
    //debugger;
    var gpVal = Xrm.Page.getAttribute('itspsa_gp').getValue()
    if (gpVal == null) { Xrm.Page.getAttribute('itspsa_gp').setValue(0) }
}

function poZero() {
    //debugger;
    var poVal = Xrm.Page.getAttribute('itspsa_povalue').getValue()
    if (poVal == null) { Xrm.Page.getAttribute('itspsa_povalue').setValue(0) }
}



