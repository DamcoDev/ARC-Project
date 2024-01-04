//Registered on Change of Email ID
function createContacts() {
    var InputEmails = Xrm.Page.getAttribute("its_emailids").getValue();
    var EmailArray = InputEmails.split(";");
    var Flag = true;
    for (var i = 0; i < EmailArray.length; i++) {
        if (!EmailTest(EmailArray[i])) {
            Flag = false;
        }
        if (Flag != true) {
            alert("The list of emails entered contain invalid email format. Please re-enter");
            Xrm.Page.getAttribute("its_emailids").setValue(null);
            Xrm.Page.getControl("its_emailids").setFocus();
        }
    }

}
function EmailTest(EmailField) {
    var Email = /^([a-zA-Z0-9_.-])+@([a-zA-Z0-9_.-])+\.([a-zA-Z])+([a-zA-Z])+/;
    if (Email.test(EmailField)) {
        Xrm.WebApi.retrieveMultipleRecords("contact", "?$filter=emailaddress1 eq ('" + EmailField + "')").then(successCallback, errorCallback);
        function successCallback(result) {
            if (result.entities.length > 0) {
            }
            else {
                var emailName = EmailField.split("@");
                var contactObj = null;
                contactObj = new Object();
                //contactObj.firstname = "Mike";
                contactObj.lastname = emailName[0];
                contactObj.emailaddress1 = EmailField;
                Xrm.WebApi.createRecord("contact", contactObj).then(function (ContactResult) {
                    //get the guid of created record
                    var recordId = ContactResult.id;

                    //below code is used to open the created record
                    var windowOptions = {
                        openInNewWindow: true
                    };
                    //check if XRM.Utility is not null
                    if (Xrm.Utility != null) {

                        //open the entity record
                        Xrm.Utility.openEntityForm("contact", recordId, null, windowOptions);
                        var contactLookup = new Array();
                        contactLookup[0] = new Object();
                        contactLookup[0].id = recordId;
                        contactLookup[0].name = emailName[0];
                        contactLookup[0].entityType = "contact";
                        Xrm.Page.getAttribute("cc").setValue(contactLookup);
                    }
                },
                    function (error) {
                        Xrm.Utility.alertDialog(error.message);
                    });
            }
        }
        function errorCallback(result) {
            Xrm.Utility.alertDialog(error.message, null);
        };
        return true;
    }
    else {
        return false;
    }
}
