function customerDuplicateCheck() {
    var lookup = new Array();
    lookup = Xrm.Page.getAttribute("itscs_customer").getValue();
    if (lookup != null) {
        var customerID = lookup[0].id;
        customerID = customerID.slice(1, -1).toLowerCase();
        var fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='false'><entity name='itscs_casecreationforspecificcustomers'><attribute name='itscs_name'/>  <attribute name='itscs_customer' /></entity ></fetch > ";
        //Xrm.WebApi.retrieveMultipleRecords("itscs_casecreationforspecificcustomers", "?fetchXml= " + fetchXml).then(
        Xrm.WebApi.retrieveMultipleRecords("itscs_casecreationforspecificcustomers", "?$select=itscs_name,_itscs_customer_value").then(
            function success(result) {
                for (var i = 0; i < result.entities.length; i++) {
                    var CustomerIDResult = result.entities[i]["_itscs_customer_value"];
                    if (customerID == CustomerIDResult) {
                        Xrm.Utility.alertDialog("Existing Configuration is already presented with this Customer.");
                        Xrm.Page.getAttribute("itscs_customer").setValue(null);
                    }
                }
            },
            function (error) {
                Xrm.Utility.alertDialog(error.message, function () { return; });
            }
        );
    }
 
}