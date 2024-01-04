//Registered on Onchange of GP%
function validateGPPercentage(executionContext) {
    var formContext = executionContext.getFormContext();
    var GPPercentage = formContext.getAttribute("its_gp").getValue();
    var OpportunityGUID = formContext.getAttribute('regardingobjectid').getValue();
    if (OpportunityGUID != null) {
        var OpportunityId = formContext.getAttribute("regardingobjectid").getValue()[0].id.slice(1, -1);
        Xrm.WebApi.retrieveMultipleRecords("its_presalesrequest", "?$filter=_regardingobjectid_value eq (" + OpportunityId + ")").then(successCallback, errorCallback);
        var tfeesValue = 0;
        function successCallback(CompanyResult) {
            for (var i = 0; i < CompanyResult.entities.length; i++) {
                var feesValue = CompanyResult.entities[i].its_gp;
                {
                    tfeesValue = tfeesValue + feesValue
                }
            }
            var balancePercentage = 100 - tfeesValue;
            if (GPPercentage > balancePercentage)
            {
                formContext.getAttribute('its_gp').setValue(null);
                Xrm.Navigation.openAlertDialog({ confirmButtonLabel: "OK", text: "You are only allowed to give " + balancePercentage + "." });
            }

        }
        function errorCallback(CompanyResult) {
            Xrm.Utility.alertDialog(error.message, null);
        };
    }
}