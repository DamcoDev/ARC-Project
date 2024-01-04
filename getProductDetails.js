function getProductDetails() {
    try {
        var lookup = new Array();
        lookup = Xrm.Page.getAttribute("emitac_product").getValue();
        if (lookup != null) {
            if (lookup[0].id != null) {
                var id = lookup[0].id;
                var name = lookup[0].name;
                Xrm.WebApi.retrieveRecord("product", id, "?$select=productnumber")
                    .then(function (data) {
                        retrieveContactSuccess(data);
                    });
            }
        }

        else {
            Xrm.Page.getAttribute("emitac_product").setValue(null);
        }
    }
    catch (e) {
        Xrm.Utility.alertDialog(e.message);
    }
}

function retrieveContactSuccess(data) {

    try {
        var vendorProductNumber = data["productnumber"];
        Xrm.Page.getAttribute("arc_vendorpartnumber").setValue(vendorProductNumber);
    }
    catch (e) {
        Xrm.Utility.alertDialog(e.message);
    }
}