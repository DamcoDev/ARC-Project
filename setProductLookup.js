function setProductLookup() {
    var vendorProductNumber = Xrm.Page.getAttribute("arc_vendorpartnumber").getValue();
    var fetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' aggregate='false'><entity name='product'><attribute name='productnumber'/><attribute name='name'/></entity ></fetch > ";
    Xrm.WebApi.retrieveMultipleRecords("product", "?fetchXml= " + fetchXml).then(
        function success(result) {
            for (var i = 0; i < result.entities.length; i++) {
                if (vendorProductNumber == result.entities[i]["productnumber"]) {
                    var setLooup = new Array();
                    setLooup[0] = new Object();
                    setLooup[0].id = result.entities[i]["productid"];
                    setLooup[0].name = result.entities[i]["name"];
                    setLooup[0].entityType = "product";
                    Xrm.Page.getAttribute("emitac_product").setValue(setLooup);
                }
            }
        },
        function (error) {
            Xrm.Utility.alertDialog(error.message, function () { return; });
        }
    );
}