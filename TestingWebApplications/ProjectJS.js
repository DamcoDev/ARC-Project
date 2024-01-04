
function displayContract() {
    var value = Xrm.Page.getAttribute("itspsa_projectclassification").getValue();
    if (value == 110920002) {
        Xrm.Page.ui.tabs.get("Contracts").sections.get("Section").setVisible(true);
    }
    else {
        Xrm.Page.ui.tabs.get("Contracts").sections.get("Section").setVisible(false);
    }
}