sap.ui.define([
    "sap/base/util/ObjectPath",
    "sap/ushell/services/Container"
], function (ObjectPath) {
    "use strict";

    var oFlpSandbox = {
        init: function () {
            ObjectPath.set(["sap-ushell-config"], {
                defaultRenderer: "fiori2",
                applications: {
                    "tproject1-display": {
                        title: "Gestión de Cuotas de Producto",
                        description: "ZQUOT V1",
                        additionalInformation: "SAPUI5.Component=t_project1",
                        applicationType: "URL",
                        url: "../",
                        navigationMode: "embedded"
                    }
                }
            });
        }
    };

    oFlpSandbox.init();
});
