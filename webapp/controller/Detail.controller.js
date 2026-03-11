sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/routing/History",
    "sap/m/MessageToast"
], function (Controller, History, MessageToast) {
    "use strict";

    return Controller.extend("t_project1.controller.Detail", {

        onInit: function () {
            var oRouter = this.getOwnerComponent().getRouter();
            oRouter.getRoute("RouteDetail").attachPatternMatched(this._onRouteMatched, this);
        },

        _onRouteMatched: function () {
            var oDetailModel = this.getOwnerComponent().getModel("detailModel");
            if (oDetailModel) {
                this.getView().setModel(oDetailModel, "detailModel");
            }
        },

        onNavBack: function () {
            var oHistory = History.getInstance();
            var sPreviousHash = oHistory.getPreviousHash();

            if (sPreviousHash !== undefined) {
                window.history.go(-1);
            } else {
                this.getOwnerComponent().getRouter().navTo("RouteListReport", {}, true);
            }
        },

        formatActivationStatus: function (sStatus) {
            switch (sStatus) {
                case "1": return "Success";
                case "2": return "None";
                case "3": return "Error";
                default:  return "None";
            }
        },

        formatDate: function (oDate) {
            if (!oDate) { return ""; }
            try {
                var oDateFormat = sap.ui.core.format.DateFormat.getDateInstance({ pattern: "dd/MM/yyyy" });
                return oDateFormat.format(new Date(oDate));
            } catch (e) {
                return oDate;
            }
        }

    });
});
