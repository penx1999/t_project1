sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageToast",
    "sap/m/MessageBox"
], function (Controller, JSONModel, Filter, FilterOperator, MessageToast, MessageBox) {
    "use strict";

    return Controller.extend("t_project1.controller.ListReport", {

        onInit: function () {
            var oModel = new JSONModel({
                filterProdAlloc: "",
                QuotaResults: [],
                detailEnabled: false,
                selectedItems: []
            });
            this.getView().setModel(oModel);
        },

        onSearch: function () {
            var oModel = this.getView().getModel();
            var sProdAlloc = oModel.getProperty("/filterProdAlloc");
            var oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (!sProdAlloc) {
                MessageBox.warning(oBundle.getText("msgProdAllocRequired"));
                return;
            }

            var oODataModel = this.getOwnerComponent().getModel();
            var aFilters = [];

            aFilters.push(new Filter("Productallocationobject", FilterOperator.EQ, sProdAlloc));

            var that = this;

            oODataModel.read("/PROD_ALLOCSet", {
                filters: aFilters,
                success: function (oData) {
                    var aResults = oData.results || [];
                    oModel.setProperty("/QuotaResults", aResults);
                    oModel.setProperty("/detailEnabled", false);
                    oModel.setProperty("/selectedItems", []);

                    var sCountText = aResults.length + " " + oBundle.getText("records");
                    var oCountText = that.byId("idRecordCount");
                    if (oCountText) { oCountText.setText(sCountText); }
                    var oSnappedCount = that.byId("idSnappedCount");
                    if (oSnappedCount) { oSnappedCount.setText(sCountText); }

                    if (aResults.length === 0) {
                        MessageToast.show(oBundle.getText("msgNoRecords"));
                    }
                },
                error: function (oError) {
                    var sMsg = oBundle.getText("msgReadError");
                    try {
                        var oResp = JSON.parse(oError.responseText);
                        sMsg = oResp.error.message.value || sMsg;
                    } catch (e) { }
                    MessageBox.error(sMsg);
                }
            });
        },

        onClear: function () {
            var oModel = this.getView().getModel();
            oModel.setProperty("/filterProdAlloc", "");
            oModel.setProperty("/QuotaResults", []);
            oModel.setProperty("/detailEnabled", false);
            oModel.setProperty("/selectedItems", []);

            var oCountText = this.byId("idRecordCount");
            if (oCountText) { oCountText.setText(""); }
            var oSnappedCount = this.byId("idSnappedCount");
            if (oSnappedCount) { oSnappedCount.setText(""); }
        },

        onSelectionChange: function () {
            var oTable = this.byId("idQuotaTable");
            var aSelectedItems = oTable.getSelectedItems();
            var oModel = this.getView().getModel();

            oModel.setProperty("/detailEnabled", aSelectedItems.length === 1);

            var aSelected = aSelectedItems.map(function (oItem) {
                return oItem.getBindingContext().getObject();
            });
            oModel.setProperty("/selectedItems", aSelected);
        },

        onItemPress: function (oEvent) {
            var oContext = oEvent.getSource().getBindingContext();
            if (oContext) {
                this._navigateToDetail(oContext.getObject());
            }
        },

        onNavToDetail: function () {
            var oModel = this.getView().getModel();
            var aSelected = oModel.getProperty("/selectedItems");
            if (aSelected && aSelected.length === 1) {
                this._navigateToDetail(aSelected[0]);
            }
        },

        _navigateToDetail: function (oItem) {
            var sId = encodeURIComponent(oItem.Productallocationobject);

            if (!this.getOwnerComponent().getModel("detailModel")) {
                this.getOwnerComponent().setModel(new JSONModel(oItem), "detailModel");
            } else {
                this.getOwnerComponent().getModel("detailModel").setData(oItem);
            }

            this.getOwnerComponent().getRouter().navTo("RouteDetail", {
                quotaId: sId
            });
        },

        onExport: function () {
            var oModel = this.getView().getModel();
            var aResults = oModel.getProperty("/QuotaResults");
            var oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (!aResults || aResults.length === 0) {
                MessageToast.show(oBundle.getText("msgNoDataExport"));
                return;
            }

            MessageToast.show(oBundle.getText("msgExportPending"));
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
