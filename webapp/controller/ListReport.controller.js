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
                filterDescription: "",
                filterAllocationObject: "",
                QuotaResults: [],
                detailEnabled: false,
                selectedItems: []
            });
            this.getView().setModel(oModel);

            var oOwnerComp = this.getOwnerComponent();
            oOwnerComp._oListModel = oModel;

            var sPendingKey = null;
            try {
                sPendingKey = window.sessionStorage.getItem("zquot_pendingListAppStateKey");
            } catch (e) { /* ignore */ }
            console.log("[ListReport] onInit - sPendingKey desde sessionStorage:", sPendingKey, "| sap.ushell disponible:", !!(sap.ushell && sap.ushell.Container));
            if (sPendingKey && sap.ushell && sap.ushell.Container) {
                sap.ushell.Container.getServiceAsync("CrossApplicationNavigation").then(function (oCrossAppNav) {
                    console.log("[ListReport] CrossApplicationNavigation obtenido, llamando getAppState con key:", sPendingKey);
                    oCrossAppNav.getAppState(oOwnerComp, sPendingKey).done(function (oAppState) {
                        var oSaved = oAppState.getData();
                        console.log("[ListReport] getAppState resuelto. oSaved:", oSaved);
                        if (oSaved && oSaved.listModel) {
                            oModel.setData(oSaved.listModel);
                            console.log("[ListReport] Modelo de lista restaurado con exito.");
                        } else {
                            console.warn("[ListReport] oSaved.listModel no existe, no se restauro nada.");
                        }
                        try {
                            window.sessionStorage.removeItem("zquot_pendingListAppStateKey");
                        } catch (e2) { /* ignore */ }
                    }).fail(function () {
                        console.error("[ListReport] getAppState fallo (posiblemente la key expiro o es invalida):", sPendingKey);
                    });
                });
            }

            this.getOwnerComponent().getRouter()
                .getRoute("RouteListReport")
                .attachPatternMatched(this._onRouteMatched, this);

            var that = this;
            this.getView().addEventDelegate({
                onkeydown: function (oEvent) {
                    if (oEvent.key === "Enter" || oEvent.keyCode === 13) {
                        var sTag = oEvent.target ? oEvent.target.tagName.toUpperCase() : "";
                        if (sTag !== "BUTTON") {
                            that.onSearch();
                        }
                    }
                }
            });
        },

        _onRouteMatched: function () {
            var oOwner = this.getOwnerComponent();
            var oDetailModel = oOwner._oDetailModel;
            if (oDetailModel && oDetailModel.getProperty("/hasChanges")) {
                var sQuotaId = oDetailModel.getProperty("/productAllocationObject");
                oOwner._bPreventDetailReload = true;
                oOwner.getRouter().navTo("RouteDetail", {
                    quotaId: encodeURIComponent(sQuotaId)
                });
            }
        },

        onAllocationObjectLiveChange: function (oEvent) {
            var oInput = oEvent.getSource();
            var sValue = oInput.getValue();
            var sClean = sValue.replace(/[^a-zA-Z0-9-]/g, "").toUpperCase();
            if (sClean !== sValue) {
                oInput.setValue(sClean);
            }
        },

        onSearch: function () {
            var oModel = this.getView().getModel();
            var sProdAlloc = (oModel.getProperty("/filterProdAlloc") || "").trim();
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var sAllocationObject = (oModel.getProperty("/filterAllocationObject") || "").trim();

            console.log("[ListReport] Botón 'Go' ejecutado.", {
                filterProdAlloc: sProdAlloc,
                filterDescription: (oModel.getProperty("/filterDescription") || "").trim(),
                filterAllocationObject: sAllocationObject
            });

            if (sAllocationObject) {
                console.log("[ListReport] Allocation Object con valor: navegando directo a pantalla 2 y ejecutando su Go.");
                this._navigateToDetail({ PRODUCTALLOCATIONOBJECT: sProdAlloc || "-" }, sAllocationObject);
                return;
            }

            var oODataModel = this.getOwnerComponent().getModel();
            var aFilters = [];

            aFilters.push(new Filter("PRODUCTALLOCATIONOBJECT", FilterOperator.EQ, sProdAlloc || "*"));

            var sDescription = (oModel.getProperty("/filterDescription") || "").trim();
            aFilters.push(new Filter("DESCRIPTION", FilterOperator.EQ, sDescription || "*"));

            aFilters.push(new Filter("DATA_ELEMENT", FilterOperator.EQ, sAllocationObject || "*"));

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
            oModel.setProperty("/filterDescription", "");
            oModel.setProperty("/filterAllocationObject", "");
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

        _navigateToDetail: function (oItem, sAllocationObjectFilter) {
            var sId = encodeURIComponent(oItem.PRODUCTALLOCATIONOBJECT || oItem.DESCRIPTION);

            if (!this.getOwnerComponent().getModel("detailModel")) {
                this.getOwnerComponent().setModel(new JSONModel(oItem), "detailModel");
            } else {
                this.getOwnerComponent().getModel("detailModel").setData(oItem);
            }

            if (sAllocationObjectFilter) {
                this.getOwnerComponent()._sPendingAllocationObjectFilter = sAllocationObjectFilter;
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
