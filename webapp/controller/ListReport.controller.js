sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageToast",
    "sap/m/MessageBox",
    "sap/m/Dialog",
    "sap/m/SearchField",
    "sap/m/Text",
    "sap/m/Label",
    "sap/m/VBox",
    "sap/m/Button",
    "sap/m/Table",
    "sap/m/Column",
    "sap/m/ColumnListItem",
    "sap/ui/core/BusyIndicator"
], function (Controller, JSONModel, Filter, FilterOperator, MessageToast, MessageBox,
    Dialog, SearchField, Text, Label, VBox, Button, MTable, MColumn, ColumnListItem, BusyIndicator) {
    "use strict";

    return Controller.extend("t_project1.controller.ListReport", {

        onInit: function () {
            var oModel = new JSONModel({
                filterDivision: "",
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
            if (sPendingKey && sap.ushell && sap.ushell.Container) {
                sap.ushell.Container.getServiceAsync("CrossApplicationNavigation").then(function (oCrossAppNav) {
                    oCrossAppNav.getAppState(oOwnerComp, sPendingKey).done(function (oAppState) {
                        var oSaved = oAppState.getData();
                        if (oSaved && oSaved.listModel) {
                            oModel.setData(oSaved.listModel);
                        }
                        try {
                            window.sessionStorage.removeItem("zquot_pendingListAppStateKey");
                        } catch (e2) { /* ignore */ }
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
            var sDivision = (oModel.getProperty("/filterDivision") || "").trim();

            console.log("[ListReport] Botón 'Go' ejecutado.", {
                filterDivision: sDivision,
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

            aFilters.push(new Filter("DIVISION", FilterOperator.EQ, sDivision || "*"));

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
            oModel.setProperty("/filterDivision", "");
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

        onDivisionValueHelp: function (oEvent) {
            var oInput = oEvent.getSource();
            var that = this;

            var oVHModel = new JSONModel({
                allItems: [],
                items: [],
                displayedCount: 0,
                totalCount: 0,
                pageSize: 100,
                currentPage: 1,
                totalPages: 1,
                moreText: "[ 0 / 0 ]",
                canMore: false
            });
            var oDialog;
            var fnApplyValueHelpPage = function (iPage) {
                var aAllItems = oVHModel.getProperty("/allItems") || [];
                var iPageSize = oVHModel.getProperty("/pageSize") || 100;
                var iTotalPages = Math.max(Math.ceil(aAllItems.length / iPageSize), 1);
                var iCurrentPage = Math.min(Math.max(iPage || 1, 1), iTotalPages);
                var iDisplayCount = Math.min(iCurrentPage * iPageSize, aAllItems.length);
                var aPageItems = aAllItems.slice(0, iDisplayCount);

                oVHModel.setProperty("/items", aPageItems);
                oVHModel.setProperty("/displayedCount", aPageItems.length);
                oVHModel.setProperty("/totalCount", aAllItems.length);
                oVHModel.setProperty("/currentPage", iCurrentPage);
                oVHModel.setProperty("/totalPages", iTotalPages);
                oVHModel.setProperty("/moreText", "[ " + aPageItems.length + " / " + aAllItems.length + " ]");
                oVHModel.setProperty("/canMore", aPageItems.length < aAllItems.length);
            };
            var oSearchField = new SearchField({
                width: "100%",
                placeholder: "Search",
                search: function (oEv) {
                    var sQuery = oEv.getParameter("query") || oEv.getParameter("value") || "";
                    that._loadDivisionValueHelp(sQuery || "*", oVHModel, fnApplyValueHelpPage);
                }
            });
            var oValueHelpTable = new MTable({
                width: "100%",
                mode: "None",
                fixedLayout: true,
                columns: [
                    new MColumn({
                        width: "30rem",
                        header: new Label({ text: "Division" })
                    }),
                    new MColumn({
                        header: new Label({ text: "Description" })
                    })
                ],
                items: {
                    path: "/items",
                    template: new ColumnListItem({
                        type: "Active",
                        cells: [
                            new Text({ text: "{Clave}", wrapping: false }),
                            new Text({ text: "{Desc}", wrapping: false })
                        ],
                        press: function (oEv) {
                            var oRowContext = oEv.getSource().getBindingContext();
                            if (!oRowContext) { return; }
                            var sClave = oRowContext.getProperty("Clave");
                            oInput.setValue(sClave);
                            oDialog.close();
                        }
                    })
                }
            });

            oValueHelpTable.setModel(oVHModel);

            oDialog = new Dialog({
                title: "Search Help: Division",
                contentWidth: "80rem",
                contentHeight: "42rem",
                verticalScrolling: true,
                resizable: true,
                draggable: true,
                content: [
                    new VBox({
                        width: "100%",
                        items: [
                            oSearchField,
                            oValueHelpTable,
                            new VBox({
                                width: "100%",
                                alignItems: "Center",
                                items: [
                                    new Button({
                                        text: "More",
                                        type: "Transparent",
                                        enabled: "{/canMore}",
                                        press: function () {
                                            fnApplyValueHelpPage((oVHModel.getProperty("/currentPage") || 1) + 1);
                                        }
                                    }),
                                    new Text({ text: "{/moreText}" })
                                ]
                            })
                        ]
                    })
                ],
                endButton: new Button({
                    text: "Cancel",
                    press: function () { oDialog.close(); }
                }),
                afterClose: function () { oDialog.destroy(); }
            });

            oDialog.setModel(oVHModel);
            oDialog.open();
            this._loadDivisionValueHelp("*", oVHModel, fnApplyValueHelpPage);
        },

        _loadDivisionValueHelp: function (sSource, oVHModel, fnApplyValueHelpPage) {
            var oODataModel = this.getOwnerComponent().getModel();
            if (!oODataModel) { return; }
            var oModel = this.getView().getModel();
            var sAlloc = (oModel.getProperty("/filterProdAlloc") || "").trim() || "*";
            if (sAlloc.length > 300) {
                sAlloc = sAlloc.substring(0, 300);
            }
            var aFilters = [
                new Filter("source",           FilterOperator.EQ, sSource),
                new Filter("allocationObject", FilterOperator.EQ, sAlloc),
                new Filter("data_element",     FilterOperator.EQ, "SPART")
            ];
            console.log("[ValueHelp] GET /ValueHelpSet?$filter=source eq '" + sSource +
                "' and allocationObject eq '" + sAlloc + "' and data_element eq 'SPART'");
            BusyIndicator.show(0);
            var iStartTime = Date.now();
            oVHModel.__iReqSeq = (oVHModel.__iReqSeq || 0) + 1;
            var iReqSeq = oVHModel.__iReqSeq;
            oODataModel.read("/ValueHelpSet", {
                filters: aFilters,
                success: function (oData) {
                    console.log("ValueHelp OData response time ms:", Date.now() - iStartTime);
                    if (iReqSeq !== oVHModel.__iReqSeq) {
                        BusyIndicator.hide();
                        return;
                    }
                    var aItems = (oData && oData.results) ? oData.results : (oData ? [oData] : []);
                    console.log("ValueHelp OData records returned:", aItems.length);
                    oVHModel.setSizeLimit(Math.max(aItems.length, 100));
                    oVHModel.setProperty("/allItems", aItems);
                    if (fnApplyValueHelpPage) {
                        fnApplyValueHelpPage(1);
                    } else {
                        oVHModel.setProperty("/items", aItems);
                        oVHModel.setProperty("/displayedCount", aItems.length);
                        oVHModel.setProperty("/totalCount", aItems.length);
                    }
                    BusyIndicator.hide();
                },
                error: function (oErr) {
                    console.log("ValueHelp OData response time ms:", Date.now() - iStartTime);
                    if (iReqSeq !== oVHModel.__iReqSeq) {
                        BusyIndicator.hide();
                        return;
                    }
                    var sStatus = (oErr && oErr.statusCode) ? oErr.statusCode : "";
                    var sDetail = "";
                    try {
                        var oResp = JSON.parse(oErr.responseText);
                        sDetail = oResp.error.message.value || "";
                    } catch (e) {
                        sDetail = (oErr && oErr.message) ? oErr.message : "";
                    }
                    jQuery.sap.log.error("ValueHelp call failed (" + sStatus + "): " + sDetail);
                    console.error("[ValueHelp] OData error", sStatus, sDetail, oErr);
                    oVHModel.setProperty("/allItems", []);
                    oVHModel.setProperty("/items", []);
                    oVHModel.setProperty("/displayedCount", 0);
                    oVHModel.setProperty("/totalCount", 0);
                    oVHModel.setProperty("/currentPage", 1);
                    oVHModel.setProperty("/totalPages", 1);
                    oVHModel.setProperty("/moreText", "[ 0 / 0 ]");
                    oVHModel.setProperty("/canMore", false);
                    BusyIndicator.hide();
                }
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
