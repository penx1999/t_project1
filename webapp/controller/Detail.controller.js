sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/routing/History",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox",
    "sap/m/Column",
    "sap/m/Text",
    "sap/m/ColumnListItem",
    "sap/m/Label"
], function (Controller, History, JSONModel, Filter, FilterOperator, MessageBox, Column, Text, ColumnListItem, Label) {
    "use strict";

    return Controller.extend("t_project1.controller.Detail", {

        onInit: function () {
            var oModel = new JSONModel({
                productAllocationObject: "",
                columns: [],
                rows: [],
                busy: false
            });
            this.getView().setModel(oModel, "detailModel");

            var oRouter = this.getOwnerComponent().getRouter();
            oRouter.getRoute("RouteDetail").attachPatternMatched(this._onRouteMatched, this);
        },

        _onRouteMatched: function (oEvent) {
            var sQuotaId = decodeURIComponent(oEvent.getParameter("arguments").quotaId);
            var oModel = this.getView().getModel("detailModel");
            oModel.setProperty("/productAllocationObject", sQuotaId);
            oModel.setProperty("/busy", true);
            this._loadDynamicFields(sQuotaId);
        },

        _loadDynamicFields: function (sProductAllocationObject) {
            var oODataModel = this.getOwnerComponent().getModel();
            var oModel = this.getView().getModel("detailModel");
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var that = this;

            oODataModel.read("/DynamicFieldSet", {
                filters: [new Filter("tablename", FilterOperator.EQ, sProductAllocationObject)],
                urlParameters: { "$expand": "DataSetAsoc" },
                success: function (oData) {
                    var aFields = oData.results || [];
                    aFields.sort(function (a, b) {
                        return parseInt(a.position) - parseInt(b.position);
                    });

                    var aColumns = aFields.map(function (oField) {
                        return {
                            name: oField.name,
                            label: oField.description || oField.name
                        };
                    });

                    var mRows = {};
                    aFields.forEach(function (oField) {
                        var aData = (oField.DataSetAsoc && oField.DataSetAsoc.results) ? oField.DataSetAsoc.results : [];
                        aData.forEach(function (oEntry) {
                            var sUUID = oEntry.CHARCVALUECOMBINATIONUUID;
                            if (!mRows[sUUID]) {
                                mRows[sUUID] = {
                                    CHARCVALUECOMBINATIONUUID: sUUID,
                                    PRODALLOCPERDSTARTUTCDATETIME: oEntry.PRODALLOCPERDSTARTUTCDATETIME,
                                    PRODALLOCPERIODENDUTCDATETIME: oEntry.PRODALLOCPERIODENDUTCDATETIME
                                };
                            }
                            mRows[sUUID][oField.name] = oEntry.Value;
                        });
                    });

                    var aRows = Object.values(mRows);

                    oModel.setProperty("/columns", aColumns);
                    oModel.setProperty("/rows", aRows);
                    oModel.setProperty("/busy", false);

                    that._buildTable(aColumns);
                },
                error: function (oError) {
                    oModel.setProperty("/busy", false);
                    var sMsg = oBundle.getText("msgReadError");
                    try {
                        var oResp = JSON.parse(oError.responseText);
                        sMsg = oResp.error.message.value || sMsg;
                    } catch (e) {}
                    MessageBox.error(sMsg);
                }
            });
        },

        _buildTable: function (aColumns) {
            var oTable = this.byId("idDynamicTable");
            if (!oTable) { return; }

            oTable.destroyColumns();

            oTable.addColumn(new Column({ header: new Label({ text: "Período Inicio" }) }));
            oTable.addColumn(new Column({ header: new Label({ text: "Período Fin" }) }));

            aColumns.forEach(function (oCol) {
                oTable.addColumn(new Column({ header: new Label({ text: oCol.label }) }));
            });

            var oTemplate = oTable.getBindingInfo("items") && oTable.getBindingInfo("items").template;
            if (oTemplate) { oTemplate.destroy(); }

            var oCells = [];
            oCells.push(new Text({ text: "{detailModel>PRODALLOCPERDSTARTUTCDATETIME}" }));
            oCells.push(new Text({ text: "{detailModel>PRODALLOCPERIODENDUTCDATETIME}" }));
            aColumns.forEach(function (oCol) {
                oCells.push(new Text({ text: "{detailModel>" + oCol.name + "}" }));
            });

            oTable.bindItems({
                path: "detailModel>/rows",
                template: new ColumnListItem({ cells: oCells })
            });
        },

        onNavBack: function () {
            var oHistory = History.getInstance();
            var sPreviousHash = oHistory.getPreviousHash();
            if (sPreviousHash !== undefined) {
                window.history.go(-1);
            } else {
                this.getOwnerComponent().getRouter().navTo("RouteListReport", {}, true);
            }
        }

    });
});
