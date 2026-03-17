sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/routing/History",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox",
    "sap/m/Text",
    "sap/m/Input",
    "sap/m/Label",
    "sap/ui/table/Column"
], function (Controller, History, JSONModel, Filter, FilterOperator, MessageBox, Text, Input, Label, UIColumn) {
    "use strict";

    var EDITABLE_FIELDS = [
        "ZZRFCUT",
        "PRODALLOCATIONACTIVATIONSTATUS",
        "PRODALLOCCHARCCONSTRAINTSTATUS",
        "PRODALLOCCHARCVALUECOMBNCMNT",
        "PRODUCTALLOCATIONQUANTITY"
    ];

    return Controller.extend("t_project1.controller.Detail", {

        onInit: function () {
            var oToday = new Date();
            var oNextYear = new Date();
            oNextYear.setFullYear(oNextYear.getFullYear() + 1);

            var oModel = new JSONModel({
                productAllocationObject: "",
                tableTitle: "",
                columns: [],
                rows: [],
                rowCount: 5,
                busy: false,
                fec_ini: this._formatDateValue(oToday),
                fec_fin: this._formatDateValue(oNextYear)
            });
            this.getView().setModel(oModel, "detailModel");

            var oRouter = this.getOwnerComponent().getRouter();
            oRouter.getRoute("RouteDetail").attachPatternMatched(this._onRouteMatched, this);
        },

        _formatDateValue: function (oDate) {
            var sDay   = ("0" + oDate.getDate()).slice(-2);
            var sMon   = ("0" + (oDate.getMonth() + 1)).slice(-2);
            var sYear  = oDate.getFullYear();
            return sDay + "/" + sMon + "/" + sYear;
        },

        _parseDateValue: function (sDate) {
            var aParts = sDate.split("/");
            return new Date(parseInt(aParts[2]), parseInt(aParts[1]) - 1, parseInt(aParts[0]));
        },

        _toODataDate: function (sDate) {
            if (!sDate) { return ""; }
            var aParts = sDate.split("/");
            return aParts[2] + aParts[1] + aParts[0];
        },

        _onRouteMatched: function (oEvent) {
            var sQuotaId = decodeURIComponent(oEvent.getParameter("arguments").quotaId);
            var oModel = this.getView().getModel("detailModel");
            oModel.setProperty("/productAllocationObject", sQuotaId);
            oModel.setProperty("/busy", true);
            this._loadDynamicFields(sQuotaId);
        },

        onDateChange: function () {
            var oModel = this.getView().getModel("detailModel");
            var sQuotaId = oModel.getProperty("/productAllocationObject");
            if (sQuotaId) {
                oModel.setProperty("/busy", true);
                this._loadDynamicFields(sQuotaId);
            }
        },

        _loadDynamicFields: function (sProductAllocationObject) {
            var oODataModel = this.getOwnerComponent().getModel();
            var oModel = this.getView().getModel("detailModel");
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var that = this;

            var sFecIni = oModel.getProperty("/fec_ini");
            var sFecFin = oModel.getProperty("/fec_fin");

            var aFilters = [
                new Filter("tablename", FilterOperator.EQ, sProductAllocationObject)
            ];

            if (sFecIni) {
                aFilters.push(new Filter("fec_ini", FilterOperator.EQ, this._toODataDate(sFecIni)));
            }
            if (sFecFin) {
                aFilters.push(new Filter("fec_fin", FilterOperator.EQ, this._toODataDate(sFecFin)));
            }

            oODataModel.read("/DynamicFieldSet", {
                filters: aFilters,
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

                    var iMaxRows = 0;
                    aFields.forEach(function (oField) {
                        var aData = (oField.DataSetAsoc && oField.DataSetAsoc.results) ? oField.DataSetAsoc.results : [];
                        if (aData.length > iMaxRows) { iMaxRows = aData.length; }
                    });

                    var aRows = [];
                    for (var i = 0; i < iMaxRows; i++) { aRows.push({}); }

                    aFields.forEach(function (oField) {
                        var aData = (oField.DataSetAsoc && oField.DataSetAsoc.results) ? oField.DataSetAsoc.results : [];
                        aData.forEach(function (oEntry, iIdx) {
                            if (!aRows[iIdx].CHARCVALUECOMBINATIONUUID) {
                                aRows[iIdx].CHARCVALUECOMBINATIONUUID = oEntry.CHARCVALUECOMBINATIONUUID;
                            }
                            aRows[iIdx][oField.name] = oEntry.Value;
                        });
                    });
                    var sTitle = oBundle.getText("tableDataTitle") + " (" + aRows.length + ")";

                    oModel.setProperty("/columns", aColumns);
                    oModel.setProperty("/rows", aRows);
                    oModel.setProperty("/rowCount", aRows.length || 1);
                    oModel.setProperty("/tableTitle", sTitle);
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

            var sColWidth = "150px";

            var iStatusIdx = -1;
            aColumns.forEach(function (oCol, iIdx) {
                if (iStatusIdx === -1 && oCol.name.toUpperCase().indexOf("STATUS") !== -1) {
                    iStatusIdx = iIdx;
                }
            });

            var iFixedCount = (iStatusIdx > 0) ? iStatusIdx : 0;

            jQuery.sap.log.info("Detail._buildTable: columns=" + JSON.stringify(aColumns.map(function(c){ return c.name; })) + " | iStatusIdx=" + iStatusIdx + " | iFixedCount=" + iFixedCount);

            aColumns.forEach(function (oCol) {
                var bEditable = EDITABLE_FIELDS.indexOf(oCol.name.toUpperCase()) !== -1 ||
                                EDITABLE_FIELDS.indexOf(oCol.name) !== -1;
                var oTemplate = bEditable
                    ? new Input({ value: "{detailModel>" + oCol.name + "}", width: "100%" })
                    : new Text({ text: "{detailModel>" + oCol.name + "}", wrapping: false });

                oTable.addColumn(new UIColumn({
                    width: sColWidth,
                    label: new Label({ text: oCol.label, wrapping: false }),
                    template: oTemplate,
                    resizable: true,
                    autoResizable: true
                }));
            });

            oTable.setFixedColumnCount(iFixedCount);
            oTable.bindRows("detailModel>/rows");
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
