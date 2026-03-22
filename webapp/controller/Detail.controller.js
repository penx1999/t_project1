sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/routing/History",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/m/Text",
    "sap/m/Input",
    "sap/m/Label",
    "sap/ui/table/Column"
], function (Controller, History, JSONModel, Filter, FilterOperator, MessageBox, MessageToast, Text, Input, Label, UIColumn) {
    "use strict";

    var EDITABLE_FIELDS = [
        "ZZRFCUT",
        "PRODALLOCATIONACTIVATIONSTATUS",
        "PRODALLOCCHARCCONSTRAINTSTATUS",
        "PRODALLOCCHARCVALUECOMBNCMNT",
        "PRODUCTALLOCATIONQUANTITY"
    ];

    return Controller.extend("t_project1.controller.Detail", {

        _oOriginalData: null,
        _oFieldMetadata: null,

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
                hasChanges: false,
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

                    that._oFieldMetadata = {};

                    aFields.forEach(function (oField) {
                        var aData = (oField.DataSetAsoc && oField.DataSetAsoc.results) ? oField.DataSetAsoc.results : [];
                        aData.forEach(function (oEntry, iIdx) {
                            if (!aRows[iIdx].CHARCVALUECOMBINATIONUUID) {
                                aRows[iIdx].CHARCVALUECOMBINATIONUUID = oEntry.CHARCVALUECOMBINATIONUUID;
                            }
                            if (!aRows[iIdx].PRODALLOCPERDSTARTUTCDATETIME) {
                                aRows[iIdx].PRODALLOCPERDSTARTUTCDATETIME = oEntry.PRODALLOCPERDSTARTUTCDATETIME;
                            }
                            if (!aRows[iIdx].PRODALLOCPERIODENDUTCDATETIME) {
                                aRows[iIdx].PRODALLOCPERIODENDUTCDATETIME = oEntry.PRODALLOCPERIODENDUTCDATETIME;
                            }
                            if (!aRows[iIdx].prodallocationtimeseriesuuid) {
                                aRows[iIdx].prodallocationtimeseriesuuid = oEntry.prodallocationtimeseriesuuid;
                            }
                            aRows[iIdx][oField.name] = oEntry.Value;
                            aRows[iIdx][oField.name + "_old"] = oEntry.Value_old || oEntry.Value;

                            if (!that._oFieldMetadata[oField.name]) {
                                that._oFieldMetadata[oField.name] = {
                                    tabname: oEntry.tabname || oField.tabname || "",
                                    position: oEntry.position || oField.position || ""
                                };
                            }
                        });
                    });

                    that._oOriginalData = JSON.parse(JSON.stringify(aRows));

                    var sTitle = oBundle.getText("tableDataTitle") + " (" + aRows.length + ")";

                    oModel.setProperty("/columns", aColumns);
                    oModel.setProperty("/rows", aRows);
                    oModel.setProperty("/rowCount", aRows.length || 1);
                    oModel.setProperty("/tableTitle", sTitle);
                    oModel.setProperty("/hasChanges", false);
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

            var that = this;

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
                var oTemplate;
                if (bEditable) {
                    oTemplate = new Input({
                        value: "{detailModel>" + oCol.name + "}",
                        width: "100%",
                        change: that._onFieldChange.bind(that),
                        liveChange: that._onFieldChange.bind(that)
                    });
                } else {
                    oTemplate = new Text({ text: "{detailModel>" + oCol.name + "}", wrapping: false });
                }

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

            var oModel = this.getView().getModel("detailModel");
            oModel.attachPropertyChange(this._onModelPropertyChange, this);
        },

        _onModelPropertyChange: function (oEvent) {
            var sPath = oEvent.getParameter("path");
            if (sPath && sPath.indexOf("/rows") === 0) {
                var oModel = this.getView().getModel("detailModel");
                var aRows = oModel.getProperty("/rows");
                var bHasChanges = this._detectChanges(aRows);
                if (oModel.getProperty("/hasChanges") !== bHasChanges) {
                    oModel.setProperty("/hasChanges", bHasChanges);
                }
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

        _onFieldChange: function (oEvent) {
            jQuery.sap.log.info("Detail._onFieldChange triggered");
            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows");
            var bHasChanges = this._detectChanges(aRows);
            jQuery.sap.log.info("Detail._onFieldChange: hasChanges=" + bHasChanges);
            oModel.setProperty("/hasChanges", bHasChanges);
        },

        _detectChanges: function (aRows) {
            if (!this._oOriginalData || !aRows) {
                return false;
            }

            for (var i = 0; i < aRows.length; i++) {
                var oRow = aRows[i];
                var oOriginal = this._oOriginalData[i];
                if (!oOriginal) { continue; }

                for (var j = 0; j < EDITABLE_FIELDS.length; j++) {
                    var sField = EDITABLE_FIELDS[j];
                    if (oRow[sField] !== oOriginal[sField]) {
                        return true;
                    }
                }
            }
            return false;
        },

        _getChangedRows: function () {
            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows");
            var aChangedRows = [];

            if (!this._oOriginalData || !aRows) {
                return aChangedRows;
            }

            for (var i = 0; i < aRows.length; i++) {
                var oRow = aRows[i];
                var oOriginal = this._oOriginalData[i];
                if (!oOriginal) { continue; }

                var aChangedFields = [];
                for (var j = 0; j < EDITABLE_FIELDS.length; j++) {
                    var sField = EDITABLE_FIELDS[j];
                    if (oRow[sField] !== oOriginal[sField]) {
                        aChangedFields.push({
                            name: sField,
                            newValue: oRow[sField],
                            oldValue: oOriginal[sField]
                        });
                    }
                }

                if (aChangedFields.length > 0) {
                    aChangedRows.push({
                        rowIndex: i,
                        rowData: oRow,
                        originalData: oOriginal,
                        changedFields: aChangedFields
                    });
                }
            }

            return aChangedRows;
        },

        onSave: function () {
            var oModel = this.getView().getModel("detailModel");
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var aChangedRows = this._getChangedRows();

            if (aChangedRows.length === 0) {
                MessageToast.show(oBundle.getText("msgNoChanges"));
                return;
            }

            var sProductAllocationObject = oModel.getProperty("/productAllocationObject");
            var sFecIni = oModel.getProperty("/fec_ini");

            oModel.setProperty("/busy", true);
            MessageToast.show(oBundle.getText("msgSaving"));

            var that = this;
            var aPromises = [];

            aChangedRows.forEach(function (oChangedRow) {
                oChangedRow.changedFields.forEach(function (oChangedField) {
                    var oPayload = that._buildPayload(sProductAllocationObject, oChangedRow.rowData, oChangedField, sFecIni);
                    aPromises.push(that._executePut(oPayload));
                });
            });

            Promise.all(aPromises)
                .then(function () {
                    oModel.setProperty("/busy", false);
                    oModel.setProperty("/hasChanges", false);
                    that._oOriginalData = JSON.parse(JSON.stringify(oModel.getProperty("/rows")));
                    MessageToast.show(oBundle.getText("msgSaveSuccess"));
                })
                .catch(function (oError) {
                    oModel.setProperty("/busy", false);
                    var sMsg = oBundle.getText("msgSaveError");
                    try {
                        var oResp = JSON.parse(oError.responseText);
                        sMsg = oResp.error.message.value || sMsg;
                    } catch (e) {}
                    MessageBox.error(sMsg);
                });
        },

        _buildPayload: function (sKey, oRowData, oChangedField, sFecIni) {
            var oMetadata = this._oFieldMetadata[oChangedField.name] || {};

            return {
                key: sKey,
                tabname: oMetadata.tabname || "",
                name: oChangedField.name,
                Value: oChangedField.newValue,
                Value_old: oChangedField.oldValue,
                position: oMetadata.position || "",
                prodallocationtimeseriesuuid: oRowData.prodallocationtimeseriesuuid || "",
                productallocationobject: sKey,
                CHARCVALUECOMBINATIONUUID: oRowData.CHARCVALUECOMBINATIONUUID || "",
                PRODALLOCPERDSTARTUTCDATETIME: oRowData.PRODALLOCPERDSTARTUTCDATETIME || "",
                PRODALLOCPERIODENDUTCDATETIME: oRowData.PRODALLOCPERIODENDUTCDATETIME || "",
                fec_ini: this._toODataDate(sFecIni)
            };
        },

        _executePut: function (oPayload) {
            var that = this;
            var oODataModel = this.getOwnerComponent().getModel();
            var sPath = "/DynamicDataSet(key='" + encodeURIComponent(oPayload.key) + "')";

            return new Promise(function (resolve, reject) {
                oODataModel.update(sPath, oPayload, {
                    success: function () {
                        resolve();
                    },
                    error: function (oError) {
                        reject(oError);
                    }
                });
            });
        }

    });
});
