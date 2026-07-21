sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/routing/History",
    "sap/ui/model/json/JSONModel",
    "sap/ui/model/Filter",
    "sap/ui/model/FilterOperator",
    "sap/ui/core/CustomData",
    "sap/m/MessageBox",
    "sap/m/MessageToast",
    "sap/m/Text",
    "sap/m/Input",
    "sap/m/Label",
    "sap/m/DatePicker",
    "sap/m/Dialog",
    "sap/m/SearchField",
    "sap/m/VBox",
    "sap/m/HBox",
    "sap/m/Button",
    "sap/m/Table",
    "sap/m/Column",
    "sap/m/ColumnListItem",
    "sap/ui/table/Column",
    "sap/ui/table/RowSettings",
    "sap/ui/core/format/DateFormat",
    "sap/ui/core/BusyIndicator"
], function (Controller, History, JSONModel, Filter, FilterOperator, CustomData, MessageBox, MessageToast, Text, Input, Label, DatePicker, Dialog, SearchField, VBox, HBox, Button, MTable, MColumn, ColumnListItem, UIColumn, RowSettings, DateFormat, BusyIndicator) {
    "use strict";

    var EDITABLE_FIELDS = [
        "ZZRFCUT",
        "PRODALLOCCHARCVALUECOMBNCMNT",
        "PRODUCTALLOCATIONQUANTITY",
        "PRODALLOCPERDSTARTUTCDATE",
        "PRODALLOCPERIODENDUTCDATE",
        "PRODALLOCATIONACTIVATIONSTATUS",
        "PRODALLOCCHARCCONSTRAINTSTATUS"
    ];

    var NON_EDITABLE_FIELDS = [
        "PRODUCTALLOCATIONOBJECT",
        "PRODALLOCATIONACTIVATIONSTATUS",
        "PRODALLOCCHARCCONSTRAINTSTATUS"
    ];

    return Controller.extend("t_project1.controller.Detail", {

        _oOriginalData: null,
        _oFieldMetadata: null,
        _hasDeletedRows: false,
        _aDeletedRows: [],

        onInit: function () {
            var oToday = new Date();
            var oFirstOfMonth = new Date(oToday.getFullYear(), oToday.getMonth(), 1);
            var oNextYear = new Date(oFirstOfMonth.getFullYear() + 1, oFirstOfMonth.getMonth(), 0);

            var oModel = new JSONModel({
                productAllocationObject: "",
                allocationObjectFilter: "",
                l_key_char: "",
                tableTitle: "",
                columns: [],
                rows: [],
                rowCount: 5,
                busy: false,
                hasChanges: false,
                editMode: false,
                fec_ini: this._formatDateValue(oFirstOfMonth),
                fec_fin: this._formatDateValue(oNextYear),
                messageText: "",
                messageType: "None",
                messageVisible: false
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
            var oOwner = this.getOwnerComponent();
            if (oOwner._bPreventDetailReload) {
                oOwner._bPreventDetailReload = false;
                var that = this;
                this._showUnsavedPopup(function () {
                    var oDetailModel = that.getView().getModel("detailModel");
                    oDetailModel.setProperty("/hasChanges", false);
                    oDetailModel.setProperty("/messageVisible", false);
                    oDetailModel.setProperty("/messageText", "");
                    oDetailModel.setProperty("/messageType", "None");
                    that.getOwnerComponent().getRouter().navTo("RouteListReport", {}, true);
                });
                return;
            }
            var oModel = this.getView().getModel("detailModel");
            oOwner._oDetailModel = oModel;

            var oCompDetailModel = oOwner.getModel("detailModel");
            var sKeyChar = oCompDetailModel ? (oCompDetailModel.getProperty("/PRODUCTALLOCATIONOBJECT") || "") : "";
            oModel.setProperty("/l_key_char", sKeyChar);

            oModel.setProperty("/productAllocationObject", sQuotaId);
            oModel.setProperty("/busy", true);
            oModel.setProperty("/messageVisible", false);
            oModel.setProperty("/messageText", "");
            oModel.setProperty("/messageType", "None");
            oModel.setProperty("/editMode", false);
            this._loadDynamicFields(sQuotaId);
        },

        onEdit: function () {
            var oModel = this.getView().getModel("detailModel");
            var sQuotaId = oModel.getProperty("/productAllocationObject");
            oModel.setProperty("/messageVisible", false);
            oModel.setProperty("/messageText", "");
            oModel.setProperty("/messageType", "None");
            if (sQuotaId) {
                oModel.setProperty("/busy", true);
                this._loadDynamicFields(sQuotaId, function () {
                    oModel.setProperty("/editMode", true);
                }, "EDIT");
            } else {
                oModel.setProperty("/editMode", true);
            }
        },

        onDateChange: function () {
            var oModel = this.getView().getModel("detailModel");
            var sQuotaId = oModel.getProperty("/productAllocationObject");
            var sFilterValue = oModel.getProperty("/allocationObjectFilter") || "";
            if (sQuotaId) {
                oModel.setProperty("/busy", true);
                this._loadDynamicFields(sQuotaId, null, sFilterValue);
            }
        },

        onAllocationObjectChange: function () {
            var oModel = this.getView().getModel("detailModel");
            var sFilterValue = (oModel.getProperty("/allocationObjectFilter") || "").trim();
            oModel.setProperty("/allocationObjectFilter", sFilterValue);

            var sQuotaId = oModel.getProperty("/productAllocationObject");
            if (sQuotaId) {
                oModel.setProperty("/busy", true);
                this._loadDynamicFields(sQuotaId, null, sFilterValue);
            }
        },

        _loadDynamicFields: function (sProductAllocationObject, fnAfterSuccess, sType) {
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
            if (sType) {
                aFilters.push(new Filter("data_element", FilterOperator.EQ, sType));
            }

            console.log("[DynamicTable] Ejecutando OData /DynamicFieldSet", {
                productAllocationObject: sProductAllocationObject,
                fec_ini: sFecIni || "",
                fec_fin: sFecFin || "",
                data_element: sType || ""
            });
            oODataModel.read("/DynamicFieldSet", {
                filters: aFilters,
                urlParameters: { "$expand": "DataSetAsoc" },
                success: function (oData) {
                    var aFields = oData.results || [];
                    aFields.sort(function (a, b) {
                        return parseInt(a.position) - parseInt(b.position);
                    });

                    var bEn = that._getSapLang() === "en";
                    var aColumns = aFields.map(function (oField) {
                        var sLabel = oField.description || oField.name;
                        if (bEn && (oField.name || "").toUpperCase() === "PRODALLOCCHARCVALUECOMBNCMNT") {
                            sLabel = "Comment";
                        }
                        return {
                            name: oField.name,
                            label: sLabel
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

                    that._oCellKeys = {};

                    // Seed per-field metadata from the DynamicField level (data_element here applies to all cells of the column)
                    aFields.forEach(function (oField) {
                        that._oFieldMetadata[oField.name] = {
                            tabname: oField.tablename || "",
                            position: oField.position || "",
                            data_element: oField.data_element || ""
                        };
                    });

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
                            if (!aRows[iIdx].productallocationsequence) {
                                aRows[iIdx].productallocationsequence = oEntry.productallocationsequence || "";
                            }
                            var sValue = (oEntry.Value || "").toString().trim();
                            var sValueOld = (oEntry.Value_old || oEntry.Value || "").toString().trim();
                            aRows[iIdx][oField.name] = sValue;
                            aRows[iIdx][oField.name + "_old"] = sValueOld;

                            var sCellKey = iIdx + "_" + oField.name;
                            that._oCellKeys[sCellKey] = {
                                key: oEntry.key,
                                tabname: oEntry.tabname || oField.tablename || "",
                                position: oField.position || "",
                                prodallocationtimeseriesuuid: oEntry.prodallocationtimeseriesuuid,
                                productallocationobject: oEntry.productallocationobject,
                                CHARCVALUECOMBINATIONUUID: oEntry.CHARCVALUECOMBINATIONUUID,
                                PRODALLOCPERDSTARTUTCDATETIME: oEntry.PRODALLOCPERDSTARTUTCDATETIME,
                                PRODALLOCPERIODENDUTCDATETIME: oEntry.PRODALLOCPERIODENDUTCDATETIME,
                                productallocationsequence: oEntry.productallocationsequence || "",
                                ind_ope: oEntry.ind_ope || "",
                                data_element: oEntry.data_element || ""
                            };

                            if (!that._oFieldMetadata[oField.name]) {
                                that._oFieldMetadata[oField.name] = {
                                    tabname: oEntry.tabname || oField.tablename || "",
                                    position: oField.position || "",
                                    data_element: oEntry.data_element || ""
                                };
                            } else if (!that._oFieldMetadata[oField.name].data_element && oEntry.data_element) {
                                that._oFieldMetadata[oField.name].data_element = oEntry.data_element;
                            }
                        });
                    });

                    var oVarCharCol = { name: "VAR_CHAR", label: "Var_CHAR" };
                    var iVarCharIdx = -1;
                    aColumns.forEach(function (oCol, iIdx) {
                        if (iVarCharIdx === -1 && oCol.name.toUpperCase() === "PRODUCTALLOCATIONOBJECT") {
                            iVarCharIdx = iIdx;
                        }
                    });
                    if (iVarCharIdx === -1) { iVarCharIdx = aColumns.length - 1; }
                    aColumns.splice(iVarCharIdx + 1, 0, oVarCharCol);

                    var sKeyCharValue = oModel.getProperty("/l_key_char") || "";
                    var oKeyCharCol = { name: "KEY_CHAR", label: "Key_Char" };
                    aColumns.splice(iVarCharIdx + 2, 0, oKeyCharCol);

                    aRows.forEach(function (oRow) {
                        oRow.VAR_CHAR = sProductAllocationObject || "";
                        oRow.VAR_CHAR_old = sProductAllocationObject || "";
                        oRow.KEY_CHAR = sKeyCharValue;
                        oRow.KEY_CHAR_old = sKeyCharValue;
                    });

                    that._oOriginalData = JSON.parse(JSON.stringify(aRows));
                    that._hasDeletedRows = false;
                    that._aDeletedRows = [];

                    var sTitle = oBundle.getText("tableDataTitle") + " (" + aRows.length + ")";

                    oModel.setProperty("/columns", aColumns);
                    oModel.setProperty("/rows", aRows);
                    oModel.setProperty("/rowCount", aRows.length || 1);
                    oModel.setProperty("/tableTitle", sTitle);
                    oModel.setProperty("/hasChanges", false);
                    oModel.setProperty("/editMode", false);
                    oModel.setProperty("/busy", false);

                    that._buildTable(aColumns);
                    if (typeof fnAfterSuccess === "function") {
                        fnAfterSuccess();
                    }
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
            var aVisibleColumns = aColumns.filter(function (oCol) {
                var sName = (oCol.name || "").toUpperCase();
                if (sName.indexOf("PROD") === 0 && sName.lastIndexOf("DESC") === sName.length - 4) { return false; }
                if (sName === "PRODUCTALLOCATIONOBJECTUUID" || sName === "VAR_CHAR" || sName === "KEY_CHAR") { return false; }
                return true;
            });

            oTable.destroyColumns();

            var sColWidth = "150px";

            var iStatusIdx = -1;
            aVisibleColumns.forEach(function (oCol, iIdx) {
                if (iStatusIdx === -1 && oCol.name.toUpperCase().indexOf("STATUS") !== -1) {
                    iStatusIdx = iIdx;
                }
            });

            var iFixedCount = (iStatusIdx > 0) ? iStatusIdx : 0;

            jQuery.sap.log.info("Detail._buildTable: columns=" + JSON.stringify(aVisibleColumns.map(function(c){ return c.name; })) + " | iStatusIdx=" + iStatusIdx + " | iFixedCount=" + iFixedCount);

            var oDateFormat = DateFormat.getDateInstance({
                style: "medium"
            });

            var sConsumedQtyField = "";
            aVisibleColumns.forEach(function (oCol) {
                if ((oCol.label || "").toUpperCase().trim() === "CNSMD QTY") {
                    sConsumedQtyField = oCol.name;
                }
            });

            aVisibleColumns.forEach(function (oCol) {
                var sFieldName = oCol.name;
                var sFieldUpper = sFieldName.toUpperCase();

                var bNonEditableText = false;
                
                var sLabelUpper = (oCol.label || "").toUpperCase().trim();
                var bNeverEditable = (sLabelUpper === "AVBL QTY" || sLabelUpper === "CNSMD QTY" ||
                                      sFieldUpper === "PRODUCTALLOCATIONOBJECT");

                var bNonEditableInput = (sFieldUpper === "PRODALLOCATIONACTIVATIONSTATUS" ||
                                         sFieldUpper === "PRODALLOCCHARCCONSTRAINTSTATUS");
                
                var bEditableField = (sFieldUpper === "ZZRFCUT" ||
                                      sFieldUpper === "PRODALLOCCHARCVALUECOMBNCMNT" ||
                                      sFieldUpper === "PRODUCTALLOCATIONQUANTITY");

                var bDateField = (sFieldUpper === "PRODALLOCPERDSTARTUTCDATE" ||
                                  sFieldUpper === "PRODALLOCPERIODENDUTCDATE");

                var bIsComment = (sFieldUpper === "PRODALLOCCHARCVALUECOMBNCMNT");
                var bIsRequired = !bNonEditableText && !bNeverEditable && !bNonEditableInput && !bIsComment;

                var oTemplate;
                if (bNonEditableText) {
                    oTemplate = new Text({ text: "{detailModel>" + sFieldName + "}", wrapping: false });
                } else if (bNeverEditable) {
                    oTemplate = new Input({
                        value: "{detailModel>" + sFieldName + "}",
                        editable: false
                    }).addStyleClass("sapUiSizeCompact");
                } else if (bDateField) {
                    var sDateFieldName = sFieldName;
                    var fnDateChange = (function (sField) {
                        return function (oEvent) {
                            var oDP = oEvent.getSource();
                            var sNewValue = oDP.getValue().replace(/-/g, "");
                            var oCtx = oDP.getBindingContext("detailModel");
                            if (oCtx) {
                                oCtx.getModel().setProperty(oCtx.getPath() + "/" + sField, sNewValue);
                            }
                            that._onFieldChange(oEvent);
                        };
                    }(sDateFieldName));
                    oTemplate = new DatePicker({
                        value: {
                            path: "detailModel>" + sFieldName,
                            formatter: function (vValue) {
                                if (!vValue) {
                                    return "";
                                }
                                var sValue = String(vValue);
                                if (/^\d{8}$/.test(sValue)) {
                                    return sValue.substring(0, 4) + "-" + sValue.substring(4, 6) + "-" + sValue.substring(6, 8);
                                }
                                return sValue;
                            }
                        },
                        valueFormat: "yyyy-MM-dd",
                        displayFormat: "medium",
                        placeholder: " ",
                        editable: "{= ${detailModel>/editMode} === true && ${detailModel>_isNew} === true }",
                        required: "{= ${detailModel>_isNew} === true }",
                        valueState: "{= ${detailModel>_err_" + sFieldName + "} ? 'Error' : 'None' }",
                        change: fnDateChange
                    }).addStyleClass("sapUiSizeCompact");
                } else if (bNonEditableInput) {
                    oTemplate = new Input({
                        value: "{detailModel>" + sFieldName + "}",
                        editable: false
                    }).addStyleClass("sapUiSizeCompact");
                } else if (bEditableField) {
                    var bLockWhenConsumed = sConsumedQtyField && sLabelUpper === "ROC";
                    var oInputCfg = {
                        value: "{detailModel>" + sFieldName + "}",
                        editable: bLockWhenConsumed ? {
                            parts: [
                                { path: "detailModel>/editMode" },
                                { path: "detailModel>" + sConsumedQtyField }
                            ],
                            formatter: function (bEditMode, vConsumedQty) {
                                var fConsumedQty = parseFloat(String(vConsumedQty || "0").replace(/,/g, ""));
                                return bEditMode === true && !(fConsumedQty > 0);
                            }
                        } : "{detailModel>/editMode}",
                        change: that._onFieldChange.bind(that),
                        liveChange: that._onFieldChange.bind(that)
                    };
                    if (!bIsComment) {
                        oInputCfg.valueState = "{= ${detailModel>_err_" + sFieldName + "} ? 'Error' : 'None' }";
                        oInputCfg.required = true;
                    }
                    oTemplate = new Input(oInputCfg).addStyleClass("sapUiSizeCompact");
                } else {
                    var sVHField = sFieldName;
                    var sVHLabel = oCol.label || sFieldName;
                    oTemplate = new Input({
                        value: "{detailModel>" + sFieldName + "}",
                        editable: "{= ${detailModel>/editMode} === true && ${detailModel>_isNew} === true }",
                        required: "{= ${detailModel>_isNew} === true }",
                        valueState: "{= ${detailModel>_err_" + sFieldName + "} ? 'Error' : 'None' }",
                        showValueHelp: "{= ${detailModel>/editMode} === true && ${detailModel>_isNew} === true }",
                        valueHelpRequest: function (oEvent) {
                            that._onValueHelpRequest(oEvent, sVHField, sVHLabel);
                        },
                        change: that._onFieldChange.bind(that),
                        liveChange: that._onFieldChange.bind(that)
                    }).addStyleClass("sapUiSizeCompact");
                }

                oTable.addColumn(new UIColumn({
                    width: "150px",
                    label: new Label({ text: oCol.label + (bIsRequired ? " *" : ""), wrapping: false }),
                    template: oTemplate,
                    resizable: true,
                    autoResizable: true
                }));
            });

            oTable.setFixedColumnCount(iFixedCount);
            oTable.setRowSettingsTemplate(new RowSettings({
                highlight: "{= ${detailModel>_overlapError} ? 'Error' : 'None' }"
            }));
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
            var oModel = this.getView().getModel("detailModel");
            if (oModel.getProperty("/hasChanges")) {
                var that = this;
                this._showUnsavedPopup(function () {
                    var oDetailModel = that.getView().getModel("detailModel");
                    oDetailModel.setProperty("/hasChanges", false);
                    oDetailModel.setProperty("/messageVisible", false);
                    oDetailModel.setProperty("/messageText", "");
                    oDetailModel.setProperty("/messageType", "None");
                    that.getOwnerComponent().getRouter().navTo("RouteListReport", {}, true);
                });
                return;
            }
            this._doNavBack();
        },

        _doNavBack: function () {
            var oModel = this.getView().getModel("detailModel");
            oModel.setProperty("/messageVisible", false);
            oModel.setProperty("/messageText", "");
            oModel.setProperty("/messageType", "None");
            var oHistory = History.getInstance();
            var sPreviousHash = oHistory.getPreviousHash();
            if (sPreviousHash !== undefined) {
                window.history.go(-1);
            } else {
                this.getOwnerComponent().getRouter().navTo("RouteListReport", {}, true);
            }
        },

        _showUnsavedPopup: function (fnOnAbandon, fnOnContinue) {
            if (this._bUnsavedPopupOpen) { return; }
            this._bUnsavedPopupOpen = true;
            var that = this;
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var oDialog = new sap.m.Dialog({
                title: oBundle.getText("warningTitle"),
                type: "Message",
                state: "Warning",
                content: [new sap.m.Text({ text: oBundle.getText("msgUnsavedChanges") })],
                beginButton: new sap.m.Button({
                    text: oBundle.getText("continuarEdicion"),
                    press: function () {
                        oDialog.close();
                        if (fnOnContinue) { fnOnContinue(); }
                    }
                }),
                endButton: new sap.m.Button({
                    text: oBundle.getText("abandonar"),
                    type: "Reject",
                    press: function () { oDialog.close(); fnOnAbandon(); }
                }),
                afterClose: function () {
                    that._bUnsavedPopupOpen = false;
                    oDialog.destroy();
                }
            });
            oDialog.open();
        },

        onAddNewRow: function () {
            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows") || [];
            var aColumns = oModel.getProperty("/columns") || [];
            var sProductAllocationObject = oModel.getProperty("/productAllocationObject") || "";

            var oNewRow = {};
            aColumns.forEach(function (oCol) {
                oNewRow[oCol.name] = "";
                oNewRow[oCol.name + "_old"] = "";
            });

            oNewRow["PRODUCTALLOCATIONOBJECT"] = "";
            oNewRow["PRODUCTALLOCATIONOBJECT_old"] = "";
            oNewRow["VAR_CHAR"] = sProductAllocationObject;
            oNewRow["VAR_CHAR_old"] = sProductAllocationObject;
            var sKeyChar = oModel.getProperty("/l_key_char") || "";
            oNewRow["KEY_CHAR"] = sKeyChar;
            oNewRow["KEY_CHAR_old"] = sKeyChar;
            var bEn = this._getSapLang() === "en";
            var sDefStatus     = "Active";
            var sDefConstraint = bEn ? "As in Sequence Constraint" : "Como en restricci\u00f3n de secuencia";
            oNewRow["PRODALLOCATIONACTIVATIONSTATUS"] = sDefStatus;
            oNewRow["PRODALLOCATIONACTIVATIONSTATUS_old"] = sDefStatus;
            oNewRow["PRODALLOCCHARCCONSTRAINTSTATUS"] = sDefConstraint;
            oNewRow["PRODALLOCCHARCCONSTRAINTSTATUS_old"] = sDefConstraint;
            oNewRow["ZZRFCUT"] = "08";
            oNewRow["ZZRFCUT_old"] = "08";
            oNewRow["_isNew"] = true;

            aRows.push(oNewRow);
            oModel.setProperty("/rows", aRows);

            var iNewRowCount = Math.min(aRows.length, 15);
            oModel.setProperty("/rowCount", iNewRowCount);

            oModel.setProperty("/hasChanges", true);

            this._oOriginalData.push(JSON.parse(JSON.stringify(oNewRow)));
        },

        onDeleteRows: function () {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();

            if (aSelectedIndices.length === 0) {
                MessageToast.show(this.getView().getModel("i18n").getResourceBundle().getText("msgSelectItem"));
                return;
            }

            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows") || [];

            aSelectedIndices.sort(function (a, b) { return b - a; });

            var aColumns = oModel.getProperty("/columns") || [];

            for (var i = 0; i < aSelectedIndices.length; i++) {
                var iIndex = aSelectedIndices[i];
                var oRowCopy = JSON.parse(JSON.stringify(aRows[iIndex]));
                aColumns.forEach(function (oCol) {
                    var sField = oCol.name;
                    var sOldKey = sField + "_old";
                    if (!oRowCopy[sOldKey]) {
                        oRowCopy[sOldKey] = oRowCopy[sField] || "";
                    }
                });
                if (!aRows[iIndex]._isNew) {
                    this._aDeletedRows.push({ rowIndex: iIndex, rowData: oRowCopy });
                }
                aRows.splice(iIndex, 1);
            }

            oModel.setProperty("/rows", aRows);

            var iNewRowCount = Math.min(aRows.length, 15);
            oModel.setProperty("/rowCount", iNewRowCount);

            oTable.clearSelection();

            this._hasDeletedRows = true;
            oModel.setProperty("/hasChanges", true);
        },

        onCopyRow: function () {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();

            if (aSelectedIndices.length === 0) {
                MessageToast.show(this.getView().getModel("i18n").getResourceBundle().getText("msgSelectOneItem"));
                return;
            }

            if (aSelectedIndices.length > 1) {
                MessageToast.show(this.getView().getModel("i18n").getResourceBundle().getText("msgSelectOneItem"));
                return;
            }

            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows") || [];
            var iSelectedIndex = aSelectedIndices[0];
            var oSelectedRow = aRows[iSelectedIndex];

            var oNewRow = JSON.parse(JSON.stringify(oSelectedRow));

            var aColumns = oModel.getProperty("/columns") || [];
            aColumns.forEach(function (oCol) {
                var sLabel = (oCol.label || "").toLowerCase().trim();
                oNewRow[oCol.name + "_old"] = "";
                var sColUpper = (oCol.name || "").toUpperCase();
                if (sLabel === "avbl qty" || sLabel === "cnsmd qty" ||
                    sColUpper === "PRODUCTALLOCATIONOBJECT" ||
                    sColUpper === "PRODUCTALLOCATIONOBJECTUUID") {
                    oNewRow[oCol.name] = "";
                }
            });

            oNewRow["_isNew"] = true;

            aRows.push(oNewRow);
            oModel.setProperty("/rows", aRows);

            var iNewRowCount = Math.min(aRows.length, 15);
            oModel.setProperty("/rowCount", iNewRowCount);

            oTable.clearSelection();

            oModel.setProperty("/hasChanges", true);

            this._oOriginalData.push(JSON.parse(JSON.stringify(oNewRow)));
        },

        onOpenChangeStatusMenu: function (oEvent) {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();
            var oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (aSelectedIndices.length === 0) {
                MessageToast.show(oBundle.getText("msgSelectRows"));
                return;
            }

            var oButton = oEvent.getSource();
            if (!this._oStatusMenu) {
                var that = this;
                var bEnS = this._getSapLang() === "en";
                this._oStatusMenu = new sap.m.Menu({
                    title: oBundle.getText("changeStatusButton"),
                    items: [
                        new sap.m.MenuItem({ text: bEnS ? "Active"   : "Activos",  press: function () { that.onChangeStatus(bEnS ? "Active"   : "Activos");  } }),
                        new sap.m.MenuItem({ text: bEnS ? "Inactive" : "Inactivos", press: function () { that.onChangeStatus(bEnS ? "Inactive" : "Inactivos"); } })
                    ]
                });
                this.getView().addDependent(this._oStatusMenu);
            }
            this._oStatusMenu.openBy(oButton);
        },

        onChangeStatus: function (sEnglishValue) {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();
            if (aSelectedIndices.length === 0) { return; }

            var oModel = this.getView().getModel("detailModel");
            aSelectedIndices.forEach(function (iIdx) {
                oModel.setProperty("/rows/" + iIdx + "/PRODALLOCATIONACTIVATIONSTATUS", sEnglishValue);
            });
            oModel.setProperty("/hasChanges", true);
        },

        onOpenChangeConstraintStatusMenu: function (oEvent) {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();
            var oBundle = this.getView().getModel("i18n").getResourceBundle();

            if (aSelectedIndices.length === 0) {
                MessageToast.show(oBundle.getText("msgSelectRows"));
                return;
            }

            var oButton = oEvent.getSource();
            if (!this._oConstraintStatusMenu) {
                var that = this;
                var bEnC = this._getSapLang() === "en";
                this._oConstraintStatusMenu = new sap.m.Menu({
                    title: oBundle.getText("changeConstraintStatusButton"),
                    items: [
                        new sap.m.MenuItem({ text: bEnC ? "Unrestricted Availability" : "Disponibilidad no restringida",    press: that.onChangeConstraintStatus.bind(that) }),
                        new sap.m.MenuItem({ text: bEnC ? "Restricted Availablity"    : "Disponibilidad restringida",      press: that.onChangeConstraintStatus.bind(that) }),
                        new sap.m.MenuItem({ text: bEnC ? "No Availability"           : "Sin disponibilidad",              press: that.onChangeConstraintStatus.bind(that) }),
                        new sap.m.MenuItem({ text: bEnC ? "Not Relevant"              : "No relevante",                    press: that.onChangeConstraintStatus.bind(that) }),
                        new sap.m.MenuItem({ text: bEnC ? "As in Sequence Constraint" : "Como en restricci\u00f3n de secuencia", press: that.onChangeConstraintStatus.bind(that) })
                    ]
                });
                this.getView().addDependent(this._oConstraintStatusMenu);
            }
            this._oConstraintStatusMenu.openBy(oButton);
        },

        onChangeConstraintStatus: function (oEvent) {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();
            if (aSelectedIndices.length === 0) { return; }

            var sNewStatus = oEvent.getSource().getText();
            var oModel = this.getView().getModel("detailModel");
            aSelectedIndices.forEach(function (iIdx) {
                oModel.setProperty("/rows/" + iIdx + "/PRODALLOCCHARCCONSTRAINTSTATUS", sNewStatus);
            });
            oModel.setProperty("/hasChanges", true);
        },

        onDownload: function () {
            var that = this;
            if (!this._oDownloadDialog) {
                var oCbDesc = new sap.m.CheckBox({ text: "With Descriptions", selected: false });
                var oRbgFormat = new sap.m.RadioButtonGroup({
                    columns: 1,
                    selectedIndex: 0,
                    buttons: [
                        new sap.m.RadioButton({ text: "As Spreadsheet (.xlsx)" }),
                        new sap.m.RadioButton({ text: "As Comma Separated Values File (.csv)" })
                    ]
                });
                this._oDownloadDialog = new sap.m.Dialog({
                    title: "Download Data",
                    contentWidth: "22rem",
                    content: [
                        new sap.m.VBox({
                            items: [oCbDesc, oRbgFormat]
                        }).addStyleClass("sapUiSmallMargin")
                    ],
                    beginButton: new sap.m.Button({
                        text: "OK",
                        type: "Emphasized",
                        press: function () {
                            var bWithDesc = oCbDesc.getSelected();
                            var sFormat = oRbgFormat.getSelectedIndex() === 0 ? "xlsx" : "csv";
                            that._oDownloadDialog.close();
                            that._executeDownload(sFormat, bWithDesc);
                        }
                    }),
                    endButton: new sap.m.Button({
                        text: "Cancel",
                        press: function () { that._oDownloadDialog.close(); }
                    })
                });
                this.getView().addDependent(this._oDownloadDialog);
            }
            this._oDownloadDialog.open();
        },

        _executeDownload: function (sFormat, bWithDesc) {
            var that = this;
            if (sFormat === "csv") {
                this._generateCsv(bWithDesc);
                return;
            }
            this._loadSheetJS().then(function (XLSX) {
                that._generateXlsx(XLSX, bWithDesc);
            }).catch(function (oError) {
                jQuery.sap.log.error("Could not load Excel library.", oError && (oError.message || oError.toString ? oError.toString() : ""));
                MessageToast.show("Excel library could not be loaded. Downloading CSV instead.");
                that._generateCsv(bWithDesc);
            });
        },

        _generateCsv: function (bWithDesc) {
            var oModel = this.getView().getModel("detailModel");
            var aColumns = this._getDownloadColumns(bWithDesc);
            var aRows = oModel.getProperty("/rows") || [];

            var aAoA = [
                ["Activation Status",  "Activation Status - Description"],
                ["01",                 "Inactive"],
                ["02",                 "Active"],
                [],
                ["Constraint Status",  "Constraint Status - Description"],
                ["01",                 "Unrestricted Availablity"],
                ["02",                 "Restricted Availablity"],
                ["03",                 "No Availability"],
                ["04",                 "Not Relevant"],
                ["05",                 "As in Sequence Constraint"],
                [],
                ["Delete CVC",         "Delete CVC - Description"],
                ["",                   "No Deletion"],
                ["1",                  "Standard Deletion"],
                ["2",                  "Deletion with Consumptions"],
                []
            ];

            var aHeader = aColumns.map(function (oCol) { return oCol.label || oCol.name; });
            aAoA.push(aHeader);

            aRows.forEach(function (oRow) {
                var aRow = aColumns.map(function (oCol) {
                    var v = oRow[oCol.name];
                    return (v === undefined || v === null) ? "" : v;
                });
                aAoA.push(aRow);
            });

            var sSep = ";";
            var fnEscape = function (v) {
                var s = String(v == null ? "" : v);
                if (s.indexOf(sSep) !== -1 || s.indexOf("\"") !== -1 || s.indexOf("\n") !== -1 || s.indexOf("\r") !== -1) {
                    s = "\"" + s.replace(/"/g, "\"\"") + "\"";
                }
                return s;
            };
            var sCsv = aAoA.map(function (aRow) {
                return aRow.map(fnEscape).join(sSep);
            }).join("\r\n");

            // UTF-8 BOM so Excel detects encoding correctly
            var sContent = "\uFEFF" + sCsv;
            var oBlob = new Blob([sContent], { type: "text/csv;charset=utf-8;" });
            var sFileName = this._buildDownloadFileName("csv");
            var sUrl = URL.createObjectURL(oBlob);
            var oLink = document.createElement("a");
            oLink.href = sUrl;
            oLink.download = sFileName;
            document.body.appendChild(oLink);
            oLink.click();
            document.body.removeChild(oLink);
            URL.revokeObjectURL(sUrl);
        },

        _isProdDescColumn: function (oCol) {
            var sName = (oCol.name || "").toUpperCase();
            return sName.indexOf("PROD") === 0 && sName.lastIndexOf("DESC") === sName.length - 4;
        },

        _getDownloadColumns: function (bWithDesc) {
            var oModel = this.getView().getModel("detailModel");
            var aExcludedLabels = ["avbl qty", "cnsmd qty"];

            return (oModel.getProperty("/columns") || []).reduce(function (aAcc, oCol) {
                var sName = (oCol.name  || "").toUpperCase();
                var sLbl  = (oCol.label || "").toLowerCase().trim();
                var bProdDesc = this._isProdDescColumn(oCol);

                if (sName === "PRODUCTALLOCATIONOBJECTUUID" || sName === "VAR_CHAR" || sName === "KEY_CHAR") { return aAcc; }
                if (aExcludedLabels.indexOf(sLbl) !== -1) { return aAcc; }
                if (bProdDesc && !bWithDesc) { return aAcc; }

                var sLabel = oCol.label || oCol.name;
                if (bProdDesc) {
                    var oPreviousColumn = aAcc[aAcc.length - 1];
                    var sPreviousLabel = oPreviousColumn ? oPreviousColumn.label : sLabel;
                    sLabel = sPreviousLabel + " - Description";
                }

                aAcc.push({
                    name: oCol.name,
                    label: sLabel
                });

                return aAcc;
            }.bind(this), []);
        },

        _loadSheetJS: function () {
            return new Promise(function (resolve, reject) {
                if (window.XLSX) { resolve(window.XLSX); return; }
                var oScript = document.createElement("script");
                oScript.src = "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js";
                oScript.onload = function () {
                    if (window.XLSX) {
                        resolve(window.XLSX);
                    } else {
                        reject(new Error("SheetJS script loaded but XLSX object is not available."));
                    }
                };
                oScript.onerror = function () { reject(new Error("SheetJS script could not be loaded from CDN.")); };
                document.head.appendChild(oScript);
            });
        },

        _buildDownloadFileName: function (sExt) {
            var oModel = this.getView().getModel("detailModel");
            var sObj = oModel.getProperty("/productAllocationObject") || "data";
            var d = new Date();
            var pad = function (n) { return ("0" + n).slice(-2); };
            var sTs = d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
                      pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
            return sObj + "_" + sTs + "." + sExt;
        },

        _generateXlsx: function (XLSX, bWithDesc) {
            var oModel = this.getView().getModel("detailModel");
            var aColumns = this._getDownloadColumns(bWithDesc);
            var aRows = oModel.getProperty("/rows") || [];

            // First 16 reference lines (lines 1-15 with content, line 16 blank)
            var aAoA = [
                ["Activation Status",  "Activation Status - Description"],
                ["01",                 "Inactive"],
                ["02",                 "Active"],
                [],
                ["Constraint Status",  "Constraint Status - Description"],
                ["01",                 "Unrestricted Availablity"],
                ["02",                 "Restricted Availablity"],
                ["03",                 "No Availability"],
                ["04",                 "Not Relevant"],
                ["05",                 "As in Sequence Constraint"],
                [],
                ["Delete CVC",         "Delete CVC - Description"],
                ["",                   "No Deletion"],
                ["1",                  "Standard Deletion"],
                ["2",                  "Deletion with Consumptions"],
                []
            ];

            // Column headers from table
            var aHeader = aColumns.map(function (oCol) { return oCol.label || oCol.name; });
            aAoA.push(aHeader);

            // Data rows
            aRows.forEach(function (oRow) {
                var aRow = aColumns.map(function (oCol) {
                    var v = oRow[oCol.name];
                    return (v === undefined || v === null) ? "" : v;
                });
                aAoA.push(aRow);
            });

            var oWs = XLSX.utils.aoa_to_sheet(aAoA);

            // Apply font Aptos Narrow size 11 to every cell
            var oFontStyle = { font: { name: "Aptos Narrow", sz: 11 } };
            var oRange = XLSX.utils.decode_range(oWs["!ref"]);

            // Column widths: double the Excel default (8.43 ch -> ~17 ch)
            var aCols = [];
            for (var c = oRange.s.c; c <= oRange.e.c; c++) {
                aCols.push({ wch: 17 });
            }
            oWs["!cols"] = aCols;
            for (var R = oRange.s.r; R <= oRange.e.r; R++) {
                for (var C = oRange.s.c; C <= oRange.e.c; C++) {
                    var sAddr = XLSX.utils.encode_cell({ r: R, c: C });
                    var oCell = oWs[sAddr];
                    if (!oCell) {
                        oWs[sAddr] = { t: "s", v: "", s: oFontStyle };
                    } else {
                        oCell.s = Object.assign({}, oCell.s || {}, oFontStyle);
                    }
                }
            }

            var oWb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(oWb, oWs, "Data");
            var sFileName = this._buildDownloadFileName("xlsx");
            XLSX.writeFile(oWb, sFileName);
        },

        onUpload: function () {
            var that = this;
            if (!this._oUploadInput) {
                var oInput = document.createElement("input");
                oInput.type = "file";
                oInput.accept = ".xlsx,.xls";
                oInput.style.display = "none";
                oInput.addEventListener("change", function (e) {
                    var oFile = e.target.files && e.target.files[0];
                    if (oFile) { that._onUploadFileSelected(oFile); }
                    oInput.value = "";
                });
                document.body.appendChild(oInput);
                this._oUploadInput = oInput;
            }
            this._oUploadInput.click();
        },

        _onUploadFileSelected: function (oFile) {
            var that = this;
            this._loadSheetJS().then(function (XLSX) {
                var oReader = new FileReader();
                oReader.onload = function (e) {
                    try {
                        var oData = new Uint8Array(e.target.result);
                        var oWb = XLSX.read(oData, { type: "array", cellDates: true });
                        var sFirstSheet = oWb.SheetNames[0];
                        var oWs = oWb.Sheets[sFirstSheet];
                        var aAoA = XLSX.utils.sheet_to_json(oWs, { header: 1, defval: "", blankrows: true });
                        that._processUploadedRows(aAoA);
                    } catch (err) {
                        MessageBox.error("Error reading Excel file: " + err.message);
                    }
                };
                oReader.onerror = function () {
                    MessageBox.error("Could not read the selected file.");
                };
                oReader.readAsArrayBuffer(oFile);
            }).catch(function () {
                MessageBox.error("Could not load Excel library.");
            });
        },

        _processUploadedRows: function (aAoA) {
            // Skip first 16 reference lines, header is row 17 (index 16), data starts row 18 (index 17)
            if (!aAoA || aAoA.length < 17) {
                MessageBox.error("Invalid file structure. Expected the same layout as the downloaded template.");
                return;
            }
            var aHeader = aAoA[16] || [];
            var aDataRows = aAoA.slice(17).filter(function (aRow) {
                return aRow && aRow.some(function (v) { return v !== "" && v !== null && v !== undefined; });
            });

            if (aDataRows.length === 0) {
                MessageToast.show("No data rows found in file.");
                return;
            }

            var oModel = this.getView().getModel("detailModel");
            var aColumns = oModel.getProperty("/columns") || [];
            var sProductAllocationObject = oModel.getProperty("/productAllocationObject") || "";

            // Map header label -> column index in header array
            var oLabelToHeaderIdx = {};
            aHeader.forEach(function (sLbl, i) {
                var sHeaderLabel = String(sLbl == null ? "" : sLbl).toLowerCase().trim();
                if (sHeaderLabel && sHeaderLabel.lastIndexOf(" - description") !== sHeaderLabel.length - 14) {
                    oLabelToHeaderIdx[sHeaderLabel] = i;
                }
            });

            // Validate header labels match what is expected (same exclusions as download)
            var aExpectedLabels = aColumns.filter(function (oCol) {
                var sName = (oCol.name  || "").toUpperCase();
                var sLbl  = (oCol.label || "").toLowerCase().trim();
                if (sName === "PRODUCTALLOCATIONOBJECTUUID" || sName === "VAR_CHAR" || sName === "KEY_CHAR") { return false; }
                if (sLbl === "avbl qty" || sLbl === "cnsmd qty") { return false; }
                if (this._isProdDescColumn(oCol)) { return false; }
                return !!sLbl;
            }.bind(this)).map(function (oCol) { return (oCol.label || "").toLowerCase().trim(); });

            var aFileLabels = Object.keys(oLabelToHeaderIdx);
            var bLabelsOk = aExpectedLabels.length === aFileLabels.length &&
                aExpectedLabels.every(function (s) { return oLabelToHeaderIdx[s] !== undefined; });
            if (!bLabelsOk) {
                MessageBox.error("ERROR!: labels in file", { actions: ["OK"] });
                return;
            }

            // Validate Allocation Object in file matches screen
            var oAllocCol = aColumns.filter(function (oCol) {
                return (oCol.name || "").toUpperCase() === "PRODUCTALLOCATIONOBJECT";
            })[0];
            var iAllocIdx = oAllocCol ? oLabelToHeaderIdx[(oAllocCol.label || "").toLowerCase().trim()] : undefined;
            var bAllocOk = true;
            if (iAllocIdx === undefined) {
                bAllocOk = false;
            } else {
                for (var i = 0; i < aDataRows.length; i++) {
                    var sFileAlloc = String(aDataRows[i][iAllocIdx] == null ? "" : aDataRows[i][iAllocIdx]).trim();
                    if (sFileAlloc !== sProductAllocationObject) { bAllocOk = false; break; }
                }
            }
            if (!bAllocOk) {
                MessageBox.error("ERROR! Allocation Object in file does not correspond", { actions: ["OK"] });
                return;
            }

            var bEn = this._getSapLang() === "en";
            var sDefStatus     = "Active";
            var sDefConstraint = bEn ? "As in Sequence Constraint" : "Como en restricci\u00f3n de secuencia";

            var aExistingRows = oModel.getProperty("/rows") || [];

            // Identify date fields up front
            var oDateFieldSet = {};
            var sQuotaQtyField = null;
            var sConsumedQtyField = null;
            var sRocField = null;
            var sCommentField = null;
            aColumns.forEach(function (oCol) {
                var u = (oCol.name || "").toUpperCase();
                var sLabelUpper = (oCol.label || "").toUpperCase().trim();
                if (u === "PRODALLOCPERDSTARTUTCDATE" || u === "PRODALLOCPERIODENDUTCDATE") {
                    oDateFieldSet[oCol.name] = true;
                }
                if (sLabelUpper === "QUOTA QTY") { sQuotaQtyField = oCol.name; }
                if (sLabelUpper === "CNSMD QTY") { sConsumedQtyField = oCol.name; }
                if (sLabelUpper === "ROC") { sRocField = oCol.name; }
                if (sLabelUpper === "COMMENT" || u === "PRODALLOCCHARCVALUECOMBNCMNT") { sCommentField = oCol.name; }
            });

            // Robust date parser: accepts Date, number (Excel serial), or string in common formats.
            // Returns "yyyymmdd" if valid, "" if input is empty, null if invalid.
            var fnParseDateCell = function (v) {
                if (v === undefined || v === null || v === "") { return ""; }
                var d = null;
                if (v instanceof Date) {
                    d = v;
                } else if (typeof v === "number") {
                    var n = Math.round(v);
                    if (n >= 18000101 && n <= 22001231) {
                        // yyyymmdd numeric
                        var ny = Math.floor(n / 10000);
                        var nm = Math.floor((n % 10000) / 100);
                        var nd = n % 100;
                        d = new Date(Date.UTC(ny, nm - 1, nd));
                        if (d.getUTCFullYear() !== ny || (d.getUTCMonth() + 1) !== nm || d.getUTCDate() !== nd) { return null; }
                    } else if (n >= 1 && n <= 2958465) {
                        // Excel serial date (days since 1899-12-30)
                        var iEpoch = Date.UTC(1899, 11, 30);
                        d = new Date(iEpoch + n * 86400000);
                    } else {
                        return null;
                    }
                } else {
                    var s = String(v).trim();
                    if (!s) { return ""; }
                    var m;
                    if ((m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) {
                        d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
                        if (d.getUTCFullYear() !== +m[1] || (d.getUTCMonth() + 1) !== +m[2] || d.getUTCDate() !== +m[3]) { return null; }
                    } else if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) {
                        d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
                        if (d.getUTCFullYear() !== +m[1] || (d.getUTCMonth() + 1) !== +m[2] || d.getUTCDate() !== +m[3]) { return null; }
                    } else if ((m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/))) {
                        d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
                        if (d.getUTCFullYear() !== +m[3] || (d.getUTCMonth() + 1) !== +m[2] || d.getUTCDate() !== +m[1]) { return null; }
                    } else {
                        var t = Date.parse(s);
                        if (isNaN(t)) { return null; }
                        d = new Date(t);
                    }
                }
                if (!d || isNaN(d.getTime())) { return null; }
                var yyyy = String(d.getUTCFullYear());
                var mm = ("0" + (d.getUTCMonth() + 1)).slice(-2);
                var dd = ("0" + d.getUTCDate()).slice(-2);
                return yyyy + mm + dd;
            };

            // Build candidate new rows (not yet pushed)
            var bInvalidDate = false;
            var oController = this;
            var aCandidateRows = aDataRows.map(function (aXlsxRow, iDataIdx) {
                var oNewRow = {};
                oNewRow._excelLine = iDataIdx + 18;
                aColumns.forEach(function (oCol) {
                    oNewRow[oCol.name] = "";
                    oNewRow[oCol.name + "_old"] = "";
                });
                oNewRow["PRODUCTALLOCATIONOBJECT"] = sProductAllocationObject;
                oNewRow["PRODUCTALLOCATIONOBJECT_old"] = sProductAllocationObject;
                oNewRow["VAR_CHAR"] = sProductAllocationObject;
                oNewRow["VAR_CHAR_old"] = sProductAllocationObject;
                oNewRow["KEY_CHAR"] = oController.getView().getModel("detailModel").getProperty("/l_key_char") || "";
                oNewRow["KEY_CHAR_old"] = oNewRow["KEY_CHAR"];
                oNewRow["PRODALLOCATIONACTIVATIONSTATUS"] = sDefStatus;
                oNewRow["PRODALLOCATIONACTIVATIONSTATUS_old"] = sDefStatus;
                oNewRow["PRODALLOCCHARCCONSTRAINTSTATUS"] = sDefConstraint;
                oNewRow["PRODALLOCCHARCCONSTRAINTSTATUS_old"] = sDefConstraint;
                oNewRow["ZZRFCUT"] = "08";
                oNewRow["ZZRFCUT_old"] = "08";
                oNewRow["_isNew"] = true;

                aColumns.forEach(function (oCol) {
                    if (oController._isProdDescColumn(oCol)) { return; }
                    var sLbl = (oCol.label || "").toLowerCase().trim();
                    if (!sLbl) { return; }
                    var iIdx = oLabelToHeaderIdx[sLbl];
                    if (iIdx === undefined) { return; }
                    var v = aXlsxRow[iIdx];
                    if (v === undefined || v === null) { return; }
                    var sVal;
                    if (oDateFieldSet[oCol.name]) {
                        var sParsed = fnParseDateCell(v);
                        if (sParsed === null) { bInvalidDate = true; sVal = ""; }
                        else { sVal = sParsed; }
                    } else {
                        sVal = String(v).trim();
                    }
                    oNewRow[oCol.name] = sVal;
                    oNewRow[oCol.name + "_old"] = sVal;
                });
                return oNewRow;
            });

            if (bInvalidDate) {
                MessageBox.error("ERROR! Dates in file!", { actions: ["OK"] });
                return;
            }

            // Per-row date range validation: end date must be after start date
            var sStartField = null, sEndField = null;
            aColumns.forEach(function (oCol) {
                var u = oCol.name.toUpperCase();
                if (u === "PRODALLOCPERDSTARTUTCDATE") { sStartField = oCol.name; }
                if (u === "PRODALLOCPERIODENDUTCDATE")  { sEndField   = oCol.name; }
            });
            var fnNormDate = function (s) {
                if (!s) { return ""; }
                var str = String(s).trim();
                if (/^\d{8}$/.test(str)) {
                    return str.substring(0, 4) + "-" + str.substring(4, 6) + "-" + str.substring(6, 8);
                }
                return str;
            };
            if (sStartField && sEndField) {
                var bRangeError = aCandidateRows.some(function (oRow) {
                    var s = fnNormDate(oRow[sStartField]);
                    var e = fnNormDate(oRow[sEndField]);
                    return s && e && e <= s;
                });
                if (bRangeError) {
                    MessageBox.error("ERROR! Dates in file!", { actions: ["OK"] });
                    return;
                }
            }

            // Compute key fields (same rule as _hasDateOverlap / Save validation)
            var aKeyFields = this._getOverlapKeyFields(aColumns);

            // Build match key = keyFields + normalized start + normalized end
            var fnMatchKey = function (oRow) {
                var aParts = aKeyFields.map(function (f) {
                    var v = oRow[f];
                    return String(v == null ? "" : v).trim();
                });
                aParts.push(sStartField ? fnNormDate(oRow[sStartField]) : "");
                aParts.push(sEndField   ? fnNormDate(oRow[sEndField])   : "");
                return aParts.join("|");
            };

            // Try to merge each candidate into an existing row with identical match key.
            // If matched: copy non-key (and non-date) field values from candidate into existing,
            // and remove the candidate from the new-rows list.
            var aWorkingRows = JSON.parse(JSON.stringify(aExistingRows));
            var oExistingByKey = {};
            aWorkingRows.forEach(function (oRow, iIdx) {
                var k = fnMatchKey(oRow);
                if (!oExistingByKey[k]) { oExistingByKey[k] = iIdx; }
            });

            var aRemainingCandidates = [];
            var aUpdatedDuplicateLogs = [];
            var iUpdatedCount = 0;
            aCandidateRows.forEach(function (oCand) {
                var k = fnMatchKey(oCand);
                var iIdx = oExistingByKey[k];
                if (iIdx !== undefined) {
                    var oTarget = aWorkingRows[iIdx];
                    var aUpdatedFields = [];
                    if (sQuotaQtyField && oCand[sQuotaQtyField] !== undefined) {
                        oTarget[sQuotaQtyField] = oCand[sQuotaQtyField];
                        aUpdatedFields.push("Quota Qty");
                    }
                    if (sRocField && oCand[sRocField] !== undefined) {
                        var fTargetConsumedQty = sConsumedQtyField ?
                            parseFloat(String(oTarget[sConsumedQtyField] || "0").replace(/,/g, "")) : 0;
                        var bRocEditable = !(fTargetConsumedQty > 0);
                        if (bRocEditable) {
                            oTarget[sRocField] = oCand[sRocField];
                            aUpdatedFields.push("RoC");
                        }
                    }
                    if (sCommentField && oCand[sCommentField] !== undefined) {
                        oTarget[sCommentField] = oCand[sCommentField];
                        aUpdatedFields.push("Comment");
                    }
                    aUpdatedDuplicateLogs.push({
                        excelLine: oCand._excelLine,
                        tableRow: iIdx + 1,
                        updatedFields: aUpdatedFields,
                        matchKey: k
                    });
                    iUpdatedCount++;
                } else {
                    aRemainingCandidates.push(oCand);
                }
            });
            console.log("[UploadExcel] Duplicate match summary:", {
                excelRows: aCandidateRows.length,
                existingRows: aWorkingRows.length,
                updatedDuplicates: aUpdatedDuplicateLogs.length,
                newRows: aRemainingCandidates.length,
                keyFields: aKeyFields,
                startField: sStartField || "",
                endField: sEndField || ""
            });
            if (aUpdatedDuplicateLogs.length > 0) {
                console.log("[UploadExcel] Duplicate lines updated:", aUpdatedDuplicateLogs);
            } else {
                console.log("[UploadExcel] No duplicate lines matched. Sample keys:", {
                    excel: aCandidateRows.slice(0, 5).map(function (oRow) {
                        return { excelLine: oRow._excelLine, matchKey: fnMatchKey(oRow) };
                    }),
                    existing: aWorkingRows.slice(0, 5).map(function (oRow, iIdx) {
                        return { tableRow: iIdx + 1, matchKey: fnMatchKey(oRow) };
                    })
                });
            }

            // Date overlap validation across existing + remaining new candidates
            var aAllRows = aWorkingRows.concat(aRemainingCandidates);
            if (this._hasDateOverlap(aAllRows, aColumns)) {
                MessageBox.error("ERROR! Dates in file!", { actions: ["OK"] });
                return;
            }

            if (sQuotaQtyField && sConsumedQtyField) {
                var bQuotaConsumedError = aAllRows.some(function (oRow) {
                    var fQuotaQty = parseFloat(String(oRow[sQuotaQtyField] || "0").replace(/,/g, ""));
                    var fConsumedQty = parseFloat(String(oRow[sConsumedQtyField] || "0").replace(/,/g, ""));
                    return !isNaN(fQuotaQty) && !isNaN(fConsumedQty) && fQuotaQty < fConsumedQty;
                });
                if (bQuotaConsumedError) {
                    MessageBox.error("Quota Qty must be greater than or equal to Cnsmd QTy.");
                    return;
                }
            }

            // Validation passed: commit rows
            var aRows = aWorkingRows;
            aRemainingCandidates.forEach(function (oNewRow) {
                aRows.push(oNewRow);
                this._oOriginalData.push(JSON.parse(JSON.stringify(oNewRow)));
            }, this);

            oModel.setProperty("/rows", aRows);
            var iNewRowCount = Math.min(aRows.length, 15);
            oModel.setProperty("/rowCount", iNewRowCount);
            oModel.setProperty("/hasChanges", true);

            var sMsg = aRemainingCandidates.length + " new row(s) loaded and " +
                iUpdatedCount + " duplicate row(s) updated from file.";
            MessageToast.show(sMsg);
        },

        _getOverlapKeyFields: function (aColumns) {
            var that = this;
            var iCsIdx = -1;
            aColumns.forEach(function (oCol, iIdx) {
                if (oCol.name.toUpperCase().indexOf("STATUS") !== -1) { iCsIdx = iIdx; }
            });
            var aNonKey = ["PRODALLOCPERDSTARTUTCDATE", "PRODALLOCPERIODENDUTCDATE",
                           "PRODUCTALLOCATIONQUANTITY", "ZZRFCUT",
                           "PRODALLOCCHARCVALUECOMBNCMNT", "PRODUCTALLOCATIONOBJECTUUID"];
            return (iCsIdx >= 0 ? aColumns.slice(0, iCsIdx + 1) : aColumns).filter(function (c) {
                var u = c.name.toUpperCase();
                return aNonKey.indexOf(u) === -1 &&
                       u.indexOf("STATUS") === -1 &&
                       u.indexOf("AVBL")   === -1 &&
                       u.indexOf("CNSMD")  === -1 &&
                       !that._isProdDescColumn(c);
            }).map(function (c) { return c.name; });
        },

        _hasDateOverlap: function (aRows, aColumns) {
            var sStartField = null, sEndField = null;
            aColumns.forEach(function (oCol) {
                var u = oCol.name.toUpperCase();
                if (u === "PRODALLOCPERDSTARTUTCDATE") { sStartField = oCol.name; }
                if (u === "PRODALLOCPERIODENDUTCDATE")  { sEndField   = oCol.name; }
            });
            if (!sStartField || !sEndField) { return false; }

            var fnNormDate = function (s) {
                if (!s) { return ""; }
                var str = String(s).trim();
                if (/^\d{8}$/.test(str)) {
                    return str.substring(0, 4) + "-" + str.substring(4, 6) + "-" + str.substring(6, 8);
                }
                return str;
            };

            // Determine key fields (same rule used in Save validation)
            var that = this;
            var iCsIdx = -1;
            aColumns.forEach(function (oCol, iIdx) {
                if (oCol.name.toUpperCase().indexOf("STATUS") !== -1) { iCsIdx = iIdx; }
            });
            var aNonKey = ["PRODALLOCPERDSTARTUTCDATE", "PRODALLOCPERIODENDUTCDATE",
                           "PRODUCTALLOCATIONQUANTITY", "ZZRFCUT",
                           "PRODALLOCCHARCVALUECOMBNCMNT", "PRODUCTALLOCATIONOBJECTUUID"];
            var aKeyFields = (iCsIdx >= 0
                ? aColumns.slice(0, iCsIdx + 1)
                : aColumns
            ).filter(function (c) {
                var u = c.name.toUpperCase();
                return aNonKey.indexOf(u) === -1 &&
                       u.indexOf("STATUS") === -1 &&
                       u.indexOf("AVBL")   === -1 &&
                       u.indexOf("CNSMD")  === -1 &&
                       !that._isProdDescColumn(c);
            }).map(function (c) { return c.name; });

            var oGroups = {};
            aRows.forEach(function (oRow, iIdx) {
                var sGK = aKeyFields.map(function (f) {
                    var v = oRow[f];
                    return String(v == null ? "" : v).trim();
                }).join("|");
                if (!oGroups[sGK]) { oGroups[sGK] = []; }
                oGroups[sGK].push({ idx: iIdx, row: oRow });
            });

            var bOverlap = false;
            Object.keys(oGroups).forEach(function (sGK) {
                var aGrp = oGroups[sGK];
                if (aGrp.length < 2) { return; }
                for (var ii = 0; ii < aGrp.length; ii++) {
                    for (var jj = ii + 1; jj < aGrp.length; jj++) {
                        var asStart = fnNormDate(aGrp[ii].row[sStartField]);
                        var asEnd   = fnNormDate(aGrp[ii].row[sEndField]);
                        var bsStart = fnNormDate(aGrp[jj].row[sStartField]);
                        var bsEnd   = fnNormDate(aGrp[jj].row[sEndField]);
                        if (asStart && asEnd && bsStart && bsEnd) {
                            if (asStart <= bsEnd && bsStart <= asEnd) { bOverlap = true; }
                        }
                    }
                }
            });
            return bOverlap;
        },

        _onValueHelpRequest: function (oEvent, sFieldName, sLabel) {
            var oInput = oEvent.getSource();
            var oCtx = oInput.getBindingContext("detailModel");
            var that = this;

            // Resolve data_element from cell metadata (existing rows) or field-level fallback (new rows)
            var sDataElement = "";
            if (oCtx) {
                var sPath = oCtx.getPath();
                var iRowIdx = parseInt((sPath.match(/\/rows\/(\d+)/) || [])[1], 10);
                if (!isNaN(iRowIdx)) {
                    var oCellMeta = (this._oCellKeys || {})[iRowIdx + "_" + sFieldName];
                    if (oCellMeta && oCellMeta.data_element) {
                        sDataElement = oCellMeta.data_element;
                    }
                }
            }
            if (!sDataElement) {
                var oFieldMeta = (this._oFieldMetadata || {})[sFieldName];
                if (oFieldMeta && oFieldMeta.data_element) {
                    sDataElement = oFieldMeta.data_element;
                }
            }

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
                console.log("ValueHelp records displayed:", aPageItems.length, "total:", aAllItems.length);
            };
            var oSearchField = new SearchField({
                width: "100%",
                placeholder: "Search",
                search: function (oEv) {
                    var sQuery = oEv.getParameter("query") || oEv.getParameter("value") || "";
                    that._loadValueHelp(sQuery || "*", "", oVHModel, sDataElement, oDialog, fnApplyValueHelpPage);
                }
            });
            var oValueHelpTable = new MTable({
                width: "100%",
                mode: "None",
                fixedLayout: true,
                columns: [
                    new MColumn({
                        width: "30rem",
                        header: new Label({ text: sLabel })
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
                            if (!oRowContext || !oCtx) { return; }
                            var sClave = oRowContext.getProperty("Clave");
                            oCtx.getModel().setProperty(oCtx.getPath() + "/" + sFieldName, sClave);
                            that._onFieldChange({});
                            oDialog.close();
                        }
                    })
                }
            });

            oValueHelpTable.setModel(oVHModel);

            oDialog = new Dialog({
                title: "Search Help: " + sLabel,
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
            this._loadValueHelp("*", "", oVHModel, sDataElement, oDialog, fnApplyValueHelpPage);
        },

        _loadValueHelp: function (sSource, sSearch, oVHModel, sDataElement, oDialog, fnApplyValueHelpPage) {
            var oODataModel = this.getOwnerComponent().getModel();
            if (!oODataModel) { return; }
            var sAlloc = this.getView().getModel("detailModel").getProperty("/productAllocationObject") || "";
            var sServiceUrl = (oODataModel.sServiceUrl || "").replace(/\/$/, "");
            var aFilters = [
                new Filter("source",           FilterOperator.EQ, sSource),
                new Filter("allocationObject", FilterOperator.EQ, sAlloc),
                new Filter("data_element",     FilterOperator.EQ, sDataElement || "")
            ];
            console.log("[ValueHelp] GET " + sServiceUrl + "/ValueHelpSet?$filter=" +
                "source eq '" + sSource + "' and allocationObject eq '" + sAlloc +
                "' and data_element eq '" + (sDataElement || "") + "'");
            BusyIndicator.show(0);
            var iStartTime = Date.now();
            oODataModel.read("/ValueHelpSet", {
                filters: aFilters,
                success: function (oData) {
                    var aItems = (oData && oData.results) ? oData.results : (oData ? [oData] : []);
                    console.log("ValueHelp OData response time ms:", Date.now() - iStartTime);
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
                    jQuery.sap.log.error("ValueHelp call failed: " + (oErr && oErr.message ? oErr.message : ""));
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

        _onFieldChange: function (oEvent) {
            jQuery.sap.log.info("Detail._onFieldChange triggered");
            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows");
            if (aRows) {
                aRows.forEach(function (oRow) {
                    Object.keys(oRow).forEach(function (sKey) {
                        if (sKey.indexOf("_err_") === 0) { oRow[sKey] = false; }
                    });
                });
                oModel.setProperty("/rows", aRows);
            }
            oModel.setProperty("/messageVisible", false);
            var bHasChanges = this._detectChanges(aRows);
            jQuery.sap.log.info("Detail._onFieldChange: hasChanges=" + bHasChanges);
            oModel.setProperty("/hasChanges", bHasChanges);
        },

        onMessageClose: function () {
            this.getView().getModel("detailModel").setProperty("/messageVisible", false);
        },

        _getSapLang: function () {
            var sLang = (sap.ui.getCore().getConfiguration().getLanguage() || "").toLowerCase().substring(0, 2);
            return sLang === "en" ? "en" : "es";
        },

        _detectChanges: function (aRows) {
            if (!this._oOriginalData || !aRows) {
                return false;
            }

            if (this._hasDeletedRows) {
                return true;
            }

            for (var i = 0; i < aRows.length; i++) {
                var oRow = aRows[i];
                if (oRow._isNew) {
                    return true;
                }
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

                if (oRow._isNew || aChangedFields.length > 0) {
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

        onCancel: function () {
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var that = this;
            MessageBox.confirm(oBundle.getText("cancelConfirmMsg"), {
                actions: [oBundle.getText("confirmYes"), oBundle.getText("confirmNo")],
                emphasizedAction: oBundle.getText("confirmNo"),
                onClose: function (sAction) {
                    if (sAction === oBundle.getText("confirmYes")) {
                        that.getView().getModel("detailModel").setProperty("/hasChanges", false);
                        that._doNavBack();
                    }
                }
            });
        },

        onSave: function () {
            var oModel = this.getView().getModel("detailModel");
            var oBundle = this.getView().getModel("i18n").getResourceBundle();
            var aChangedRows = this._getChangedRows();

            if (aChangedRows.length === 0 && this._aDeletedRows.length === 0) {
                MessageToast.show(oBundle.getText("msgNoChanges"));
                return;
            }

            var aRows = oModel.getProperty("/rows") || [];
            var aColumns = oModel.getProperty("/columns") || [];
            var bDateError = false;
            var bRequiredError = false;
            var bQuotaConsumedError = false;
            var oController = this;

            var aNonRequired = [
                "PRODALLOCCHARCVALUECOMBNCMNT",
                "PRODALLOCATIONACTIVATIONSTATUS",
                "PRODALLOCCHARCCONSTRAINTSTATUS",
                "PRODUCTALLOCATIONOBJECT",
                "PRODUCTALLOCATIONOBJECTUUID",
                "VAR_CHAR",
                "KEY_CHAR"
            ];

            aRows.forEach(function (oRow) {
                aColumns.forEach(function (oCol) {
                    oRow["_err_" + oCol.name] = false;
                });
                oRow._overlapError = false;
            });

            var fnNormDate = function (s) {
                if (!s) { return ""; }
                var str = String(s).trim();
                if (/^\d{8}$/.test(str)) {
                    return str.substring(0, 4) + "-" + str.substring(4, 6) + "-" + str.substring(6, 8);
                }
                return str;
            };

            var sStartField = null, sEndField = null, sQuotaQtyField = null, sConsumedQtyField = null;
            aColumns.forEach(function (oCol) {
                var u = oCol.name.toUpperCase();
                var sLabelUpper = (oCol.label || "").toUpperCase().trim();
                if (u === "PRODALLOCPERDSTARTUTCDATE") { sStartField = oCol.name; }
                if (u === "PRODALLOCPERIODENDUTCDATE")  { sEndField   = oCol.name; }
                if (sLabelUpper === "QUOTA QTY") { sQuotaQtyField = oCol.name; }
                if (sLabelUpper === "CNSMD QTY") { sConsumedQtyField = oCol.name; }
            });

            aChangedRows.forEach(function (oChangedRow) {
                var oRowData = oChangedRow.rowData;

                if (oRowData._isNew) {
                    aColumns.forEach(function (oCol) {
                        var sFieldName = oCol.name;
                        var sFieldUpper = sFieldName.toUpperCase();
                        var sColLabelUpper = (oCol.label || "").toUpperCase().trim();
                        if (oController._isProdDescColumn(oCol)) { return; }
                        if (aNonRequired.indexOf(sFieldUpper) !== -1) { return; }
                        if (sColLabelUpper === "AVBL QTY" || sColLabelUpper === "CNSMD QTY") { return; }
                        var sValue = (oRowData[sFieldName] || "").toString().trim();
                        if (!sValue) {
                            oRowData["_err_" + sFieldName] = true;
                            bRequiredError = true;
                        }
                    });
                }

                var sStart = sStartField ? fnNormDate(oRowData[sStartField]) : "";
                var sEnd   = sEndField   ? fnNormDate(oRowData[sEndField])   : "";
                if (sStart && sEnd && sEnd < sStart) {
                    if (sStartField) { oRowData["_err_" + sStartField] = true; }
                    if (sEndField)   { oRowData["_err_" + sEndField]   = true; }
                    bDateError = true;
                }

                if (sQuotaQtyField && sConsumedQtyField) {
                    var fQuotaQty = parseFloat(String(oRowData[sQuotaQtyField] || "0").replace(/,/g, ""));
                    var fConsumedQty = parseFloat(String(oRowData[sConsumedQtyField] || "0").replace(/,/g, ""));
                    if (!isNaN(fQuotaQty) && !isNaN(fConsumedQty) && fQuotaQty < fConsumedQty) {
                        oRowData["_err_" + sQuotaQtyField] = true;
                        bQuotaConsumedError = true;
                    }
                }
            });

            oModel.setProperty("/rows", aRows);

            if (bRequiredError) {
                MessageBox.error(oBundle.getText("msgRequiredFields"));
                return;
            }

            if (bDateError) {
                MessageBox.error(oBundle.getText("msgDateRangeError"));
                return;
            }

            if (bQuotaConsumedError) {
                MessageBox.error("Quota Qty must be greater than or equal to Cnsmd QTy.");
                return;
            }

            // --- Date overlap validation across all rows with same key fields ---
            // Key fields = all columns up to and including the last STATUS column (Constraint Status)
            var iCsIdx = -1;
            aColumns.forEach(function (oCol, iIdx) {
                if (oCol.name.toUpperCase().indexOf("STATUS") !== -1) { iCsIdx = iIdx; }
            });
            var aNonKey = ["PRODALLOCPERDSTARTUTCDATE", "PRODALLOCPERIODENDUTCDATE",
                           "PRODUCTALLOCATIONQUANTITY", "ZZRFCUT",
                           "PRODALLOCCHARCVALUECOMBNCMNT", "PRODUCTALLOCATIONOBJECTUUID"];
            var aKeyFields = (iCsIdx >= 0
                ? aColumns.slice(0, iCsIdx + 1)
                : aColumns
            ).filter(function (c) {
                var u = c.name.toUpperCase();
                return aNonKey.indexOf(u) === -1 &&
                       u.indexOf("STATUS") === -1 &&
                       u.indexOf("AVBL")   === -1 &&
                       u.indexOf("CNSMD")  === -1 &&
                       !oController._isProdDescColumn(c);
            }).map(function (c) { return c.name; });

            var oGroups = {};
            aRows.forEach(function (oRow, iIdx) {
                var sGroupKey = aKeyFields.map(function (f) { return oRow[f] || ""; }).join("|");
                if (!oGroups[sGroupKey]) { oGroups[sGroupKey] = []; }
                oGroups[sGroupKey].push({ idx: iIdx, row: oRow });
            });

            var bOverlapError = false;
            Object.keys(oGroups).forEach(function (sGK) {
                var aGrp = oGroups[sGK];
                if (aGrp.length < 2) { return; }
                for (var ii = 0; ii < aGrp.length; ii++) {
                    for (var jj = ii + 1; jj < aGrp.length; jj++) {
                        var rA = aGrp[ii].row, rB = aGrp[jj].row;
                        var asStart = sStartField ? fnNormDate(rA[sStartField]) : "";
                        var asEnd   = sEndField   ? fnNormDate(rA[sEndField])   : "";
                        var bsStart = sStartField ? fnNormDate(rB[sStartField]) : "";
                        var bsEnd   = sEndField   ? fnNormDate(rB[sEndField])   : "";
                        if (asStart && asEnd && bsStart && bsEnd) {
                            if (asStart <= bsEnd && bsStart <= asEnd) {
                                aRows[aGrp[ii].idx]._overlapError = true;
                                aRows[aGrp[jj].idx]._overlapError = true;
                                bOverlapError = true;
                            }
                        }
                    }
                }
            });

            if (bOverlapError) {
                oModel.setProperty("/rows", aRows);
                MessageBox.error(oBundle.getText("msgDateOverlapError"));
                return;
            }

            var sFecIni = oModel.getProperty("/fec_ini");

            oModel.setProperty("/busy", true);
            MessageToast.show(oBundle.getText("msgSaving"));

            var that = this;
            var aPayloadItems = this._buildPayloadArray(aChangedRows, sFecIni);

            this._executePost(aPayloadItems)
                .then(function (oSapMsg) {
                    if (oSapMsg && oSapMsg.text) {
                        var sType = "Information";
                        var sSev = (oSapMsg.severity || "").toLowerCase();
                        if (sSev === "success") { sType = "Success"; }
                        else if (sSev === "warning" || sSev === "w") { sType = "Warning"; }
                        else if (sSev === "error" || sSev === "e") { sType = "Error"; }
                        oModel.setProperty("/messageText", oSapMsg.text);
                        oModel.setProperty("/messageType", sType);
                        oModel.setProperty("/messageVisible", true);
                        if (sType === "Error") {
                            that._restoreDeletedRowsAfterSaveError();
                            oModel.setProperty("/busy", false);
                            return;
                        }
                    } else {
                        oModel.setProperty("/messageText", oBundle.getText("msgSaveSuccess"));
                        oModel.setProperty("/messageType", "Success");
                        oModel.setProperty("/messageVisible", true);
                    }
                    that._hasDeletedRows = false;
                    that._aDeletedRows = [];
                    var sObj = oModel.getProperty("/productAllocationObject");
                    oModel.setProperty("/busy", true);
                    that._loadDynamicFields(sObj);
                })
                .catch(function (oError) {
                    that._restoreDeletedRowsAfterSaveError();
                    oModel.setProperty("/busy", false);
                    var sMsg = oBundle.getText("msgSaveError");
                    try {
                        var oResp = JSON.parse(oError.responseText);
                        sMsg = oResp.error.message.value || sMsg;
                    } catch (e) {}
                    MessageBox.error(sMsg);
                });
        },

        _restoreDeletedRowsAfterSaveError: function () {
            if (!this._aDeletedRows || this._aDeletedRows.length === 0) { return; }

            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows") || [];
            var aDeletedRows = this._aDeletedRows.slice().sort(function (a, b) {
                return a.rowIndex - b.rowIndex;
            });

            aDeletedRows.forEach(function (oDeletedEntry) {
                var iIndex = Math.min(oDeletedEntry.rowIndex, aRows.length);
                var oRowCopy = JSON.parse(JSON.stringify(oDeletedEntry.rowData));
                aRows.splice(iIndex, 0, oRowCopy);
            });

            this._aDeletedRows = [];
            this._hasDeletedRows = false;
            oModel.setProperty("/rows", aRows);
            oModel.setProperty("/rowCount", Math.min(aRows.length, 15));
            oModel.setProperty("/hasChanges", this._getChangedRows().length > 0);
        },

        _buildPayloadArray: function (aChangedRows, sFecIni) {
            var that = this;
            var oModel = this.getView().getModel("detailModel");
            var aColumns = oModel.getProperty("/columns") || [];
            var oFieldsMap = {};

            aColumns.forEach(function (oCol) {
                var sFieldName = oCol.name;
                var oFieldMeta = that._oFieldMetadata[sFieldName] || {};

                oFieldsMap[sFieldName] = {
                    name: sFieldName,
                    tablename: oFieldMeta.tabname || "PAL",
                    description: oCol.label || "",
                    position: oFieldMeta.position || "",
                    key: "",
                    type: "C",
                    length: "",
                    fec_ini: that._toODataDate(sFecIni) || "",
                    fec_fin: "",
                    DataSetAsoc: []
                };
            });

            // --- Determine ind_ope='Q' for rows sharing material+centro+startDate+endDate ---
            var sMaterialField = null, sCentroField = null;
            aColumns.forEach(function (oCol) {
                var lbl = (oCol.label || "").toLowerCase();
                if (!sMaterialField && lbl.indexOf("material") !== -1) { sMaterialField = oCol.name; }
                if (!sCentroField  && (lbl.indexOf("centro") !== -1 || lbl.indexOf("plant") !== -1)) { sCentroField = oCol.name; }
            });
            var sStartField = "PRODALLOCPERDSTARTUTCDATE";
            var sEndField   = "PRODALLOCPERIODENDUTCDATE";

            var aAllRows = oModel.getProperty("/rows") || [];
            var oGroupCount = {};
            aAllRows.forEach(function (oRow) {
                var sGK = (sMaterialField ? (oRow[sMaterialField] || "") : "") + "|" +
                          (sCentroField   ? (oRow[sCentroField]   || "") : "") + "|" +
                          (oRow[sStartField] || "") + "|" + (oRow[sEndField] || "");
                oGroupCount[sGK] = (oGroupCount[sGK] || 0) + 1;
            });
            var oQRows = {};
            aChangedRows.forEach(function (oChangedRow) {
                var d = oChangedRow.rowData;
                var sGK = (sMaterialField ? (d[sMaterialField] || "") : "") + "|" +
                          (sCentroField   ? (d[sCentroField]   || "") : "") + "|" +
                          (d[sStartField] || "") + "|" + (d[sEndField] || "");
                var bDatesModified = d["_isNew"] ||
                    (d[sStartField + "_old"] != null && d[sStartField + "_old"] !== d[sStartField]) ||
                    (d[sEndField   + "_old"] != null && d[sEndField   + "_old"] !== d[sEndField]);
                if (oGroupCount[sGK] > 1 && bDatesModified) { oQRows[oChangedRow.rowIndex] = true; }
            });
            // -----------------------------------------------------------------------

            aChangedRows.forEach(function (oChangedRow) {
                var iRowIndex = oChangedRow.rowIndex;
                var oRowData = oChangedRow.rowData;

                aColumns.forEach(function (oCol) {
                    var sFieldName = oCol.name;
                    var sCellKey = iRowIndex + "_" + sFieldName;
                    var oCellMeta = that._oCellKeys[sCellKey] || {};

                    var sCurrentValue = oRowData[sFieldName] || "";
                    var sOldRaw = oRowData[sFieldName + "_old"];
                    var sOldValue = oRowData["_isNew"] ? "" : (sOldRaw != null ? String(sOldRaw) : sCurrentValue);
                    if (sFieldName.toUpperCase() === "PRODALLOCATIONACTIVATIONSTATUS") {
                        var oStatusEn = { "activos": "Active", "inactivos": "Inactive", "active": "Active", "inactive": "Inactive" };
                        sCurrentValue = oStatusEn[sCurrentValue.toLowerCase()] || sCurrentValue;
                        sOldValue    = oStatusEn[sOldValue.toLowerCase()]    || sOldValue;
                    }

                    var oDataItem = {
                        key: oCellMeta.key || "",
                        tabname: oCellMeta.tabname || "PAL",
                        name: sFieldName,
                        Value: sCurrentValue,
                        Value_old: sOldValue,
                        position: String(iRowIndex + 1),
                        prodallocationtimeseriesuuid: oCellMeta.prodallocationtimeseriesuuid || oRowData.prodallocationtimeseriesuuid || "",
                        productallocationobject: oCellMeta.productallocationobject || oRowData.productallocationobject || "",
                        CHARCVALUECOMBINATIONUUID: oCellMeta.CHARCVALUECOMBINATIONUUID || oRowData.CHARCVALUECOMBINATIONUUID || "",
                        PRODALLOCPERDSTARTUTCDATETIME: oCellMeta.PRODALLOCPERDSTARTUTCDATETIME || oRowData.PRODALLOCPERDSTARTUTCDATETIME || "",
                        PRODALLOCPERIODENDUTCDATETIME: oCellMeta.PRODALLOCPERIODENDUTCDATETIME || oRowData.PRODALLOCPERIODENDUTCDATETIME || "",
                        productallocationsequence: oCellMeta.productallocationsequence || oRowData.productallocationsequence || "",
                        fec_ini: that._toODataDate(sFecIni) || "",
                        ind_ope: oQRows[iRowIndex] ? "Q" : (oCellMeta.ind_ope || "")
                    };

                    oFieldsMap[sFieldName].DataSetAsoc.push(oDataItem);
                });
            });

            this._aDeletedRows.forEach(function (oDeletedEntry) {
                var iRowIndex = oDeletedEntry.rowIndex;
                var oRowData = oDeletedEntry.rowData;

                aColumns.forEach(function (oCol) {
                    var sFieldName = oCol.name;
                    var sCellKey = iRowIndex + "_" + sFieldName;
                    var oCellMeta = that._oCellKeys[sCellKey] || {};

                    var sDelOldValue = oRowData[sFieldName + "_old"] || oRowData[sFieldName] || "";
                    if (sFieldName.toUpperCase() === "PRODALLOCATIONACTIVATIONSTATUS") {
                        var oStatusEnDel = { "activos": "Active", "inactivos": "Inactive", "active": "Active", "inactive": "Inactive" };
                        sDelOldValue = oStatusEnDel[sDelOldValue.toLowerCase()] || sDelOldValue;
                    }
                    var oDataItem = {
                        key: oCellMeta.key || "",
                        tabname: oCellMeta.tabname || "PAL",
                        name: sFieldName,
                        Value: "",
                        Value_old: sDelOldValue,
                        position: String(iRowIndex + 1),
                        prodallocationtimeseriesuuid: oCellMeta.prodallocationtimeseriesuuid || oRowData.prodallocationtimeseriesuuid || "",
                        productallocationobject: oCellMeta.productallocationobject || oRowData.productallocationobject || "",
                        CHARCVALUECOMBINATIONUUID: oCellMeta.CHARCVALUECOMBINATIONUUID || oRowData.CHARCVALUECOMBINATIONUUID || "",
                        PRODALLOCPERDSTARTUTCDATETIME: oCellMeta.PRODALLOCPERDSTARTUTCDATETIME || oRowData.PRODALLOCPERDSTARTUTCDATETIME || "",
                        PRODALLOCPERIODENDUTCDATETIME: oCellMeta.PRODALLOCPERIODENDUTCDATETIME || oRowData.PRODALLOCPERIODENDUTCDATETIME || "",
                        productallocationsequence: oCellMeta.productallocationsequence || oRowData.productallocationsequence || "",
                        fec_ini: that._toODataDate(sFecIni) || "",
                        ind_ope: oCellMeta.ind_ope || ""
                    };

                    if (oFieldsMap[sFieldName]) {
                        oFieldsMap[sFieldName].DataSetAsoc.push(oDataItem);
                    }
                });
            });

            var aPayload = [];
            for (var sKey in oFieldsMap) {
                if (oFieldsMap.hasOwnProperty(sKey) && oFieldsMap[sKey].DataSetAsoc.length > 0) {
                    aPayload.push(oFieldsMap[sKey]);
                }
            }

            return aPayload;
        },

        _executePost: function (aPayload) {
            var oODataModel = this.getOwnerComponent().getModel();

            var oData = {
                name: "",
                tablename: "PAL",
                DataSetAsoc: []
            };

            aPayload.forEach(function (oFieldSet) {
                var aDataItems = oFieldSet.DataSetAsoc || [];
                aDataItems.forEach(function (oItem) {
                    oData.DataSetAsoc.push(oItem);
                });
            });

            console.log("[Save] POST /DynamicFieldSet payload:", JSON.stringify(oData, null, 2));

            return new Promise(function (resolve, reject) {
                oODataModel.create("/DynamicFieldSet", oData, {
                    success: function (oData, oResponse) {
                        var oSapMsg = null;
                        try {
                            var sHeader = oResponse && oResponse.headers && oResponse.headers["sap-message"];
                            if (sHeader) {
                                var oParsed = JSON.parse(sHeader);
                                oSapMsg = { text: oParsed.message || "", severity: oParsed.severity || "" };
                            }
                        } catch (e) {}
                        resolve(oSapMsg);
                    },
                    error: function (oError) {
                        reject(oError);
                    }
                });
            });
        }

    });
});
