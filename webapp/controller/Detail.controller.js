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
    "sap/m/DatePicker",
    "sap/ui/table/Column",
    "sap/ui/table/RowSettings",
    "sap/ui/core/format/DateFormat"
], function (Controller, History, JSONModel, Filter, FilterOperator, MessageBox, MessageToast, Text, Input, Label, DatePicker, UIColumn, RowSettings, DateFormat) {
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
                    that.getView().getModel("detailModel").setProperty("/hasChanges", false);
                    that._doNavBack();
                });
                return;
            }
            var oModel = this.getView().getModel("detailModel");
            oOwner._oDetailModel = oModel;
            oModel.setProperty("/productAllocationObject", sQuotaId);
            oModel.setProperty("/busy", true);
            oModel.setProperty("/messageVisible", false);
            oModel.setProperty("/messageText", "");
            oModel.setProperty("/messageType", "None");
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

                    that._oCellKeys = {};

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
                                productallocationsequence: oEntry.productallocationsequence || ""
                            };

                            if (!that._oFieldMetadata[oField.name]) {
                                that._oFieldMetadata[oField.name] = {
                                    tabname: oEntry.tabname || oField.tablename || "",
                                    position: oField.position || ""
                                };
                            }
                        });
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

            var oDateFormat = DateFormat.getDateInstance({
                style: "medium"
            });

            aColumns.forEach(function (oCol) {
                var sFieldName = oCol.name;
                var sFieldUpper = sFieldName.toUpperCase();

                if (sFieldUpper === "PRODUCTALLOCATIONOBJECTUUID") {
                    return;
                }
                
                var bNonEditableText = (sFieldUpper === "PRODUCTALLOCATIONOBJECT");
                
                var sLabelUpper = (oCol.label || "").toUpperCase().trim();
                var bNeverEditable = (sLabelUpper === "AVBL QTY" || sLabelUpper === "CNSMD QTY");

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
                        editable: "{= ${detailModel>_isNew} === true }",
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
                    var oInputCfg = {
                        value: "{detailModel>" + sFieldName + "}",
                        change: that._onFieldChange.bind(that),
                        liveChange: that._onFieldChange.bind(that)
                    };
                    if (!bIsComment) {
                        oInputCfg.valueState = "{= ${detailModel>_err_" + sFieldName + "} ? 'Error' : 'None' }";
                        oInputCfg.required = true;
                    }
                    oTemplate = new Input(oInputCfg).addStyleClass("sapUiSizeCompact");
                } else {
                    oTemplate = new Input({
                        value: "{detailModel>" + sFieldName + "}",
                        editable: "{= ${detailModel>_isNew} === true }",
                        required: "{= ${detailModel>_isNew} === true }",
                        valueState: "{= ${detailModel>_err_" + sFieldName + "} ? 'Error' : 'None' }",
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
                this._showUnsavedPopup(function () { that._doNavBack(); });
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
                afterClose: function () { oDialog.destroy(); }
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

            oNewRow["PRODUCTALLOCATIONOBJECT"] = sProductAllocationObject;
            oNewRow["PRODUCTALLOCATIONOBJECT_old"] = sProductAllocationObject;
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
                MessageToast.show("Seleccione al menos una fila para eliminar");
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
                this._aDeletedRows.push({ rowIndex: iIndex, rowData: oRowCopy });
                aRows.splice(iIndex, 1);
                if (this._oOriginalData && this._oOriginalData[iIndex]) {
                    this._oOriginalData.splice(iIndex, 1);
                }
            }

            oModel.setProperty("/rows", aRows);

            var iNewRowCount = Math.max(aRows.length, 5);
            oModel.setProperty("/rowCount", Math.min(iNewRowCount, 15));

            oTable.clearSelection();

            this._hasDeletedRows = true;
            oModel.setProperty("/hasChanges", true);
        },

        onCopyRow: function () {
            var oTable = this.byId("idDynamicTable");
            var aSelectedIndices = oTable.getSelectedIndices();

            if (aSelectedIndices.length === 0) {
                MessageToast.show("Seleccione una fila para copiar");
                return;
            }

            if (aSelectedIndices.length > 1) {
                MessageToast.show("Solo puede copiar una fila a la vez");
                return;
            }

            var oModel = this.getView().getModel("detailModel");
            var aRows = oModel.getProperty("/rows") || [];
            var iSelectedIndex = aSelectedIndices[0];
            var oSelectedRow = aRows[iSelectedIndex];

            var oNewRow = JSON.parse(JSON.stringify(oSelectedRow));

            var aColumns = oModel.getProperty("/columns") || [];
            aColumns.forEach(function (oCol) {
                oNewRow[oCol.name + "_old"] = "";
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
                        new sap.m.MenuItem({ text: bEnC ? "Restricted Availability"   : "Disponibilidad restringida",      press: that.onChangeConstraintStatus.bind(that) }),
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
                        that.onNavBack();
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

            var aNonRequired = [
                "PRODALLOCCHARCVALUECOMBNCMNT",
                "PRODALLOCATIONACTIVATIONSTATUS",
                "PRODALLOCCHARCCONSTRAINTSTATUS",
                "PRODUCTALLOCATIONOBJECT",
                "PRODUCTALLOCATIONOBJECTUUID"
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

            var sStartField = null, sEndField = null;
            aColumns.forEach(function (oCol) {
                var u = oCol.name.toUpperCase();
                if (u === "PRODALLOCPERDSTARTUTCDATE") { sStartField = oCol.name; }
                if (u === "PRODALLOCPERIODENDUTCDATE")  { sEndField   = oCol.name; }
            });

            aChangedRows.forEach(function (oChangedRow) {
                var oRowData = oChangedRow.rowData;

                if (oRowData._isNew) {
                    aColumns.forEach(function (oCol) {
                        var sFieldName = oCol.name;
                        var sFieldUpper = sFieldName.toUpperCase();
                        var sColLabelUpper = (oCol.label || "").toUpperCase().trim();
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
                : aColumns.filter(function (c) {
                    var u = c.name.toUpperCase();
                    return aNonKey.indexOf(u) === -1 && u.indexOf("AVBL") === -1 && u.indexOf("CNSMD") === -1;
                })
            ).map(function (c) { return c.name; });

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

            console.log("POST Payload Array:", JSON.stringify(aPayloadItems, null, 2));

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
                    oModel.setProperty("/busy", false);
                    var sMsg = oBundle.getText("msgSaveError");
                    try {
                        var oResp = JSON.parse(oError.responseText);
                        sMsg = oResp.error.message.value || sMsg;
                    } catch (e) {}
                    MessageBox.error(sMsg);
                });
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
                        fec_ini: that._toODataDate(sFecIni) || ""
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
                        fec_ini: that._toODataDate(sFecIni) || ""
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

            console.log("=== POST REQUEST ===");
            console.log("POST Path: /DynamicFieldSet");
            console.log("POST Payload:", JSON.stringify(aPayload, null, 2));
            console.log("===================");

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

            console.log("Flattened DataSetAsoc:", JSON.stringify(oData, null, 2));

            return new Promise(function (resolve, reject) {
                oODataModel.create("/DynamicFieldSet", oData, {
                    success: function (oData, oResponse) {
                        console.log("POST Success");
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
                        console.log("POST Error:", oError);
                        reject(oError);
                    }
                });
            });
        }

    });
});
