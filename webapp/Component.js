sap.ui.define([
    "sap/ui/core/UIComponent",
    "sap/ui/Device",
    "t_project1/model/models"
], function(UIComponent, Device, models) {
    "use strict";

    return UIComponent.extend("t_project1.Component", {
        metadata: {
            manifest: "json",
            config: {
                fullWidth: true
            }
        },

        init: function() {
            UIComponent.prototype.init.apply(this, arguments);

            this.getRouter().initialize();

            this.setModel(models.createDeviceModel(), "device");
        }
    });
});
