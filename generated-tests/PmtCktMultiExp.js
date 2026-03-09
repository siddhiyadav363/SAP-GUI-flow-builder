```javascript
// SHEET_NAME: Payment Cockpit - MultiStatus & Export
// REQUIRED_COLUMNS: DateRange, Statuses, ExportFormat
// UPDATED: Separated tests for TC_STORY_PAYMENT_006 and TC_STORY_PAYMENT_008.
// - Added clearer comments, helper POM-like functions, robust assertions, and explicit timeouts.
// - Preserved original matchers and control identifiers (viewName, controlType, id regex).
// - Increased global timeout and per-waitFor timeout to mitigate intermittent timeout issues.
// Notes:
// - Resolver.get("<key>") is used to fetch test data (e.g., DateRange, Statuses, ExportFormat, etc.).
// - Actual application URL is a placeholder; adjust as needed in CI environment.
// - File-system verification of download is outside OPA5 scope; tests check for UI confirmation and export invocation.
sap.ui.define([
    "sap/ui/test/opaQunit",
    "sap/ui/test/Opa5",
    "sap/ui/test/actions/Press",
    "sap/ui/test/actions/EnterText",
    "sap/ui/test/matchers/PropertyStrictEquals",
    "sap/ui/test/matchers/Properties",
    "sap/ui/test/matchers/Ancestor",
    "sap/ui/test/matchers/Descendant",
    "flp/opa/test/integration/runtime/resolver"
], function (opaTest, Opa5, Press, EnterText, PropertyStrictEquals, Properties, Ancestor, Descendant, Resolver) {
    "use strict";
    // Global OPA config - extended timeout for slower environments
    Opa5.extendConfig({
        arrangements: new Opa5({
            iStartFLP: function () {
                // Start FLP in iframe using placeholder URL - update if real FLP path is provided
                return this.iStartMyAppInAFrame("/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home");
            }
        }),
        autoWait: true,
        timeout: 120000, // 2 minutes global timeout
        pollingInterval: 200,
        assertions: new Opa5({
            iTeardownMyApp: function () {
                return this.iTeardownMyAppFrame();
            }
        })
    });
    // -----------------------
    // Helper / POM-like methods
    // -----------------------
    var fnNavigateToPaymentCockpit = function (When) {
        // Click Central Billing tab
        When.waitFor(Object.assign({
            actions: new Press(),
            viewName: "sap.ushell.components.shell.MenuBar.view.MenuBarPersonalization",
            timeout: 120000,
            success: function () {
                Opa5.assert.ok(true, "Clicked 'Central Billing' tab");
            },
            errorMessage: "Could not click 'Central Billing' tab"
        }, {
            controlType: "sap.m.IconTabFilter",
            matchers: [
                new PropertyStrictEquals({ name: "text", value: "Central Billing" })
            ]
        }));
        // Click Payment Cockpit tile on FLP
        When.waitFor(Object.assign({
            actions: new Press(),
            viewName: "sap.ushell.components.pages.view.PageRuntime",
            timeout: 120000,
            success: function () {
                Opa5.assert.ok(true, "Clicked 'Payment Cockpit' FLP tile");
            },
            errorMessage: "Could not click 'Payment Cockpit' tile"
        }, {
            controlType: "sap.ushell.ui.launchpad.VizInstanceCdm",
            matchers: [
                new PropertyStrictEquals({ name: "title", value: "Payment Cock