// ## PrecisionScript
import { test, expect, Page } from '@playwright/test';
// ## PrecisionScript
test.use({
  ignoreHTTPSErrors: true,
  httpCredentials: {
    username: '{{username}}',
    password: '{{password}}'
  }
});

test.describe('Fiori Sales Order Creation', () => {
  test('Create Standard Sales Order with Product', async ({ page }) => {
    test.setTimeout(320_000);

    // Step 1: Navigate to Fiori Launchpad
    await page.goto('{{fioriLaunchUrl}}');

    // Step 2: Enter username
    await page.getByRole('textbox', { name: 'User' }).click();
    await page.getByRole('textbox', { name: 'User' }).fill('{{username}}');

    // Step 3: Enter password
    await page.getByRole('textbox', { name: 'Password' }).click();
    await page.getByRole('textbox', { name: 'Password' }).fill('{{password}}');

    // Step 4: Click Log On button
    await page.getByRole('button', { name: 'Log On' }).click();

    // Step 5: Navigate to home after login
    // await page.goto('{{fioriLaunchUrl}}');

    // Step 6: Navigate to Central Billing
    await page.getByText('Central Billing').click();

    // Step 7: Open Create/Edit Orders app
    await page.getByRole('link', { name: 'Create/Edit Orders Navigation' }).click();

    // Step 8: Navigate to Manage Sales display
    // await page.goto('{{fioriLaunchUrl}}');

    // Step 9: Click Create Sales Order button
    await page.getByRole('button', { name: 'Create Sales Order' }).click();

    // Step 10: Select Sales Order Type
    await page.locator('[id="APD_::SalesOrderType-inner-vhi"]').click();
    await page.getByText('Standard Order (ZPFU)').click();

    // Step 11: Select Sales Organization
    await page.getByRole('textbox', { name: 'Sales Organization' }).click();
    await page.locator('[id="APD_::SalesOrganization-inner-vhi"]').click();
    await page.getByText('Freeman Expo US').click();

    // Step 12: Select Distribution Channel
    await page.locator('[id="APD_::DistributionChannel-inner-vhi"]').click();
    await page.getByText('Transactional').click();

    // Step 13: Select Organization Division
    await page.locator('[id="APD_::OrganizationDivision-inner-vhi"]').click();
    await page.getByText('Common').click();

    // Step 14: Click Create Sales Order in footer
    await page.getByLabel('Footer actions').getByRole('button', { name: 'Create Sales Order' }).click();

    // Step 15: Select Sold-to Party
    await page.locator('[id="cus.sd.salesorderv2.manage::SalesOrderManageObjectPage--fe::FormContainer::OrderData::FormElement::DataField::SoldToParty::Field-edit-inner-vhi"]').click();
    await page.getByText('EPE100000004406').click();

    // Step 16: Fill Ordering Person First Name
    await page.getByRole('form', { name: 'Ordering Person Contact' }).getByLabel('First Name').click();
    await page.getByRole('form', { name: 'Ordering Person Contact' }).getByLabel('First Name').fill('{{contact_first_name}}');

    // Step 17: Fill Ordering Person Last Name
    await page.getByRole('form', { name: 'Ordering Person Contact' }).getByLabel('Last Name').click();
    await page.getByRole('form', { name: 'Ordering Person Contact' }).getByLabel('Last Name').fill('{{contact_last_name}}');

    // Step 18: Fill Ordering Person Email
    await page.getByRole('form', { name: 'Ordering Person Contact' }).getByLabel('Email').click();
    await page.getByRole('form', { name: 'Ordering Person Contact' }).getByLabel('Email').fill('{{contact_email}}');

    // Step 19: Fill Bill-To Person Email
    await page.locator('[id="__input33-inner"]').click();
    await page.locator('[id="__input33-inner"]').fill('{{bill_to_email}}');

    // Step 20: Navigate to Items tab
    await page.getByRole('tab', { name: 'Items' }).click();

    // Step 21: Open Event Product Catalog
    await page.getByRole('button', { name: 'Event Product Catalog' }).click();

    // Step 22: Search for material
    await page.getByRole('searchbox', { name: 'Search by Material' }).click();
    await page.getByRole('searchbox', { name: 'Search by Material' }).fill('{{material_number}}');

    // Step 23: Click Go to search
    await page.getByRole('button', { name: 'Go', exact: true }).click();

    // Step 24: Select product row
    await page.locator('[id="cus.sd.salesorderv2.manage::SalesOrderManageObjectPage--TAB_products-rows-row0-col0"]').click();

    // Step 25: Click quantity cell
    await page.locator('[id="cus.sd.salesorderv2.manage::SalesOrderManageObjectPage--TAB_products-rows-row0-col4"] > .sapUiTableCellInner').click();

    // Step 26: Fill quantity
    await page.locator('[id="__input55-__clone513-inner"]').fill('1');

    // Step 27: Click Select button
    await page.getByRole('button', { name: 'Select' }).click();

    // Step 28: Click Create button
    await page.getByRole('button', { name: 'Create' }).click();

    // Step 29: Click OK on confirmation
    await page.getByRole('button', { name: 'OK' }).click();

    // Step 30: Capture Sales Order number
    const salesOrderNumber = await page.locator('[id="__title5-inner"]').innerText();
    console.log('Output: ' + String(salesOrderNumber).trim());

    // Step 31: Click Sales Order title to verify
    await page.locator('[id="__title5-inner"]').click();
  });
});

/*
 * ═══════════════════════════════════════════════════════════════════
 * ELEMENT CONTEXT  (captured during recording — locators for reference)
 * Copy-paste these into your Playwright script as needed.
 * ═══════════════════════════════════════════════════════════════════
 * 
 * ┌─ Page : (untitled)
 * │  URL  : chrome-error://chromewebdata/
 * │  Time : 2026-04-28T08:41:50.957Z
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : (untitled)
 * │  URL  : chrome-error://chromewebdata/
 * │  Time : 2026-04-28T08:41:54.648Z
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Logon
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:41:57.750Z
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('User')  // {{user}}  [REQUIRED]
 * │    page.getByLabel('Password')  // {{password}}  [REQUIRED]
 * │
 * │  ── Dropdowns
 * │    page.getByLabel('Language')
 * │    page.selectOption(page.getByLabel('Language'), '<value>')
 * │      // Options: 'DE - Deutsch', 'EN - English', 'FR - Français'
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Log On' })
 * │    page.getByRole('button', { name: 'Change Password' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Logon
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:02.590Z
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('User')  // {{user}}  [REQUIRED]
 * │
 * │  ── Dropdowns
 * │    page.getByLabel('Language')
 * │    page.selectOption(page.getByLabel('Language'), '<value>')
 * │      // Options: 'DE - Deutsch', 'EN - English', 'FR - Français'
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Log On' })
 * │    page.getByRole('button', { name: 'Change Password' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Logon
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:06.573Z
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('User')  // {{user}}  [REQUIRED]
 * │
 * │  ── Dropdowns
 * │    page.getByLabel('Language')
 * │    page.selectOption(page.getByLabel('Language'), '<value>')
 * │      // Options: 'DE - Deutsch', 'EN - English', 'FR - Français'
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Log On' })
 * │    page.getByRole('button', { name: 'Change Password' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Logon
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:10.610Z
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('User')  // {{user}}  [REQUIRED]
 * │    page.getByLabel('Password')  // {{password}}  [REQUIRED]
 * │
 * │  ── Dropdowns
 * │    page.getByLabel('Language')
 * │    page.selectOption(page.getByLabel('Language'), '<value>')
 * │      // Options: 'DE - Deutsch', 'EN - English', 'FR - Français'
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Log On' })
 * │    page.getByRole('button', { name: 'Change Password' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Loading
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?_sap-hash=JTIzU2hlbGwtaG9tZQ&sap-system-login=X&sap-system-login-cookie=X&sap-contextid=SID:ANON:vhtfmqsrci_QSR_00:McIzWmzUVCq7M_zw2xWxJjHxYVadCKr3Dbt33aUW-ATT
 * │  Time : 2026-04-28T08:42:13.129Z
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : (untitled)
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:14.618Z
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : (untitled)
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:18.650Z
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : (untitled)
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:22.679Z
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:26.683Z
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Edit Page' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-04-28T08:42:31.019Z
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Edit Page' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Launchpad-openFLPPage?pageId=ZP_FIN_W_CNTRL_BLNG&spaceId=ZS_FIN_W_CNTRL_BLNG
 * │  Time : 2026-04-28T08:42:34.730Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │    page.getByRole('button', { name: 'Edit Page' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Launchpad-openFLPPage?pageId=ZP_FIN_W_CNTRL_BLNG&spaceId=ZS_FIN_W_CNTRL_BLNG
 * │  Time : 2026-04-28T08:42:38.719Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Display Business Partner BUP3 Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Document Navigation Tile' })
 * │    page.getByRole('link', { name: 'List Blocked Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'My Inbox All Items ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Payment Cockpit Navigation Tile' })
 * │    page.getByRole('link', { name: 'PDF Generation Dashboard Navigation Tile' })
 * │    page.getByRole('link', { name: 'Sales Documents Blocked for Billing Navigation Tile' })
 * │    page.getByRole('link', { name: 'General Invoice Report General Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Outstanding Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Proforma Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Invoice Search ZINV_PRINT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Customer List Navigation Tile' })
 * │    page.getByRole('link', { name: 'WBS Element Overview Navigation Tile' })
 * │    page.getByRole('link', { name: 'Cancel Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Clear Incoming Payments Manual Clearing 211 Open Payments Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create Billing Documents ... Billing Due List Items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Exhibitor Portal Setup Exhibitor/Booths Navigation Tile' })
 * │    page.getByRole('link', { name: 'File Upload of Billiable items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Import Sales Order - EXT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Document Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Debit Memo Requests ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Customer Down Payment Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Sales Orders Billing Block Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create/Edit Orders Navigation Tile' })
 * │    page.getByRole('link', { name: 'Maintain Billing Due List Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Journal Entries New Version Recommended Navigation Tile' })
 * │    page.getByRole('link', { name: 'Event Statement Print ZEVENT_STMT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Credit Decisions Blocked SD Documents Navigation Tile' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:42:42.732Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Display Business Partner BUP3 Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Document Navigation Tile' })
 * │    page.getByRole('link', { name: 'List Blocked Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'My Inbox All Items ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Payment Cockpit Navigation Tile' })
 * │    page.getByRole('link', { name: 'PDF Generation Dashboard Navigation Tile' })
 * │    page.getByRole('link', { name: 'Sales Documents Blocked for Billing Navigation Tile' })
 * │    page.getByRole('link', { name: 'General Invoice Report General Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Outstanding Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Proforma Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Invoice Search ZINV_PRINT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Customer List Navigation Tile' })
 * │    page.getByRole('link', { name: 'WBS Element Overview Navigation Tile' })
 * │    page.getByRole('link', { name: 'Cancel Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Clear Incoming Payments Manual Clearing 211 Open Payments Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create Billing Documents ... Billing Due List Items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Exhibitor Portal Setup Exhibitor/Booths Navigation Tile' })
 * │    page.getByRole('link', { name: 'File Upload of Billiable items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Import Sales Order - EXT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Document Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Debit Memo Requests ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Customer Down Payment Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Sales Orders Billing Block Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create/Edit Orders Navigation Tile' })
 * │    page.getByRole('link', { name: 'Maintain Billing Due List Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Journal Entries New Version Recommended Navigation Tile' })
 * │    page.getByRole('link', { name: 'Event Statement Print ZEVENT_STMT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Credit Decisions Blocked SD Documents Navigation Tile' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:42:55.566Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Display Business Partner BUP3 Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Document Navigation Tile' })
 * │    page.getByRole('link', { name: 'List Blocked Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'My Inbox All Items ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Payment Cockpit Navigation Tile' })
 * │    page.getByRole('link', { name: 'PDF Generation Dashboard Navigation Tile' })
 * │    page.getByRole('link', { name: 'Sales Documents Blocked for Billing Navigation Tile' })
 * │    page.getByRole('link', { name: 'General Invoice Report General Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Outstanding Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Proforma Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Invoice Search ZINV_PRINT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Customer List Navigation Tile' })
 * │    page.getByRole('link', { name: 'WBS Element Overview Navigation Tile' })
 * │    page.getByRole('link', { name: 'Cancel Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Clear Incoming Payments Manual Clearing 211 Open Payments Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create Billing Documents ... Billing Due List Items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Exhibitor Portal Setup Exhibitor/Booths Navigation Tile' })
 * │    page.getByRole('link', { name: 'File Upload of Billiable items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Import Sales Order - EXT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Document Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Debit Memo Requests ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Customer Down Payment Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Sales Orders Billing Block Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create/Edit Orders Navigation Tile' })
 * │    page.getByRole('link', { name: 'Maintain Billing Due List Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Journal Entries New Version Recommended Navigation Tile' })
 * │    page.getByRole('link', { name: 'Event Statement Print ZEVENT_STMT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Credit Decisions Blocked SD Documents Navigation Tile' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:42:55.593Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Display Business Partner BUP3 Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Document Navigation Tile' })
 * │    page.getByRole('link', { name: 'List Blocked Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'My Inbox All Items ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Payment Cockpit Navigation Tile' })
 * │    page.getByRole('link', { name: 'PDF Generation Dashboard Navigation Tile' })
 * │    page.getByRole('link', { name: 'Sales Documents Blocked for Billing Navigation Tile' })
 * │    page.getByRole('link', { name: 'General Invoice Report General Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Outstanding Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Proforma Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Invoice Search ZINV_PRINT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Customer List Navigation Tile' })
 * │    page.getByRole('link', { name: 'WBS Element Overview Navigation Tile' })
 * │    page.getByRole('link', { name: 'Cancel Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Clear Incoming Payments Manual Clearing 211 Open Payments Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create Billing Documents ... Billing Due List Items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Exhibitor Portal Setup Exhibitor/Booths Navigation Tile' })
 * │    page.getByRole('link', { name: 'File Upload of Billiable items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Import Sales Order - EXT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Document Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Debit Memo Requests ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Customer Down Payment Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Sales Orders Billing Block Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create/Edit Orders Navigation Tile' })
 * │    page.getByRole('link', { name: 'Maintain Billing Due List Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Journal Entries New Version Recommended Navigation Tile' })
 * │    page.getByRole('link', { name: 'Event Statement Print ZEVENT_STMT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Credit Decisions Blocked SD Documents Navigation Tile' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:42:55.610Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Display Business Partner BUP3 Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Document Navigation Tile' })
 * │    page.getByRole('link', { name: 'List Blocked Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'My Inbox All Items ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Payment Cockpit Navigation Tile' })
 * │    page.getByRole('link', { name: 'PDF Generation Dashboard Navigation Tile' })
 * │    page.getByRole('link', { name: 'Sales Documents Blocked for Billing Navigation Tile' })
 * │    page.getByRole('link', { name: 'General Invoice Report General Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Outstanding Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Proforma Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Invoice Search ZINV_PRINT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Customer List Navigation Tile' })
 * │    page.getByRole('link', { name: 'WBS Element Overview Navigation Tile' })
 * │    page.getByRole('link', { name: 'Cancel Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Clear Incoming Payments Manual Clearing 211 Open Payments Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create Billing Documents ... Billing Due List Items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Exhibitor Portal Setup Exhibitor/Booths Navigation Tile' })
 * │    page.getByRole('link', { name: 'File Upload of Billiable items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Import Sales Order - EXT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Document Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Debit Memo Requests ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Customer Down Payment Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Sales Orders Billing Block Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create/Edit Orders Navigation Tile' })
 * │    page.getByRole('link', { name: 'Maintain Billing Due List Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Journal Entries New Version Recommended Navigation Tile' })
 * │    page.getByRole('link', { name: 'Event Statement Print ZEVENT_STMT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Credit Decisions Blocked SD Documents Navigation Tile' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Home
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:42:58.847Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Home' })
 * │    page.getByRole('button', { name: 'Personalize Navigation Bar' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Display Business Partner BUP3 Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Document Navigation Tile' })
 * │    page.getByRole('link', { name: 'List Blocked Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'My Inbox All Items ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Payment Cockpit Navigation Tile' })
 * │    page.getByRole('link', { name: 'PDF Generation Dashboard Navigation Tile' })
 * │    page.getByRole('link', { name: 'Sales Documents Blocked for Billing Navigation Tile' })
 * │    page.getByRole('link', { name: 'General Invoice Report General Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Outstanding Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Proforma Invoice Report Navigation Tile' })
 * │    page.getByRole('link', { name: 'Invoice Search ZINV_PRINT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Display Customer List Navigation Tile' })
 * │    page.getByRole('link', { name: 'WBS Element Overview Navigation Tile' })
 * │    page.getByRole('link', { name: 'Cancel Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Clear Incoming Payments Manual Clearing 211 Open Payments Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create Billing Documents ... Billing Due List Items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Exhibitor Portal Setup Exhibitor/Booths Navigation Tile' })
 * │    page.getByRole('link', { name: 'File Upload of Billiable items Navigation Tile' })
 * │    page.getByRole('link', { name: 'Import Sales Order - EXT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Documents Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Billing Document Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Debit Memo Requests ... Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Customer Down Payment Requests Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Sales Orders Billing Block Navigation Tile' })
 * │    page.getByRole('link', { name: 'Create/Edit Orders Navigation Tile' })
 * │    page.getByRole('link', { name: 'Maintain Billing Due List Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Journal Entries New Version Recommended Navigation Tile' })
 * │    page.getByRole('link', { name: 'Event Statement Print ZEVENT_STMT Navigation Tile' })
 * │    page.getByRole('link', { name: 'Manage Credit Decisions Blocked SD Documents Navigation Tile' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:03.766Z
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:06.832Z
 * │  Headings: Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search')  // {{search}}
 * │    page.getByLabel('Sales Order')  // {{sales_order}}
 * │    page.getByLabel('Sold-to Party')  // {{sold_to_party}}
 * │    page.getByLabel('Customer Reference')  // {{customer_reference}}
 * │    page.getByLabel('Requested Delivery Date')  // {{requested_delivery_date}}
 * │    page.getByLabel('Overall Status')  // {{overall_status}}
 * │    page.getByLabel('Document Date')  // {{document_date}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:11.199Z
 * │  Headings: Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search')  // {{search}}
 * │    page.getByLabel('Sales Order')  // {{sales_order}}
 * │    page.getByLabel('Sold-to Party')  // {{sold_to_party}}
 * │    page.getByLabel('Customer Reference')  // {{customer_reference}}
 * │    page.getByLabel('Requested Delivery Date')  // {{requested_delivery_date}}
 * │    page.getByLabel('Overall Status')  // {{overall_status}}
 * │    page.getByLabel('Document Date')  // {{document_date}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:14.868Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:19.378Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:22.936Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Freeman Upload Order (ZCSV)')
 * │    await expect(page.locator('body')).toContainText('FET order (ZFET)')
 * │    await expect(page.locator('body')).toContainText('Intercompany Order (ZIC)')
 * │    await expect(page.locator('body')).toContainText('Internal Sales Order (ZINT)')
 * │    await expect(page.locator('body')).toContainText('Freight Receiv Order (ZMHR)')
 * │    await expect(page.locator('body')).toContainText('Standard Order (ZPFU)')
 * │    await expect(page.locator('body')).toContainText('Consignment Pick-Up (ZPPU)')
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:26.898Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:30.913Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:34.916Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:39.010Z
 * │  Headings: Create Sales Order  ·  Select: Sales Organization  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │    page.getByLabel('Search')  // {{search}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Show Filters' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sales Organization')
 * │    await expect(page.locator('body')).toContainText('Sales Organization Description')
 * │    await expect(page.locator('body')).toContainText('0001')
 * │    await expect(page.locator('body')).toContainText('Sales Org. 001')
 * │    await expect(page.locator('body')).toContainText('0003')
 * │    await expect(page.locator('body')).toContainText('1001')
 * │    await expect(page.locator('body')).toContainText('Freeman Expo US')
 * │    await expect(page.locator('body')).toContainText('1003')
 * │    await expect(page.locator('body')).toContainText('Freeman Audio Visual')
 * │    await expect(page.locator('body')).toContainText('1004')
 * │    await expect(page.locator('body')).toContainText('Alford Media Service')
 * │    await expect(page.locator('body')).toContainText('1006')
 * │    await expect(page.locator('body')).toContainText('Freeman Chicago Elec')
 * │    await expect(page.locator('body')).toContainText('1014')
 * │    await expect(page.locator('body')).toContainText('Freeman XP Agency')
 * │    await expect(page.locator('body')).toContainText('1015')
 * │    await expect(page.locator('body')).toContainText('Freeman Digi Venture')
 * │    await expect(page.locator('body')).toContainText('1030')
 * │    await expect(page.locator('body')).toContainText('Exhibit Surveys, LLC')
 * │    await expect(page.locator('body')).toContainText('1094')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:42.969Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:47.004Z
 * │  Headings: Create Sales Order  ·  Select: Distribution Channel  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │    page.getByLabel('Distribution Channel')  // {{distribution_channel}}
 * │    page.getByLabel('Search')  // {{search}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Show Filters' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Distribution Channel')
 * │    await expect(page.locator('body')).toContainText('Distribution Channel Description')
 * │    await expect(page.locator('body')).toContainText('Transactional')
 * │    await expect(page.locator('body')).toContainText('Sales Support')
 * │    await expect(page.locator('body')).toContainText('Show Organizer')
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:50.993Z
 * │  Headings: Create Sales Order  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │    page.getByLabel('Distribution Channel')  // {{distribution_channel}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:55.192Z
 * │  Headings: Create Sales Order  ·  Select: Division  ·  Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Sales Order Type')  // {{sales_order_type}}  [REQUIRED]
 * │    page.getByLabel('Sales Organization')  // {{sales_organization}}
 * │    page.getByLabel('Distribution Channel')  // {{distribution_channel}}
 * │    page.getByLabel('Division')  // {{division}}
 * │    page.getByLabel('Search')  // {{search}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Show Filters' })
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Division')
 * │    await expect(page.locator('body')).toContainText('Division Description')
 * │    await expect(page.locator('body')).toContainText('Common')
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display
 * │  Time : 2026-04-28T08:43:59.143Z
 * │  Headings: Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:02.015Z
 * │  Headings: Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:07.708Z
 * │  Headings: Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:09.782Z
 * │  Headings: Standard  ·  Sales Orders
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Adapt Filters' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'Work Order Number/Ticket Number' })
 * │    page.getByRole('button', { name: 'Create Sales Order' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Create with Reference' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │
 * │  ── Tables
 * │    // Column headers: Select All | Sales Order | Sold-to Party | Customer Reference | Requested Delivery Date | Overall Status | Net Value
 * │    page.getByRole('columnheader', { name: 'Select All' })
 * │    page.getByRole('columnheader', { name: 'Sales Order' })
 * │    page.getByRole('columnheader', { name: 'Sold-to Party' })
 * │    page.getByRole('columnheader', { name: 'Customer Reference' })
 * │    page.getByRole('columnheader', { name: 'Requested Delivery Date' })
 * │    page.getByRole('columnheader', { name: 'Overall Status' })
 * │    page.getByRole('columnheader', { name: 'Net Value' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | To start, set the relevant filters and choose "Go". |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('To start, set the relevant filters and choose "Go".')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:15.845Z
 * │  Headings: New: Sales Order  ·  Event Details  ·  Partners  ·  Status  ·  Credit Limit Utilization  ·  Net Sales Volume (YTD)  ·  Net Amount  ·  Amount Details
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '–Empty Value' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Withdraw Approval Request' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Generate Payment Link' })
 * │    page.getByRole('button', { name: 'Launch Service Cloud' })
 * │    page.getByRole('button', { name: 'Simulate Taxes' })
 * │    page.getByRole('button', { name: 'Create Credit Order' })
 * │    page.getByRole('button', { name: 'Create Rebill Order' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Billing Plan' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Output Items' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '–Empty Value' }).first()  // ×4 on page
 * │    page.getByRole('link', { name: '–Empty Value' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | No items available. |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Event ID')
 * │    await expect(page.locator('body')).toContainText('–Empty Value')
 * │    await expect(page.locator('body')).toContainText('Booth Number')
 * │    await expect(page.locator('body')).toContainText('Product & Pricing Zone')
 * │    await expect(page.locator('body')).toContainText('Event Description')
 * │    await expect(page.locator('body')).toContainText('Show Open Date')
 * │    await expect(page.locator('body')).toContainText('Show Close Date')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-in Date')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-Out Date')
 * │    await expect(page.locator('body')).toContainText('Customer Type')
 * │    await expect(page.locator('body')).toContainText('Event Site Facility')
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Document Date')
 * │    await expect(page.locator('body')).toContainText('Source System')
 * │    await expect(page.locator('body')).toContainText('Source System Order ID')
 * │    await expect(page.locator('body')).toContainText('Order Reason')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Address')
 * │    await expect(page.locator('body')).toContainText('Terms of Payment')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Create/Edit Orders
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:17.442Z
 * │  Headings: New: Sales Order  ·  Event Details  ·  Partners  ·  Status  ·  Credit Limit Utilization  ·  Net Sales Volume (YTD)  ·  Net Amount  ·  Amount Details
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '–Empty Value' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Create/Edit Orders' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Withdraw Approval Request' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Generate Payment Link' })
 * │    page.getByRole('button', { name: 'Launch Service Cloud' })
 * │    page.getByRole('button', { name: 'Simulate Taxes' })
 * │    page.getByRole('button', { name: 'Create Credit Order' })
 * │    page.getByRole('button', { name: 'Create Rebill Order' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Billing Plan' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Output Items' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '–Empty Value' }).first()  // ×4 on page
 * │    page.getByRole('link', { name: '–Empty Value' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | No items available. |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Event ID')
 * │    await expect(page.locator('body')).toContainText('–Empty Value')
 * │    await expect(page.locator('body')).toContainText('Bill to Party Name')
 * │    await expect(page.locator('body')).toContainText('Booth Number')
 * │    await expect(page.locator('body')).toContainText('Booth Key')
 * │    await expect(page.locator('body')).toContainText('Priority Empty Zone')
 * │    await expect(page.locator('body')).toContainText('Product & Pricing Zone')
 * │    await expect(page.locator('body')).toContainText('Event Description')
 * │    await expect(page.locator('body')).toContainText('Show Open Date')
 * │    await expect(page.locator('body')).toContainText('Show Close Date')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-in Date')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-Out Date')
 * │    await expect(page.locator('body')).toContainText('Customer Type')
 * │    await expect(page.locator('body')).toContainText('Event Site Facility')
 * │    await expect(page.locator('body')).toContainText('Event Site Facility Description.')
 * │    await expect(page.locator('body')).toContainText('Door/Hall')
 * │    await expect(page.locator('body')).toContainText('Door/Hall Description')
 * │    await expect(page.locator('body')).toContainText('Event Site L3 Description')
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:21.117Z
 * │  Headings: New: Sales Order  ·  Event Details  ·  Partners  ·  Status  ·  Net Sales Volume (YTD)  ·  Net Amount  ·  Amount Details
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '–Empty Value' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Generate Payment Link' })
 * │    page.getByRole('button', { name: 'Launch Service Cloud' })
 * │    page.getByRole('button', { name: 'Simulate Taxes' })
 * │    page.getByRole('button', { name: 'Create Credit Order' })
 * │    page.getByRole('button', { name: 'Create Rebill Order' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '–Empty Value' }).first()  // ×3 on page
 * │    page.getByRole('link', { name: '–Empty Value' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | No items available. |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('No items available.')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:25.098Z
 * │  Headings: New: Sales Order  ·  Event Details  ·  Partners  ·  Status  ·  Net Sales Volume (YTD)  ·  Net Amount  ·  Amount Details
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '–Empty Value' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Generate Payment Link' })
 * │    page.getByRole('button', { name: 'Launch Service Cloud' })
 * │    page.getByRole('button', { name: 'Simulate Taxes' })
 * │    page.getByRole('button', { name: 'Create Credit Order' })
 * │    page.getByRole('button', { name: 'Create Rebill Order' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '–Empty Value' }).first()  // ×3 on page
 * │    page.getByRole('link', { name: '–Empty Value' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | No items available. |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('No items available.')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:29.145Z
 * │  Headings: Select: Sold-to Party  ·  New: Sales Order  ·  Event Details  ·  Partners  ·  Status  ·  Net Sales Volume (YTD)  ·  Net Amount  ·  Amount Details
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '–Empty Value' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search')  // {{search}}
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Show Filters' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Generate Payment Link' })
 * │    page.getByRole('button', { name: 'Launch Service Cloud' })
 * │    page.getByRole('button', { name: 'Simulate Taxes' })
 * │    page.getByRole('button', { name: 'Create Credit Order' })
 * │    page.getByRole('button', { name: 'Create Rebill Order' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '–Empty Value' }).first()  // ×3 on page
 * │    page.getByRole('link', { name: '–Empty Value' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | No items available. |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Customer')
 * │    await expect(page.locator('body')).toContainText('Customer Name')
 * │    await expect(page.locator('body')).toContainText('Customer Type')
 * │    await expect(page.locator('body')).toContainText('Evt Prtnr Eng ID')
 * │    await expect(page.locator('body')).toContainText('Booth Number')
 * │    await expect(page.locator('body')).toContainText('Event ID')
 * │    await expect(page.locator('body')).toContainText('Booth Eng ID')
 * │    await expect(page.locator('body')).toContainText('ZONE Information')
 * │    await expect(page.locator('body')).toContainText('Event Facility Venue')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-In Date')
 * │    await expect(page.locator('body')).toContainText('Prod & Pricing Zone')
 * │    await expect(page.locator('body')).toContainText('Finish Date')
 * │    await expect(page.locator('body')).toContainText('Start Date')
 * │    await expect(page.locator('body')).toContainText('Project Name')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-Out Date')
 * │    await expect(page.locator('body')).toContainText('No items available.')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:33.739Z
 * │  Headings: Select: Sold-to Party  ·  New: Sales Order  ·  Event Details  ·  Partners  ·  Status  ·  Net Sales Volume (YTD)  ·  Net Amount  ·  Amount Details
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '–Empty Value' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search')  // {{search}}
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Show Filters' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Generate Payment Link' })
 * │    page.getByRole('button', { name: 'Launch Service Cloud' })
 * │    page.getByRole('button', { name: 'Simulate Taxes' })
 * │    page.getByRole('button', { name: 'Create Credit Order' })
 * │    page.getByRole('button', { name: 'Create Rebill Order' })
 * │    page.getByRole('button', { name: 'Collapse Header' })
 * │    page.getByRole('button', { name: 'Pin Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '–Empty Value' }).first()  // ×3 on page
 * │    page.getByRole('link', { name: '–Empty Value' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ | No items available. |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Customer')
 * │    await expect(page.locator('body')).toContainText('Customer Name')
 * │    await expect(page.locator('body')).toContainText('Customer Type')
 * │    await expect(page.locator('body')).toContainText('Evt Prtnr Eng ID')
 * │    await expect(page.locator('body')).toContainText('Booth Number')
 * │    await expect(page.locator('body')).toContainText('Event ID')
 * │    await expect(page.locator('body')).toContainText('Booth Eng ID')
 * │    await expect(page.locator('body')).toContainText('ZONE Information')
 * │    await expect(page.locator('body')).toContainText('Event Facility Venue')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-In Date')
 * │    await expect(page.locator('body')).toContainText('Prod & Pricing Zone')
 * │    await expect(page.locator('body')).toContainText('Finish Date')
 * │    await expect(page.locator('body')).toContainText('Start Date')
 * │    await expect(page.locator('body')).toContainText('Project Name')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Exh. Move-Out Date')
 * │    await expect(page.locator('body')).toContainText('1000000030')
 * │    await expect(page.locator('body')).toContainText('Appdetex')
 * │    await expect(page.locator('body')).toContainText('EXH')
 * │    await expect(page.locator('body')).toContainText('EPE100000003066')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:37.181Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:41.189Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:45.262Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:49.503Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:54.320Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Last Name')  // {{ordering_person_contact_last_name}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:44:57.326Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Last Name')  // {{ordering_person_contact_last_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Email')  // {{ordering_person_contact_email}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:45:01.301Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Last Name')  // {{ordering_person_contact_last_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Email')  // {{ordering_person_contact_email}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:45:05.327Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Last Name')  // {{ordering_person_contact_last_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Email')  // {{ordering_person_contact_email}}  [REQUIRED]
 * │    // Section: Bill-To Person Contact
 * │    page.getByLabel('Email')  // {{bill_to_person_contact_email}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:45:09.512Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Last Name')  // {{ordering_person_contact_last_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Email')  // {{ordering_person_contact_email}}  [REQUIRED]
 * │    // Section: Bill-To Person Contact
 * │    page.getByLabel('Email')  // {{bill_to_person_contact_email}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:45:13.704Z
 * │  Headings: New: Sales Order
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    // Section: Event Data
 * │    page.getByLabel('Event ID')  // {{event_data_event_id}}
 * │    // Section: Order Data
 * │    page.getByLabel('Sold-to Party')  // {{order_data_sold_to_party}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('First Name')  // {{ordering_person_contact_first_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Last Name')  // {{ordering_person_contact_last_name}}  [REQUIRED]
 * │    // Section: Ordering Person Contact
 * │    page.getByLabel('Email')  // {{ordering_person_contact_email}}  [REQUIRED]
 * │    // Section: Bill-To Person Contact
 * │    page.getByLabel('Email')  // {{bill_to_person_contact_email}}  [REQUIRED]
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Change Ship-to Party Data' })
 * │    page.getByRole('button', { name: 'Show Details' })
 * │    page.getByRole('button', { name: 'Show More' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Change Address' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: 'Details' }).first()  // ×5 on page
 * │    page.getByRole('link', { name: 'Details' }).nth(0)  // replace 0 with row index
 * │
 * │  ── Tables
 * │    // Column headers: Partner Function | Partner | Address | Doc-Specific Address
 * │    page.getByRole('columnheader', { name: 'Partner Function' })
 * │    page.getByRole('columnheader', { name: 'Partner' })
 * │    page.getByRole('columnheader', { name: 'Address' })
 * │    page.getByRole('columnheader', { name: 'Doc-Specific Address' })
 * │    // Sample row locators (use a unique cell value to scope the row):
 * │    // Row data samples:
 * │      [ |  | Sold-to Party |  |  |  | Details |  | ]
 * │      [ |  | Bill-to Party |  |  |  | Details |  | ]
 * │      [ |  | Payer |  |  |  | Details |  | ]
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Sold-to Party')
 * │    await expect(page.locator('body')).toContainText('Details')
 * │    await expect(page.locator('body')).toContainText('Bill-to Party')
 * │    await expect(page.locator('body')).toContainText('Payer')
 * │    await expect(page.locator('body')).toContainText('Ship-to Party')
 * │    await expect(page.locator('body')).toContainText('Ordering Party')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')
 * │  Time : 2026-04-28T08:45:17.832Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Settings' })
 * │    page.getByRole('button', { name: 'excel-attachment' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:19.380Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:23.417Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:27.405Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:31.419Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'edit' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('4000000003407')
 * │    await expect(page.locator('body')).toContainText('1/2 Dumpster Fee')
 * │    await expect(page.locator('body')).toContainText('Cleaning Services (13000)')
 * │    await expect(page.locator('body')).toContainText('Waste Management (13700)')
 * │    await expect(page.locator('body')).toContainText('End of Show Bulk Trash Removal (00005200)')
 * │    await expect(page.locator('body')).toContainText('2000000005340')
 * │    await expect(page.locator('body')).toContainText('1/2M X 87" Double Sided Sign')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:35.424Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'edit' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('4000000003407')
 * │    await expect(page.locator('body')).toContainText('1/2 Dumpster Fee')
 * │    await expect(page.locator('body')).toContainText('Cleaning Services (13000)')
 * │    await expect(page.locator('body')).toContainText('Waste Management (13700)')
 * │    await expect(page.locator('body')).toContainText('End of Show Bulk Trash Removal (00005200)')
 * │    await expect(page.locator('body')).toContainText('2000000005340')
 * │    await expect(page.locator('body')).toContainText('1/2M X 87" Double Sided Sign')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:39.750Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'edit' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('4000000003407')
 * │    await expect(page.locator('body')).toContainText('1/2 Dumpster Fee')
 * │    await expect(page.locator('body')).toContainText('Cleaning Services (13000)')
 * │    await expect(page.locator('body')).toContainText('Waste Management (13700)')
 * │    await expect(page.locator('body')).toContainText('End of Show Bulk Trash Removal (00005200)')
 * │    await expect(page.locator('body')).toContainText('2000000005340')
 * │    await expect(page.locator('body')).toContainText('1/2M X 87" Double Sided Sign')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:43.493Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'edit' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('4000000003407')
 * │    await expect(page.locator('body')).toContainText('1/2 Dumpster Fee')
 * │    await expect(page.locator('body')).toContainText('Cleaning Services (13000)')
 * │    await expect(page.locator('body')).toContainText('Waste Management (13700)')
 * │    await expect(page.locator('body')).toContainText('End of Show Bulk Trash Removal (00005200)')
 * │    await expect(page.locator('body')).toContainText('2000000005340')
 * │    await expect(page.locator('body')).toContainText('1/2M X 87" Double Sided Sign')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:47.537Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'edit' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('4000000003407')
 * │    await expect(page.locator('body')).toContainText('1/2 Dumpster Fee')
 * │    await expect(page.locator('body')).toContainText('Cleaning Services (13000)')
 * │    await expect(page.locator('body')).toContainText('Waste Management (13700)')
 * │    await expect(page.locator('body')).toContainText('End of Show Bulk Trash Removal (00005200)')
 * │    await expect(page.locator('body')).toContainText('2000000005340')
 * │    await expect(page.locator('body')).toContainText('1/2M X 87" Double Sided Sign')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:51.536Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'edit' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('4000000003407')
 * │    await expect(page.locator('body')).toContainText('1/2 Dumpster Fee')
 * │    await expect(page.locator('body')).toContainText('Cleaning Services (13000)')
 * │    await expect(page.locator('body')).toContainText('Waste Management (13700)')
 * │    await expect(page.locator('body')).toContainText('End of Show Bulk Trash Removal (00005200)')
 * │    await expect(page.locator('body')).toContainText('2000000005340')
 * │    await expect(page.locator('body')).toContainText('1/2M X 87" Double Sided Sign')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:45:55.799Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (00034100)')
 * │    await expect(page.locator('body')).toContainText('3000000002195')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:00.043Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (00034100)')
 * │    await expect(page.locator('body')).toContainText('3000000002195')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:03.581Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Informative entry')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('181.00')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:07.592Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Informative entry')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('181.00')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:11.594Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Informative entry')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('181.00')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:15.623Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Informative entry')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('181.00')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:19.633Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Informative entry')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('181.00')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:23.671Z
 * │  Headings: New: Sales Order  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Inputs / Form Fields
 * │    page.getByLabel('Search by Material')  // {{search_by_material}}
 * │    page.getByLabel('Qty')  // {{qty}}
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Select Options' })
 * │    page.getByRole('button', { name: 'Go' })
 * │    page.getByRole('button', { name: 'Select' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Copy to Clipboard' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Select')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Product Description Extended')
 * │    await expect(page.locator('body')).toContainText('Product Description')
 * │    await expect(page.locator('body')).toContainText('Qty')
 * │    await expect(page.locator('body')).toContainText('Available Onsite Inventory')
 * │    await expect(page.locator('body')).toContainText('UoM')
 * │    await expect(page.locator('body')).toContainText('Current Date Price')
 * │    await expect(page.locator('body')).toContainText('Additional Info')
 * │    await expect(page.locator('body')).toContainText('Material Category')
 * │    await expect(page.locator('body')).toContainText('Material Sub Category')
 * │    await expect(page.locator('body')).toContainText('Material Type')
 * │    await expect(page.locator('body')).toContainText('BOM Usage')
 * │    await expect(page.locator('body')).toContainText('Informative entry')
 * │    await expect(page.locator('body')).toContainText('1000000000374')
 * │    await expect(page.locator('body')).toContainText('10\' X 10\' Carpet Padding - Single Layer')
 * │    await expect(page.locator('body')).toContainText('10\'X10\' Crpt Pddng Single Layer')
 * │    await expect(page.locator('body')).toContainText('181.00')
 * │    await expect(page.locator('body')).toContainText('Flooring (12000)')
 * │    await expect(page.locator('body')).toContainText('Padding & Protective Plastic Covering​ (13100)')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:27.662Z
 * │  Headings: New: Sales Order  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:31.679Z
 * │  Headings: Confirmation  ·  New: Sales Order  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'OK' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:35.861Z
 * │  Headings: New: Sales Order  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Incompleteness Info' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Event Product Catalog' })
 * │    page.getByRole('button', { name: 'Propose Items' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Check Availability' })
 * │    page.getByRole('button', { name: 'Delete' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │    page.getByRole('button', { name: 'Add Row' })
 * │    page.getByRole('button', { name: 'Create' })
 * │    page.getByRole('button', { name: 'Cancel' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:41.948Z
 * │  Headings: 1000376127  ·  Sales Order Items
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'Close' })
 * │    page.getByRole('button', { name: 'Expand Header' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:43.720Z
 * │  Headings: 1000376127  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '10\'X10\' Crpt Pddng Single Layer (1000000000374)' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:47.773Z
 * │  Headings: 1000376127  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '10\'X10\' Crpt Pddng Single Layer (1000000000374)' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:51.843Z
 * │  Headings: 1000376127  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '10\'X10\' Crpt Pddng Single Layer (1000000000374)' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:55.796Z
 * │  Headings: 1000376127  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '10\'X10\' Crpt Pddng Single Layer (1000000000374)' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:46:59.840Z
 * │  Headings: 1000376127  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '10\'X10\' Crpt Pddng Single Layer (1000000000374)' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ┌─ Page : Sales Order
 * │  URL  : https://test.corpsapnext.freeman.com/sap/bc/ui2/flp?sap-client=100&sap-language=EN#ManageSales-display&/SalesOrderManage('1000376127')?sap-iapp-state=TAS7SXNUSPJLDJOM1SCHQ30KCSS5LT7OZCOW6GXVM
 * │  Time : 2026-04-28T08:47:02.269Z
 * │  Headings: 1000376127  ·  Sales Order Items (1)
 * │
 * │  ── Navigation
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │    page.getByRole('link', { name: '' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Sales Order' })
 * │    page.getByRole('button', { name: 'Edit' })
 * │    page.getByRole('button', { name: 'Update Prices' })
 * │    page.getByRole('button', { name: 'Delivery Block' })
 * │    page.getByRole('button', { name: 'Billing Block' })
 * │    page.getByRole('button', { name: 'Display Change Log' })
 * │    page.getByRole('button', { name: 'Attachments (0)' })
 * │    page.getByRole('button', { name: 'Share (Ctrl+Shift+S)' })
 * │    page.getByRole('button', { name: 'General Information' })
 * │    page.getByRole('button', { name: 'Open Menu' })
 * │    page.getByRole('button', { name: 'Items' })
 * │    page.getByRole('button', { name: 'Prices' })
 * │    page.getByRole('button', { name: 'Comments' })
 * │    page.getByRole('button', { name: 'Status and Blocks' })
 * │    page.getByRole('button', { name: 'Process Flow' })
 * │    page.getByRole('button', { name: 'Select View' })
 * │    page.getByRole('button', { name: 'Rejection Reason' })
 * │    page.getByRole('button', { name: 'Additional Options' })
 * │    page.getByRole('button', { name: 'ConfirmedEntry successfully validated' })
 * │    page.getByRole('button', { name: 'Navigation' })
 * │
 * │  ── Links
 * │    page.getByRole('link', { name: '10\'X10\' Crpt Pddng Single Layer (1000000000374)' })
 * │
 * │  ── Visible Data / Text (for assertions)
 * │    await expect(page.locator('body')).toContainText('Item')
 * │    await expect(page.locator('body')).toContainText('Product')
 * │    await expect(page.locator('body')).toContainText('Requested Quantity')
 * │    await expect(page.locator('body')).toContainText('Confirmed Quantity')
 * │    await expect(page.locator('body')).toContainText('Estimated People')
 * │    await expect(page.locator('body')).toContainText('Item Category')
 * │    await expect(page.locator('body')).toContainText('Estimated Hours')
 * │    await expect(page.locator('body')).toContainText('Third Party Material Code')
 * │    await expect(page.locator('body')).toContainText('Higher-Level Item')
 * │    await expect(page.locator('body')).toContainText('Item BOM Type')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Date')
 * │    await expect(page.locator('body')).toContainText('Requested Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Estimated Start Time')
 * │    await expect(page.locator('body')).toContainText('Confirmed Delivery Date')
 * │    await expect(page.locator('body')).toContainText('Availability')
 * │    await expect(page.locator('body')).toContainText('Actual People')
 * │    await expect(page.locator('body')).toContainText('Net Value')
 * │    await expect(page.locator('body')).toContainText('Total Amount')
 * │    await expect(page.locator('body')).toContainText('Actual Hours')
 * │    await expect(page.locator('body')).toContainText('Product Name Long Text')
 * └─────────────────────────────────────────────────────────────────
 * 
 * ── Required field summary (for Excel template column marking) ──
 * REQUIRED_FIELD_LABELS: {{user}}, {{password}}, {{sales_order_type}}, {{order_data_sold_to_party}}, {{ordering_person_contact_first_name}}, {{ordering_person_contact_last_name}}, {{ordering_person_contact_email}}, {{bill_to_person_contact_email}}
 * 
 * ═══════════════════════════════════════════════════════════════════
 */