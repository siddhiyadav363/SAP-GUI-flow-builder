// ## PrecisionScript
import { test, expect, Page } from '@playwright/test';
// ## PrecisionScript
test.use({
  ignoreHTTPSErrors: true,
  actionTimeout: 30000,
  httpCredentials: {
    username: '{{username}}',
    password: '{{password}}'
  }
});

test.describe('Create Sales Inquiry', () => {
  test('Create new sales inquiry with material and quantity', async ({ page }) => {
    test.setTimeout(300_000);

    // Step 1: Navigate to Fiori Launchpad
    await page.goto('https://mhana09.mouritech.net:50100/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home');

    // Step 2: Enter username
    await page.getByRole('textbox', { name: 'User' }).click();
    await page.getByRole('textbox', { name: 'User' }).fill('{{username}}');

    // Step 3: Enter password
    await page.getByRole('textbox', { name: 'Password' }).click();
    await page.getByRole('textbox', { name: 'Password' }).fill('{{password}}');

    // Step 4: Click Log On button
    await page.getByRole('button', { name: 'Log On' }).click();

    // Step 5: Navigate to home after login
    await page.goto('https://mhana09.mouritech.net:50100/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home');

    // Step 6: Open Custom Business Catalog - SD Sales Order MNT
    await page.getByLabel('Group Navigation').getByText('Custom Business CatalogGroup - SD Sales Order MNT').click();

    // Step 7: Click Manage Sales Inquiries tile
    await page.getByRole('link', { name: 'Manage Sales Inquiries' }).click();

    // Step 8: Navigate to Manage Sales Inquiries app
    await page.goto('https://mhana09.mouritech.net:50100/sap/bc/ui2/flp?sap-client=100&sap-language=EN#SalesInquiry-manage&/?sap-iapp-state--history=TAS9AJ2K84SPFLBGNXO5PRJKIZNOWG8WU2WXW09KW&sap-iapp-state=TASND3AJ25UL33B0VKOU8KQ5ZXAWRHY0JPZAQSGJG');

    // Step 9: Click Create Inquiry button
    await page.getByRole('button', { name: 'Create Inquiry' }).click();

    // Step 10: Click value help for Sales Order Type
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().locator('#ls-inputfieldhelpbutton').click();

    // Step 11: Confirm Sales Order Type selection
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'OK  Emphasized' }).click();

    // Step 12: Click Sales Organization field
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { name: 'Sales Organization' }).click();

    // Step 13: Click value help for Sales Organization
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().locator('#ls-inputfieldhelpbutton').click();

    // Step 14: Confirm Sales Organization selection
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'OK  Emphasized' }).click();

    // Step 15: Click Distribution Channel field
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { name: 'Distribution Channel' }).click();

    // Step 16: Click value help for Distribution Channel
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().locator('#ls-inputfieldhelpbutton').click();

    // Step 17: Confirm Distribution Channel selection
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'OK  Emphasized' }).click();

    // Step 18: Click Division field
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { name: 'Division' }).click();

    // Step 19: Click value help for Division
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().locator('#ls-inputfieldhelpbutton').click();

    // Step 20: Confirm Division selection
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'OK  Emphasized' }).click();

    // Step 21: Click Continue to proceed to main inquiry screen
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'Continue  Emphasized' }).click();

    // Step 22: Click Sold-to Party field
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { name: 'Sold-to Party' }).click();

    // Step 23: Enter Sold-to Party
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { name: 'Sold-to Party' }).fill('{{sold_to_party}}');

    // Step 24: Press Enter to validate Sold-to Party
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { name: 'Sold-to Party' }).press('Enter');

    // Step 25: Click Continue after Sold-to Party validation
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'Continue' }).click();

    // Step 26: Click Material field
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { description: 'Value help available', exact: true }).nth(3).click();

    // Step 27: Enter Material number
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('textbox', { description: 'Value help available', exact: true }).nth(3).fill('{{material}}');

    // Step 28: Click Quantity field
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().locator('[id="M0:46:2:3B257:1:2[1,4]_c"]').click();

    // Step 29: Enter Quantity
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().locator('[id="M0:46:2:3B257:1:2[1,4]_c"]').fill('{{quantity}}');

    // Step 30: Click Save to create the inquiry
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('button', { name: 'Save  Emphasized' }).click();

    // Step 31: Click success message bar
    await page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('note', { name: 'Success Message Bar Inquiry' }).click();

    // Step 32: Extract Inquiry number from success message
    const successMsgLoc = page.locator('iframe[name="application-SalesInquiry-create-iframe"]').contentFrame().getByRole('note', { name: 'Success Message Bar Inquiry' });
    await successMsgLoc.waitFor({ state: 'visible' });
    const inquiryResult = await successMsgLoc.innerText();
    console.log('Output: ' + String(inquiryResult).trim());
  });
});

/*
 * ═══════════════════════════════════════════════════════════════════
 * ELEMENT CONTEXT  (clicked/focused elements only — stable locators for AI)
 * One entry per unique element you interacted with during recording.
 * ═══════════════════════════════════════════════════════════════════
 * 
 * ┌─ Page : Interacted elements
 * │  URL  : https://mhana09.mouritech.net:50100/sap/bc/ui2/flp?sap-client=100&sap-language=EN#Shell-home
 * │  Time : 2026-05-27T07:40:56.845Z
 * │
 * │  ── Inputs / Form Fields
 * │    // input id is dynamic — use stable anchor below (not #__clone… id)
 * │    // ARIA role+name (form field, not SAP table cell)
 * │    page.getByRole('textbox', { name: 'User' })  // {{user}}  [REQUIRED]
 * │    // input id is dynamic — use stable anchor below (not #__clone… id)
 * │    // ARIA role+name (form field, not SAP table cell)
 * │    page.getByRole('textbox', { name: 'Password' })  // {{password}}  [REQUIRED]
 * │
 * │  ── Interacted element profiles (full attributes, ancestors, .nth() resolution)
 * │    // ── Interacted input label="User"  // {{user}}
 * │    // dynamic id at record time: USERNAME_FIELD-inner — do NOT use in final script
 * │    // target attributes:
 * │      aria-describedby="USERNAME_LABEL"
 * │      aria-required="true"
 * │      autocapitalize="none"
 * │      autocorrect="off"
 * │      class="loginInputField"
 * │      id="USERNAME_FIELD-inner"
 * │      inputmode="verbatim"
 * │      maxlength="12 "
 * │      name="sap-user"
 * │      placeholder="User"
 * │      required="true"
 * │      tabindex="0"
 * │      title="User"
 * │      type="text"
 * │    // ancestor chain (nearest parent first):
 * │      <div>
 * │        class="loginInput sapUiLightestBG"
 * │        id="USERNAME_BLOCK"
 * │      <form>
 * │        accept-charset="UTF-8"
 * │        action="https://mhana09.mouritech.net:50100/sap/bc/ui2/flp?_sap-hash=JTIzU2hlbGwtaG9tZQ"
 * │        autocomplete="off"
 * │        class="loginForm"
 * │        id="LOGIN_FORM"
 * │        method="post"
 * │        name="loginForm"
 * │    // suggested Playwright locators (prefer first stable line):
 * │    page.getByRole('textbox', { name: 'User' })
 * │    // ── Interacted input label="Password"  // {{password}}
 * │    // dynamic id at record time: PASSWORD_FIELD-inner — do NOT use in final script
 * │    // target attributes:
 * │      aria-describedby="PASSWORD_LABEL"
 * │      aria-required="true"
 * │      class="loginInputField"
 * │      id="PASSWORD_FIELD-inner"
 * │      inputmode="verbatim"
 * │      name="sap-password"
 * │      placeholder="Password"
 * │      required="true"
 * │      tabindex="0"
 * │      title="Password"
 * │      type="password"
 * │    // ancestor chain (nearest parent first):
 * │      <div>
 * │        class="loginInput sapUiLightestBG"
 * │        id="PASSWORD_BLOCK"
 * │      <form>
 * │        accept-charset="UTF-8"
 * │        action="https://mhana09.mouritech.net:50100/sap/bc/ui2/flp?_sap-hash=JTIzU2hlbGwtaG9tZQ"
 * │        autocomplete="off"
 * │        class="loginForm"
 * │        id="LOGIN_FORM"
 * │        method="post"
 * │        name="loginForm"
 * │    // suggested Playwright locators (prefer first stable line):
 * │    page.getByRole('textbox', { name: 'Password' })
 * │
 * │  ── Buttons
 * │    page.getByRole('button', { name: 'Log On' })
 * └─────────────────────────────────────────────────────────────────
 * 
 * ── Required field summary (for Excel template column marking) ──
 * REQUIRED_FIELD_LABELS: {{user}}, {{password}}
 * 
 * ═══════════════════════════════════════════════════════════════════
 */