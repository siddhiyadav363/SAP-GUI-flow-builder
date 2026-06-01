/**
 * Test Case: TC_AIGPMM_38_001
 * Test complete checkout flow for a single product from cart to finish page
 * Story: AIGPMM-38
 * Priority: High
 * 
 * Objective: Validate that a user can complete the entire checkout process 
 * with valid information for a single product purchase
 * Test Type: Functional
 */
import { Builder, WebDriver } from 'selenium-webdriver';
import { LoginPage } from '../pages/LoginPage';
import { ProductsPage } from '../pages/ProductsPage';
import { CartPage } from '../pages/CartPage';
import { CheckoutInformationPage } from '../pages/CheckoutInformationPage';
import { CheckoutOverviewPage } from '../pages/CheckoutOverviewPage';
import { CheckoutCompletePage } from '../pages/CheckoutCompletePage';
import { TestData } from '../utils/TestData';
import * as assert from 'assert';
describe('TC_AIGPMM_38_001: Complete Checkout Flow - Single Product', () => {
    let driver: WebDriver;
    let loginPage: LoginPage;
    let productsPage: ProductsPage;
    let cartPage: CartPage;
    let checkoutInfoPage: CheckoutInformationPage;
    let checkoutOverviewPage: CheckoutOverviewPage;
    let checkoutCompletePage: CheckoutCompletePage;
    let testData: TestData;
    before(async function() {
        this.timeout(60000);
        console.log('\n=== Starting Test: TC_AIGPMM_38_001 - Complete Checkout Flow - Single Product ===\n');
        // Initialize WebDriver
        driver = await new Builder().forBrowser('chrome').build();
        await driver.manage().window().maximize();
        await driver.manage().setTimeouts({ implicit: 10000 });
        // Initialize Page Object Models
        loginPage = new LoginPage(driver);
        productsPage = new ProductsPage(driver);
        cartPage = new CartPage(driver);
        checkoutInfoPage = new CheckoutInformationPage(driver);
        checkoutOverviewPage = new CheckoutOverviewPage(driver);
        checkoutCompletePage = new CheckoutCompletePage(driver);
        // Load test data with {{placeholders}}
        testData = new TestData();
        console.log('Browser initialized: Chrome');
    });
    after(async function() {
        this.timeout(30000);
        if (driver) {
            await driver.quit();
            console.log('\n=== Test Completed Successfully ===');
            console.log('Complete checkout flow for single product verified successfully\n');
        }
    });
    it('should complete checkout process for single product purchase successfully', async function() {
        this.timeout(120000);
        // Step 1: Navigate to the application URL
        loginPage.logStep(1, 'Navigate to application URL');
        await loginPage.navigateToLoginPage(testData.baseUrl);
        // Expected Result 1: Login page loads successfully
        await loginPage.verifyLoginPageDisplayed();
        // Step 2: Enter username in the Username field
        loginPage.logStep(2, 'Enter username in the Username field');
        await loginPage.enterUsername(testData.username);
        // Expected Result 2: Username field accepts input
        await loginPage.verifyUsernameFieldEnabled();
        const usernameValue = await loginPage.getUsernameValue();
        assert.strictEqual(usernameValue, testData.username, `Username value should match input: ${testData.username}`);
        loginPage.logVerification(`Username value matches input: ${usernameValue}`);
        // Step 3: Enter password in the Password field
        loginPage.logStep(3, 'Enter password in the Password field');
        await loginPage.enterPassword(testData.password);
        // Expected Result 3: Password field accepts input and masks characters
        await loginPage.verifyPasswordFieldMasked();
        // Step 4: Click 'Login' button
        loginPage.logStep(4, 'Click Login button');
        // Expected Result 4: Login button is clickable
        await loginPage.verifyLoginButtonClickable();
        await loginPage.clickLoginButton();
        // Step 5: Wait for Products page to load
        productsPage.logStep(5, 'Wait for Products page to load');
        await productsPage.waitForProductsPageToLoad();
        // Expected Result 5: Products page loads with product listings
        await productsPage.verifyProductsPageDisplayed();
        // Step 6: Click 'Add to cart' button for Sauce Labs Backpack
        productsPage.logStep(6, "Click 'Add to cart' button for Sauce Labs Backpack");
        await productsPage.clickAddToCartForSauceLabsBackpack();
        // Expected Result 6: 'Add to cart' button changes to 'Remove' after clicking
        await productsPage.verifyRemoveButtonDisplayed();
        // Step 7: Click on 'Cart' icon in the top right corner
        productsPage.logStep(7, 'Click on Cart icon');
        // Expected Result 7: Cart icon shows badge with '1'
        await productsPage.verifyCartBadgeCount('1');
        await productsPage.clickCartIcon();
        // Step 8: Verify product appears in cart with quantity 1
        cartPage.logStep(8, 'Verify product appears in cart');
        await cartPage.waitForCartPageToLoad();
        // Expected Result 8: Cart page displays with correct product name and quantity
        await cartPage.verifyProductWithQuantityInCart(testData.productName);
        // Step 9: Click 'Checkout' button
        cartPage.logStep(9, 'Click Checkout button');
        // Expected Result 9: Checkout button is visible and clickable
        await cartPage.verifyCheckoutButtonVisibleAndClickable();
        await cartPage.clickCheckoutButton();
        // Step 10: Wait for 'Checkout: Your Information' page to load
        checkoutInfoPage.logStep(10, "Wait for 'Checkout: Your Information' page to load");
        await checkoutInfoPage.waitForCheckoutInformationPageToLoad();
        // Expected Result 10: 'Checkout: Your Information' page displays
        await checkoutInfoPage.verifyCheckoutInformationPageDisplayed();
        // Step 11: Enter first name in First Name field
        checkoutInfoPage.logStep(11, 'Enter first name in First Name field');
        await checkoutInfoPage.enterFirstName(testData.firstName);
        // Expected Result 11: First Name field accepts alphabetic input
        await checkoutInfoPage.verifyFirstNameFieldEnabled();
        // Step 12: Enter last name in Last Name field
        checkoutInfoPage.logStep(12, 'Enter last name in Last Name field');
        await checkoutInfoPage.enterLastName(testData.lastName);
        // Expected Result 12: Last Name field accepts alphabetic input
        await checkoutInfoPage.verifyLastNameFieldEnabled();
        // Step 13: Enter zip code in Zip/Postal Code field
        checkoutInfoPage.logStep(13, 'Enter zip code in Zip/Postal Code field');
        await checkoutInfoPage.enterZipPostalCode(testData.zipCode);
        // Expected Result 13: Zip/Postal Code field accepts numeric input
        await checkoutInfoPage.verifyZipPostalCodeFieldEnabled();
        // Step 14: Click 'Continue' button
        checkoutInfoPage.logStep(14, 'Click Continue button');
        await checkoutInfoPage.clickContinueButton();
        // Step 15: Wait for 'Checkout: Overview' page to load
        checkoutOverviewPage.logStep(15, "Wait for 'Checkout: Overview' page to load");
        await checkoutOverviewPage.waitForCheckoutOverviewPageToLoad();
        // Expected Result 15: 'Checkout: Overview' page displays with correct header
        await checkoutOverviewPage.verifyCheckoutOverviewHeaderDisplayed();
        // Step 16: Verify product details, payment information, and shipping information
        checkoutOverviewPage.logStep(16, 'Verify product details, payment information, and shipping information');
        // Expected Result 16: Product table shows quantity and description correctly
        await checkoutOverviewPage.verifyProductTableDisplayedCorrectly(testData.productName);
        // Expected Result 17: Payment Information and Shipping Information sections display
        await checkoutOverviewPage.verifyPaymentAndShippingInfoDisplayed();
        // Step 17: Verify Item Total, Tax, and Total amounts
        checkoutOverviewPage.logStep(17, 'Verify Item Total, Tax, and Total amounts are calculated and displayed');
        // Expected Result 18: Item Total, Tax, and Total are displayed
        await checkoutOverviewPage.verifyPriceCalculationsDisplayed();
        const itemTotal = await checkoutOverviewPage.getItemTotal();
        const tax = await checkoutOverviewPage.getTax();
        const total = await checkoutOverviewPage.getTotal();
        assert.ok(itemTotal.length > 0, 'Item Total should not be empty');
        assert.ok(tax.length > 0, 'Tax should not be empty');
        assert.ok(total.length > 0, 'Total should not be empty');
        console.log(`  - Item Total: ${itemTotal}`);
        console.log(`  - Tax: ${tax}`);
        console.log(`  - Total: ${total}`);
        // Step 18: Click 'Finish' button
        checkoutOverviewPage.logStep(18, 'Click Finish button');
        await checkoutOverviewPage.clickFinishButton();
        // Step 19: Wait for Finish page to load
        checkoutCompletePage.logStep(19, 'Wait for Finish page to load');
        await checkoutCompletePage.waitForFinishPageToLoad();
        // Expected Result 19: Finish page loads successfully
        await checkoutCompletePage.verifyFinishPageDisplayed();
        // Step 20: Verify 'Thank you for your order!' message displays
        checkoutCompletePage.logStep(20, "Verify 'Thank you for your order!' message displays");
        // Expected Result 20: Success message and Pony Express Sauce Labs logo display
        await checkoutCompletePage.verifyThankYouMessageDisplayed();
        await checkoutCompletePage.verifyPonyExpressLogoDisplayed();
    });
});