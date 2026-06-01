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
import LoginPage from '../support/pages/LoginPage';
import ProductsPage from '../support/pages/ProductsPage';
import CartPage from '../support/pages/CartPage';
import CheckoutInformationPage from '../support/pages/CheckoutInformationPage';
import CheckoutOverviewPage from '../support/pages/CheckoutOverviewPage';
import CheckoutCompletePage from '../support/pages/CheckoutCompletePage';
describe('TC_AIGPMM_38_001: Complete Checkout Flow - Single Product', () => {
    let loginPage;
    let productsPage;
    let cartPage;
    let checkoutInfoPage;
    let checkoutOverviewPage;
    let checkoutCompletePage;
    // Test data using {{placeholders}} for data-driven execution
    // These will be replaced with actual values from data source
    let testData;
    before(() => {
        // Initialize Page Object Models
        loginPage = new LoginPage();
        productsPage = new ProductsPage();
        cartPage = new CartPage();
        checkoutInfoPage = new CheckoutInformationPage();
        checkoutOverviewPage = new CheckoutOverviewPage();
        checkoutCompletePage = new CheckoutCompletePage();
        // Load test data from fixture (supports {{placeholders}})
        cy.fixture('testData').then((data) => {
            testData = data;
        });
    });
    beforeEach(() => {
        cy.log('=== Starting Test: TC_AIGPMM_38_001 - Complete Checkout Flow - Single Product ===');
    });
    it('should complete checkout process for single product purchase successfully', () => {
        // Step 1: Navigate to the application URL
        cy.logStep(1, 'Navigate to application URL');
        loginPage.navigateToLoginPage(testData.base_url);
        // Expected Result 1: Login page loads successfully
        loginPage.verifyLoginPageDisplayed();
        // Step 2: Enter username in the Username field
        cy.logStep(2, 'Enter username in the Username field');
        loginPage.enterUsername(testData.username);
        // Expected Result 2: Username field accepts input
        loginPage.verifyUsernameFieldEnabled();
        loginPage.getUsernameValue().should('equal', testData.username);
        cy.log('✓ Verified: Username value matches input: ' + testData.username);
        // Step 3: Enter password in the Password field
        cy.logStep(3, 'Enter password in the Password field');
        loginPage.enterPassword(testData.password);
        // Expected Result 3: Password field accepts input and masks characters
        loginPage.verifyPasswordFieldMasked();
        // Step 4: Click 'Login' button
        cy.logStep(4, 'Click Login button');
        // Expected Result 4: Login button is clickable
        loginPage.verifyLoginButtonClickable();
        loginPage.clickLoginButton();
        // Step 5: Wait for Products page to load
        cy.logStep(5, 'Wait for Products page to load');
        productsPage.waitForProductsPageToLoad();
        // Expected Result 5: Products page loads with product listings
        productsPage.verifyProductsPageDisplayed();
        // Step 6: Click 'Add to cart' button for Sauce Labs Backpack
        cy.logStep(6, "Click 'Add to cart' button for Sauce Labs Backpack");
        productsPage.clickAddToCartForSauceLabsBackpack();
        // Expected Result 6: 'Add to cart' button changes to 'Remove' after clicking
        productsPage.verifyRemoveButtonDisplayed();
        // Step 7: Click on 'Cart' icon in the top right corner
        cy.logStep(7, 'Click on Cart icon');
        // Expected Result 7: Cart icon shows badge with '1'
        productsPage.verifyCartBadgeCount('1');
        productsPage.clickCartIcon();
        // Step 8: Verify product appears in cart with quantity 1
        cy.logStep(8, 'Verify product appears in cart');
        cartPage.waitForCartPageToLoad();
        // Expected Result 8: Cart page displays with correct product name and quantity
        cartPage.verifyProductWithQuantityInCart(testData.product_name);
        // Step 9: Click 'Checkout' button
        cy.logStep(9, 'Click Checkout button');
        // Expected Result 9: Checkout button is visible and clickable
        cartPage.verifyCheckoutButtonVisibleAndClickable();
        cartPage.clickCheckoutButton();
        // Step 10: Wait for 'Checkout: Your Information' page to load
        cy.logStep(10, "Wait for 'Checkout: Your Information' page to load");
        checkoutInfoPage.waitForCheckoutInformationPageToLoad();
        // Expected Result 10: 'Checkout: Your Information' page displays
        checkoutInfoPage.verifyCheckoutInformationPageDisplayed();
        // Step 11: Enter first name in First Name field
        cy.logStep(11, 'Enter first name in First Name field');
        checkoutInfoPage.enterFirstName(testData.first_name);
        // Expected Result 11: First Name field accepts alphabetic input
        checkoutInfoPage.verifyFirstNameFieldEnabled();
        // Step 12: Enter last name in Last Name field
        cy.logStep(12, 'Enter last name in Last Name field');
        checkoutInfoPage.enterLastName(testData.last_name);
        // Expected Result 12: Last Name field accepts alphabetic input
        checkoutInfoPage.verifyLastNameFieldEnabled();
        // Step 13: Enter zip code in Zip/Postal Code field
        cy.logStep(13, 'Enter zip code in Zip/Postal Code field');
        checkoutInfoPage.enterZipPostalCode(testData.zip_code);
        // Expected Result 13: Zip/Postal Code field accepts numeric input
        checkoutInfoPage.verifyZipPostalCodeFieldEnabled();
        // Step 14: Click 'Continue' button
        cy.logStep(14, 'Click Continue button');
        checkoutInfoPage.clickContinueButton();
        // Step 15: Wait for 'Checkout: Overview' page to load
        cy.logStep(15, "Wait for 'Checkout: Overview' page to load");
        checkoutOverviewPage.waitForCheckoutOverviewPageToLoad();
        // Expected Result 15: 'Checkout: Overview' page displays with correct header
        checkoutOverviewPage.verifyCheckoutOverviewHeaderDisplayed();
        // Step 16: Verify product details, payment information, and shipping information
        cy.logStep(16, 'Verify product details, payment information, and shipping information');
        // Expected Result 16: Product table shows quantity and description correctly
        checkoutOverviewPage.verifyProductTableDisplayedCorrectly(testData.product_name);
        // Expected Result 17: Payment Information and Shipping Information sections display
        checkoutOverviewPage.verifyPaymentAndShippingInfoDisplayed();
        // Step 17: Verify Item Total, Tax, and Total amounts
        cy.logStep(17, 'Verify Item Total, Tax, and Total amounts are calculated and displayed');
        // Expected Result 18: Item Total, Tax, and Total are displayed
        checkoutOverviewPage.verifyPriceCalculationsDisplayed();
        checkoutOverviewPage.getItemTotal().should('not.be.empty');
        checkoutOverviewPage.getTax().should('not.be.empty');
        checkoutOverviewPage.getTotal().should('not.be.empty');
        // Step 18: Click 'Finish' button
        cy.logStep(18, 'Click Finish button');
        checkoutOverviewPage.clickFinishButton();
        // Step 19: Wait for Finish page to load
        cy.logStep(19, 'Wait for Finish page to load');
        checkoutCompletePage.waitForFinishPageToLoad();
        // Expected Result 19: Finish page loads successfully
        checkoutCompletePage.verifyFinishPageDisplayed();
        // Step 20: Verify 'Thank you for your order!' message displays
        cy.logStep(20, "Verify 'Thank you for your order!' message displays");
        // Expected Result 20: Success message and Pony Express Sauce Labs logo display
        checkoutCompletePage.verifyThankYouMessageDisplayed();
        checkoutCompletePage.verifyPonyExpressLogoDisplayed();
        cy.log('=== Test Completed Successfully ===');
        cy.log('Complete checkout flow for single product verified successfully');
    });
    afterEach(() => {
        cy.log('=== Completed Test: TC_AIGPMM_38_001 ===');
    });
});