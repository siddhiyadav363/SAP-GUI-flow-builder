package com.swaglab.tests;
import com.swaglab.base.TestBase;
import com.swaglab.pages.*;
import org.testng.Assert;
import org.testng.annotations.Test;
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
public class TC_AIGPMM_38_001_Complete_Checkout_Flow_Single_Product extends TestBase {
    private LoginPage loginPage;
    private ProductsPage productsPage;
    private CartPage cartPage;
    private CheckoutInformationPage checkoutInfoPage;
    private CheckoutOverviewPage checkoutOverviewPage;
    private CheckoutCompletePage checkoutCompletePage;
    @Test(description = "Verify successful checkout completion for a single product from cart to finish page")
    public void testCompleteCheckoutFlowSingleProduct() {
        System.out.println("\n=== Starting Test: TC_AIGPMM_38_001 - Complete Checkout Flow - Single Product ===\n");
        // Initialize Page Object Models
        loginPage = new LoginPage(page);
        productsPage = new ProductsPage(page);
        cartPage = new CartPage(page);
        checkoutInfoPage = new CheckoutInformationPage(page);
        checkoutOverviewPage = new CheckoutOverviewPage(page);
        checkoutCompletePage = new CheckoutCompletePage(page);
        // Step 1: Navigate to the application URL
        logStep(1, "Navigate to application URL");
        loginPage.navigateToLoginPage(baseUrl);
        // Expected Result 1: Login page loads successfully
        Assert.assertTrue(loginPage.isLoginPageDisplayed(), "Login page should be displayed");
        logVerification("Login page loads successfully");
        // Step 2: Enter username in the Username field
        logStep(2, "Enter username in the Username field");
        loginPage.enterUsername(username);
        // Expected Result 2: Username field accepts input
        Assert.assertTrue(loginPage.isUsernameFieldEnabled(), "Username field should be enabled");
        String usernameValue = loginPage.getUsernameValue();
        Assert.assertEquals(usernameValue, username, "Username value should match input");
        logVerification("Username field accepts input");
        // Step 3: Enter password in the Password field
        logStep(3, "Enter password in the Password field");
        loginPage.enterPassword(password);
        // Expected Result 3: Password field accepts input and masks characters
        Assert.assertTrue(loginPage.isPasswordFieldMasked(), "Password field should be masked");
        logVerification("Password field accepts input and masks characters");
        // Step 4: Click 'Login' button
        logStep(4, "Click Login button");
        // Expected Result 4: Login button is clickable
        Assert.assertTrue(loginPage.isLoginButtonClickable(), "Login button should be clickable");
        logVerification("Login button is clickable");
        loginPage.clickLoginButton();
        // Step 5: Wait for Products page to load
        logStep(5, "Wait for Products page to load");
        productsPage.waitForProductsPageToLoad();
        // Expected Result 5: Products page loads with product listings
        Assert.assertTrue(productsPage.isProductsPageDisplayed(), "Products page should be displayed");
        logVerification("Products page loads with product listings");
        // Step 6: Click 'Add to cart' button for Sauce Labs Backpack
        logStep(6, "Click 'Add to cart' button for Sauce Labs Backpack");
        productsPage.clickAddToCartForSauceLabsBackpack();
        // Expected Result 6: 'Add to cart' button changes to 'Remove' after clicking
        Assert.assertTrue(productsPage.isRemoveButtonDisplayed(), "'Remove' button should be displayed");
        logVerification("'Add to cart' button changes to 'Remove' after clicking");
        // Step 7: Click on 'Cart' icon in the top right corner
        logStep(7, "Click on Cart icon");
        // Expected Result 7: Cart icon shows badge with '1'
        Assert.assertTrue(productsPage.isCartBadgeCountCorrect("1"), "Cart badge should show '1'");
        logVerification("Cart icon shows badge with '1'");
        productsPage.clickCartIcon();
        // Step 8: Verify product appears in cart with quantity 1
        logStep(8, "Verify product appears in cart");
        cartPage.waitForCartPageToLoad();
        // Expected Result 8: Cart page displays with correct product name and quantity
        Assert.assertTrue(cartPage.verifyProductWithQuantityInCart(productName), 
            "Product '" + productName + "' with quantity 1 should be in cart");
        logVerification("Cart page displays with correct product name '" + productName + "' and quantity 1");
        // Step 9: Click 'Checkout' button
        logStep(9, "Click Checkout button");
        // Expected Result 9: Checkout button is visible and clickable
        Assert.assertTrue(cartPage.isCheckoutButtonVisibleAndClickable(), 
            "Checkout button should be visible and clickable");
        logVerification("Checkout button is visible and clickable");
        cartPage.clickCheckoutButton();
        // Step 10: Wait for 'Checkout: Your Information' page to load
        logStep(10, "Wait for 'Checkout: Your Information' page to load");
        checkoutInfoPage.waitForCheckoutInformationPageToLoad();
        // Expected Result 10: 'Checkout: Your Information' page displays
        Assert.assertTrue(checkoutInfoPage.isCheckoutInformationPageDisplayed(), 
            "'Checkout: Your Information' page should be displayed");
        logVerification("'Checkout: Your Information' page displays with header and three mandatory fields");
        // Step 11: Enter first name in First Name field
        logStep(11, "Enter first name in First Name field");
        checkoutInfoPage.enterFirstName(firstName);
        // Expected Result 11: First Name field accepts alphabetic input
        Assert.assertTrue(checkoutInfoPage.isFirstNameFieldEnabled(), "First Name field should be enabled");
        logVerification("First Name field accepts alphabetic input");
        // Step 12: Enter last name in Last Name field
        logStep(12, "Enter last name in Last Name field");
        checkoutInfoPage.enterLastName(lastName);
        // Expected Result 12: Last Name field accepts alphabetic input
        Assert.assertTrue(checkoutInfoPage.isLastNameFieldEnabled(), "Last Name field should be enabled");
        logVerification("Last Name field accepts alphabetic input");
        // Step 13: Enter zip code in Zip/Postal Code field
        logStep(13, "Enter zip code in Zip/Postal Code field");
        checkoutInfoPage.enterZipPostalCode(zipCode);
        // Expected Result 13: Zip/Postal Code field accepts numeric input
        Assert.assertTrue(checkoutInfoPage.isZipPostalCodeFieldEnabled(), 
            "Zip/Postal Code field should be enabled");
        logVerification("Zip/Postal Code field accepts numeric input");
        // Step 14: Click 'Continue' button
        logStep(14, "Click Continue button");
        checkoutInfoPage.clickContinueButton();
        // Step 15: Wait for 'Checkout: Overview' page to load
        logStep(15, "Wait for 'Checkout: Overview' page to load");
        checkoutOverviewPage.waitForCheckoutOverviewPageToLoad();
        // Expected Result 15: 'Checkout: Overview' page displays with correct header
        Assert.assertTrue(checkoutOverviewPage.isCheckoutOverviewHeaderDisplayed(), 
            "'Checkout: Overview' page header should be displayed");
        logVerification("'Checkout: Overview' page displays with correct header (hamburger menu, SWAGLABS logo, cart icon)");
        // Step 16: Verify product details, payment information, and shipping information
        logStep(16, "Verify product details, payment information, and shipping information");
        // Expected Result 16: Product table shows quantity and description correctly
        Assert.assertTrue(checkoutOverviewPage.isProductTableDisplayedCorrectly(productName), 
            "Product table should display correctly");
        logVerification("Product table shows quantity and description correctly");
        // Expected Result 17: Payment Information and Shipping Information sections display
        Assert.assertTrue(checkoutOverviewPage.arePaymentAndShippingInfoDisplayed(), 
            "Payment and Shipping info should be displayed");
        logVerification("Payment Information and Shipping Information sections display below product list");
        // Step 17: Verify Item Total, Tax, and Total amounts
        logStep(17, "Verify Item Total, Tax, and Total amounts are calculated and displayed");
        // Expected Result 18: Item Total, Tax, and Total are displayed
        Assert.assertTrue(checkoutOverviewPage.arePriceCalculationsDisplayed(), 
            "Price calculations should be displayed");
        String itemTotal = checkoutOverviewPage.getItemTotal();
        String tax = checkoutOverviewPage.getTax();
        String total = checkoutOverviewPage.getTotal();
        Assert.assertFalse(itemTotal.isEmpty(), "Item Total should not be empty");
        Assert.assertFalse(tax.isEmpty(), "Tax should not be empty");
        Assert.assertFalse(total.isEmpty(), "Total should not be empty");
        logVerification("Item Total, Tax, and Total are displayed with correct calculations");
        System.out.println("  - Item Total: " + itemTotal);
        System.out.println("  - Tax: " + tax);
        System.out.println("  - Total: " + total);
        // Step 18: Click 'Finish' button
        logStep(18, "Click Finish button");
        checkoutOverviewPage.clickFinishButton();
        // Step 19: Wait for Finish page to load
        logStep(19, "Wait for Finish page to load");
        checkoutCompletePage.waitForFinishPageToLoad();
        // Expected Result 19: Finish page loads successfully
        Assert.assertTrue(checkoutCompletePage.isFinishPageDisplayed(), "Finish page should be displayed");
        logVerification("Finish page loads successfully");
        // Step 20: Verify 'Thank you for your order!' message displays
        logStep(20, "Verify 'Thank you for your order!' message displays");
        // Expected Result 20: Success message and Pony Express Sauce Labs logo display
        Assert.assertTrue(checkoutCompletePage.isThankYouMessageDisplayed(), 
            "'Thank you for your order!' message should be displayed");
        logVerification("'Thank you for your order!' message displays");
        Assert.assertTrue(checkoutCompletePage.isPonyExpressLogoDisplayed(), 
            "Pony Express logo should be displayed");
        logVerification("Pony Express Sauce Labs logo displays");
        System.out.println("\n=== Test Completed Successfully ===");
        System.out.println("Complete checkout flow for single product verified successfully\n");
    }
}