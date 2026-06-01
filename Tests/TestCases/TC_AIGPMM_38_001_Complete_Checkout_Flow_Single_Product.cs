/**
 * Test Case: TC_AIGPMM_38_001
 * Test complete checkout flow for a single product from cart to finish page
 */
using NUnit.Framework;
using SwagLabsAutomation.POM;
using SwagLabsAutomation.Tests.TestBase;
namespace SwagLabsAutomation.Tests.TestCases
{
    /// <summary>
    /// Test fixture for TC_AIGPMM_38_001: Complete Checkout Flow - Single Product
    /// Verifies successful checkout completion for a single product from cart to finish page
    /// </summary>
    [TestFixture]
    [Description("Verify successful checkout completion for a single product from cart to finish page")]
    [Category("Functional")]
    [Category("Checkout")]
    public class TC_AIGPMM_38_001_Complete_Checkout_Flow_Single_Product : TestBase
    {
        private LoginPagePOM loginPage;
        private ProductsPagePOM productsPage;
        private CartPagePOM cartPage;
        private CheckoutInformationPagePOM checkoutInfoPage;
        private CheckoutOverviewPagePOM checkoutOverviewPage;
        private CheckoutCompletePagePOM checkoutCompletePage;
        /// <summary>
        /// Test setup - Initialize Page Object Models
        /// </summary>
        [SetUp]
        public new void SetUp()
        {
            base.SetUp();
            loginPage = new LoginPagePOM(Driver);
            productsPage = new ProductsPagePOM(Driver);
            cartPage = new CartPagePOM(Driver);
            checkoutInfoPage = new CheckoutInformationPagePOM(Driver);
            checkoutOverviewPage = new CheckoutOverviewPagePOM(Driver);
            checkoutCompletePage = new CheckoutCompletePagePOM(Driver);
        }
        /// <summary>
        /// Test Method: Verify successful checkout completion for a single product
        /// Objective: Validate that a user can complete the entire checkout process with valid information for a single product purchase
        /// Priority: High
        /// </summary>
        [Test]
        [Description("Validate that a user can complete the entire checkout process with valid information for a single product purchase")]
        [Property("Priority", "High")]
        [Property("TestCaseId", "TC_AIGPMM_38_001")]
        [Property("Story", "AIGPMM-38")]
        public void TestCompleteCheckoutFlowSingleProduct()
        {
            // Test Data - Using {{placeholders}} for data-driven execution
            string baseUrl = "{{base_url}}";
            string username = "{{username}}";
            string password = "{{password}}";
            string firstName = "{{first_name}}";
            string lastName = "{{last_name}}";
            string zipCode = "{{zip_code}}";
            string productName = "{{product_name}}";
            // Step 1: Navigate to https://www.saucedemo.com/
            TestContext.WriteLine("Step 1: Navigate to application URL");
            loginPage.NavigateToLoginPage(baseUrl);
            // Expected Result 1: Login page loads successfully
            Assert.That(loginPage.IsLoginPageDisplayed(), Is.True, 
                "Expected: Login page loads successfully");
            // Step 2: Enter 'standard_user' in the Username field
            TestContext.WriteLine("Step 2: Enter username in the Username field");
            loginPage.EnterUsername(username);
            // Expected Result 2: Username field accepts input
            Assert.That(loginPage.IsUsernameFieldAcceptsInput(), Is.True, 
                "Expected: Username field accepts input");
            // Step 3: Enter 'secret_sauce' in the Password field
            TestContext.WriteLine("Step 3: Enter password in the Password field");
            loginPage.EnterPassword(password);
            // Expected Result 3: Password field accepts input and masks characters
            Assert.That(loginPage.IsPasswordFieldMasked(), Is.True, 
                "Expected: Password field accepts input and masks characters");
            // Step 4: Click 'Login' button
            TestContext.WriteLine("Step 4: Click Login button");
            // Expected Result 4: Login button is clickable
            Assert.That(loginPage.IsLoginButtonClickable(), Is.True, 
                "Expected: Login button is clickable");
            loginPage.ClickLoginButton();
            // Step 5: Wait for Products page to load
            TestContext.WriteLine("Step 5: Wait for Products page to load");
            productsPage.WaitForProductsPageToLoad();
            // Expected Result 5: Products page loads with product listings
            Assert.That(productsPage.IsProductsPageDisplayed(), Is.True, 
                "Expected: Products page loads with product listings");
            // Step 6: Click 'Add to cart' button for 'Sauce Labs Backpack'
            TestContext.WriteLine("Step 6: Click 'Add to cart' button for Sauce Labs Backpack");
            productsPage.ClickAddToCartForSauceLabsBackpack();
            // Expected Result 6: 'Add to cart' button changes to 'Remove' after clicking
            Assert.That(productsPage.IsRemoveButtonDisplayedForSauceLabsBackpack(), Is.True, 
                "Expected: 'Add to cart' button changes to 'Remove' after clicking");
            // Step 7: Click on 'Cart' icon in the top right corner
            TestContext.WriteLine("Step 7: Click on Cart icon");
            // Expected Result 7: Cart icon shows badge with '1'
            Assert.That(productsPage.IsCartBadgeCountCorrect("1"), Is.True, 
                "Expected: Cart icon shows badge with '1'");
            productsPage.ClickCartIcon();
            // Step 8: Verify 'Sauce Labs Backpack' appears in cart with quantity 1
            TestContext.WriteLine("Step 8: Verify product appears in cart");
            cartPage.WaitForCartPageToLoad();
            // Expected Result 8: Cart page displays with correct product name and quantity
            Assert.That(cartPage.VerifyProductWithQuantityInCart(productName), Is.True, 
                "Expected: Cart page displays with correct product name and quantity");
            // Step 9: Click 'Checkout' button
            TestContext.WriteLine("Step 9: Click Checkout button");
            // Expected Result 9: Checkout button is visible and clickable
            Assert.That(cartPage.IsCheckoutButtonVisibleAndClickable(), Is.True, 
                "Expected: Checkout button is visible and clickable");
            cartPage.ClickCheckoutButton();
            // Step 10: Wait for 'Checkout: Your Information' page to load
            TestContext.WriteLine("Step 10: Wait for Checkout: Your Information page to load");
            checkoutInfoPage.WaitForCheckoutInformationPageToLoad();
            // Expected Result 10: 'Checkout: Your Information' page displays with header and three mandatory fields
            Assert.That(checkoutInfoPage.IsCheckoutInformationPageDisplayed(), Is.True, 
                "Expected: 'Checkout: Your Information' page displays with header and three mandatory fields");
            // Step 11: Enter 'Sarah' in First Name field
            TestContext.WriteLine("Step 11: Enter first name in First Name field");
            checkoutInfoPage.EnterFirstName(firstName);
            // Expected Result 11: First Name field accepts alphabetic input
            Assert.That(checkoutInfoPage.IsFirstNameFieldAcceptsInput(), Is.True, 
                "Expected: First Name field accepts alphabetic input");
            // Step 12: Enter 'Johnson' in Last Name field
            TestContext.WriteLine("Step 12: Enter last name in Last Name field");
            checkoutInfoPage.EnterLastName(lastName);
            // Expected Result 12: Last Name field accepts alphabetic input
            Assert.That(checkoutInfoPage.IsLastNameFieldAcceptsInput(), Is.True, 
                "Expected: Last Name field accepts alphabetic input");
            // Step 13: Enter '78701' in Zip/Postal Code field
            TestContext.WriteLine("Step 13: Enter zip code in Zip/Postal Code field");
            checkoutInfoPage.EnterZipPostalCode(zipCode);
            // Expected Result 13: Zip/Postal Code field accepts numeric input
            Assert.That(checkoutInfoPage.IsZipPostalCodeFieldAcceptsInput(), Is.True, 
                "Expected: Zip/Postal Code field accepts numeric input");
            // Step 14: Click 'Continue' button
            TestContext.WriteLine("Step 14: Click Continue button");
            checkoutInfoPage.ClickContinueButton();
            // Expected Result 14: Continue button navigates to next page
            // (Navigation verified by waiting for overview page in next step)
            // Step 15: Wait for 'Checkout: Overview' page to load
            TestContext.WriteLine("Step 15: Wait for Checkout: Overview page to load");
            checkoutOverviewPage.WaitForCheckoutOverviewPageToLoad();
            // Expected Result 15: 'Checkout: Overview' page displays with correct header (hamburger menu, SWAGLABS logo, cart icon)
            Assert.That(checkoutOverviewPage.IsCheckoutOverviewHeaderDisplayed(), Is.True, 
                "Expected: 'Checkout: Overview' page displays with correct header (hamburger menu, SWAGLABS logo, cart icon)");
            // Step 16: Verify product details, payment information, and shipping information display correctly
            TestContext.WriteLine("Step 16: Verify product details, payment information, and shipping information");
            // Expected Result 16: Product table shows quantity and description correctly
            Assert.That(checkoutOverviewPage.IsProductTableDisplayedCorrectly(productName), Is.True, 
                "Expected: Product table shows quantity and description correctly");
            // Expected Result 17: Payment Information and Shipping Information sections display below product list
            Assert.That(checkoutOverviewPage.ArePaymentAndShippingInfoDisplayed(), Is.True, 
                "Expected: Payment Information and Shipping Information sections display below product list");
            // Step 17: Verify Item Total, Tax, and Total amounts are calculated and displayed
            TestContext.WriteLine("Step 17: Verify Item Total, Tax, and Total amounts are calculated and displayed");
            // Expected Result 18: Item Total, Tax, and Total are displayed with correct calculations
            Assert.That(checkoutOverviewPage.ArePriceCalculationsDisplayed(), Is.True, 
                "Expected: Item Total, Tax, and Total are displayed with correct calculations");
            string itemTotal = checkoutOverviewPage.GetItemTotal();
            string tax = checkoutOverviewPage.GetTax();
            string total = checkoutOverviewPage.GetTotal();
            TestContext.WriteLine($"Item Total: {itemTotal}");
            TestContext.WriteLine($"Tax: {tax}");
            TestContext.WriteLine($"Total: {total}");
            // Step 18: Click 'Finish' button
            TestContext.WriteLine("Step 18: Click Finish button");
            checkoutOverviewPage.ClickFinishButton();
            // Step 19: Wait for Finish page to load
            TestContext.WriteLine("Step 19: Wait for Finish page to load");
            checkoutCompletePage.WaitForFinishPageToLoad();
            // Expected Result 19: Finish page loads successfully
            Assert.That(checkoutCompletePage.IsFinishPageDisplayed(), Is.True, 
                "Expected: Finish page loads successfully");
            // Step 20: Verify 'Thank you for your order!' message displays
            TestContext.WriteLine("Step 20: Verify 'Thank you for your order!' message displays");
            // Expected Result 20: Success message and Pony Express Sauce Labs logo display
            Assert.That(checkoutCompletePage.IsThankYouMessageDisplayed(), Is.True, 
                "Expected: 'Thank you for your order!' message displays");
            Assert.That(checkoutCompletePage.IsPonyExpressLogoDisplayed(), Is.True, 
                "Expected: Pony Express Sauce Labs logo displays");
            TestContext.WriteLine("Test completed successfully: Complete checkout flow for single product verified");
        }
    }
}