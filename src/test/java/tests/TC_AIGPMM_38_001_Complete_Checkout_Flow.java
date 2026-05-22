package tests;
import com.microsoft.playwright.*;
import org.testng.Assert;
import org.testng.annotations.*;
import pages.*;
/**
 * Test Case: TC_AIGPMM-38_001
 * Story: AIGPMM-38
 * Description: Verify complete checkout process from cart to order completion with all mandatory fields filled
 * Objective: Validate that user can successfully complete checkout journey with valid information and receive order confirmation
 */
public class TC_AIGPMM_38_001_Complete_Checkout_Flow {
    private Playwright playwright;
    private Browser browser;
    private BrowserContext context;
    private Page page;
    private LoginPage loginPage;
    private ProductsPage productsPage;
    private CartPage cartPage;
    private CheckoutStepOnePage checkoutStepOnePage;
    private CheckoutStepTwoPage checkoutStepTwoPage;
    private CheckoutCompletePage checkoutCompletePage;
    @BeforeClass
    public void setupBrowser() {
        playwright = Playwright.create();
        browser = playwright.chromium().launch(new BrowserType.LaunchOptions().setHeadless(false));
    }
    @BeforeMethod
    public void setupTest() {
        context = browser.newContext();
        page = context.newPage();
        // Initialize Page Objects
        loginPage = new LoginPage(page);
        productsPage = new ProductsPage(page);
        cartPage = new CartPage(page);
        checkoutStepOnePage = new CheckoutStepOnePage(page);
        checkoutStepTwoPage = new CheckoutStepTwoPage(page);
        checkoutCompletePage = new CheckoutCompletePage(page);
    }
    @Test(description = "Verify complete checkout process from cart to order completion", priority = 1)
    public void testCompleteCheckoutFlow() {
        // Step 1: Navigate to application
        loginPage.navigateTo("{{base_url}}");
        Assert.assertTrue(loginPage.isLoginPageDisplayed(), "Login page should be displayed");
        // Steps 2-4: Login with valid credentials
        loginPage.enterUsername("{{username}}");
        loginPage.enterPassword("{{password}}");
        loginPage.clickLoginButton();
        // Step 5: Wait for Products page to load
        productsPage.waitForProductsPageToLoad();
        Assert.assertTrue(productsPage.isProductsPageDisplayed(), "Products page should be displayed");
        // Step 6: Add 'Sauce Labs Backpack' to cart
        productsPage.addSauceLabsBackpackToCart();
        Assert.assertTrue(productsPage.isRemoveButtonDisplayedForBackpack(), 
            "Add to cart button should change to Remove after click");
        // Step 7: Click on Cart icon in header
        productsPage.clickCartIcon();
        Assert.assertEquals(productsPage.getCartBadgeCount(), "1", 
            "Cart icon should show badge count '1'");
        // Step 8-9: Wait for 'Your Cart' page and verify product
        cartPage.waitForCartPageToLoad();
        Assert.assertTrue(cartPage.isCartPageDisplayed(), "'Your Cart' page should be displayed");
        Assert.assertTrue(cartPage.isProductInCart("Sauce Labs Backpack"), 
            "Product 'Sauce Labs Backpack' should appear in cart");
        // Step 10: Click Checkout button
        cartPage.clickCheckoutButton();
        // Step 11: Wait for 'Checkout: Your Information' page
        checkoutStepOnePage.waitForCheckoutStepOnePageToLoad();
        Assert.assertTrue(checkoutStepOnePage.isCheckoutStepOnePageDisplayed(), 
            "'Checkout: Your Information' page should be displayed with three mandatory fields visible");
        // Steps 12-14: Enter checkout information
        checkoutStepOnePage.enterFirstName("{{first_name}}");
        checkoutStepOnePage.enterLastName("{{last_name}}");
        checkoutStepOnePage.enterZipCode("{{zip_code}}");
        Assert.assertTrue(checkoutStepOnePage.areAllFieldsAcceptingInput(), 
            "All three fields should accept text input without errors");
        // Step 15: Click Continue button
        checkoutStepOnePage.clickContinueButton();
        Assert.assertTrue(checkoutStepOnePage.isContinueButtonClickable(), 
            "Continue button should become clickable");
        // Step 16-17: Wait for 'Checkout: Overview' page and verify details
        checkoutStepTwoPage.waitForCheckoutStepTwoPageToLoad();
        Assert.assertTrue(checkoutStepTwoPage.isCheckoutStepTwoPageDisplayed(), 
            "'Checkout: Overview' page should display");
        Assert.assertTrue(checkoutStepTwoPage.isProductDetailsDisplayed(), 
            "Product table with quantity and description should be displayed");
        Assert.assertTrue(checkoutStepTwoPage.isPaymentInformationDisplayed(), 
            "Payment information should be displayed");
        Assert.assertTrue(checkoutStepTwoPage.isShippingInformationDisplayed(), 
            "Shipping information should be displayed");
        Assert.assertTrue(checkoutStepTwoPage.arePricingDetailsDisplayed(), 
            "Item Total, Tax, Total amounts should be displayed");
        // Step 18: Click Finish button
        Assert.assertTrue(checkoutStepTwoPage.isFinishButtonVisible(), 
            "Finish button should be visible");
        Assert.assertTrue(checkoutStepTwoPage.isFinishButtonClickable(), 
            "Finish button should be clickable");
        checkoutStepTwoPage.clickFinishButton();
        // Step 19: Wait for 'Finish' page and verify order completion
        checkoutCompletePage.waitForCheckoutCompletePageToLoad();
        Assert.assertTrue(checkoutCompletePage.isCheckoutCompletePageDisplayed(), 
            "'Finish' page should be displayed");
        Assert.assertTrue(checkoutCompletePage.isThankYouMessageDisplayed(), 
            "'Thank you for your order!' message should be displayed");
        Assert.assertTrue(checkoutCompletePage.isPonyExpressLogoDisplayed(), 
            "'Pony Express Sauce Labs' logo should be displayed");
    }
    @AfterMethod
    public void tearDownTest() {
        if (context != null) {
            context.close();
        }
    }
    @AfterClass
    public void tearDownBrowser() {
        if (browser != null) {
            browser.close();
        }
        if (playwright != null) {
            playwright.close();
        }
    }
}