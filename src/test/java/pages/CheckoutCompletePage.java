package pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Swag Labs Checkout Complete Page (Finish)
 * URL: https://www.saucedemo.com/checkout-complete.html
 */
public class CheckoutCompletePage {
    private final Page page;
    // Locators using provided XPaths
    private final String checkoutCompleteContainer = "//div[@id='checkout_complete_container']";
    private final String thankYouMessage = "//h2[text()='Thank you for your order!']";
    private final String ponyExpressLogo = "//img[@class='pony_express']";
    private final String backHomeButton = "//button[@id='back-to-products']";
    private final String checkoutCompleteTitle = "//span[text()='Checkout: Complete!']";
    public CheckoutCompletePage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Checkout Complete page to load
     */
    public void waitForCheckoutCompletePageToLoad() {
        page.locator(checkoutCompleteTitle).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE)
            .setTimeout(10000));
        page.locator(checkoutCompleteContainer).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
    }
    /**
     * Check if Checkout Complete page is displayed
     * @return true if Checkout Complete page is displayed
     */
    public boolean isCheckoutCompletePageDisplayed() {
        return page.locator(checkoutCompleteTitle).isVisible() && 
               page.locator(checkoutCompleteContainer).isVisible();
    }
    /**
     * Check if 'Thank you for your order!' message is displayed
     * @return true if thank you message is displayed
     */
    public boolean isThankYouMessageDisplayed() {
        page.locator(thankYouMessage).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE)
            .setTimeout(5000));
        return page.locator(thankYouMessage).isVisible();
    }
    /**
     * Check if Pony Express logo is displayed
     * @return true if Pony Express logo is displayed
     */
    public boolean isPonyExpressLogoDisplayed() {
        try {
            return page.locator(ponyExpressLogo).isVisible();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Click Back Home button
     */
    public void clickBackHomeButton() {
        page.locator(backHomeButton).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(backHomeButton).click();
        page.waitForLoadState();
    }
}