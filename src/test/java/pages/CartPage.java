package pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Swag Labs Cart Page
 * URL: https://www.saucedemo.com/cart.html
 */
public class CartPage {
    private final Page page;
    // Locators using provided XPaths
    private final String cartContentsContainer = "//div[@id='cart_contents_container']";
    private final String checkoutButton = "//button[@id='checkout']";
    private final String continueShoppingButton = "//button[@id='continue-shopping']";
    private final String yourCartTitle = "//span[text()='Your Cart']";
    public CartPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Cart page to load
     */
    public void waitForCartPageToLoad() {
        page.locator(yourCartTitle).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE)
            .setTimeout(10000));
        page.locator(cartContentsContainer).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
    }
    /**
     * Check if Cart page is displayed
     * @return true if Cart page is displayed
     */
    public boolean isCartPageDisplayed() {
        return page.locator(yourCartTitle).isVisible() && 
               page.locator(cartContentsContainer).isVisible();
    }
    /**
     * Check if specific product is in cart
     * @param productName The product name to check
     * @return true if product is in cart
     */
    public boolean isProductInCart(String productName) {
        String productLocator = String.format("//div[@class='inventory_item_name' and text()='%s']", productName);
        try {
            return page.locator(productLocator).isVisible();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Click Checkout button
     */
    public void clickCheckoutButton() {
        page.locator(checkoutButton).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(checkoutButton).click();
        page.waitForLoadState();
    }
}