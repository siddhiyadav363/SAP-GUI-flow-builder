package pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Swag Labs Products Page
 * URL: https://www.saucedemo.com/inventory.html
 */
public class ProductsPage {
    private final Page page;
    // Locators using provided XPaths
    private final String productsPageTitle = "//span[text()='Products']";
    private final String inventoryContainer = "//div[@id='inventory_container']";
    private final String addToCartBackpack = "//button[@id='add-to-cart-sauce-labs-backpack']";
    private final String removeBackpack = "//button[@id='remove-sauce-labs-backpack']";
    private final String cartIcon = "//div[@id='shopping_cart_container']";
    private final String cartBadge = "//div[@id='shopping_cart_container']//span";
    public ProductsPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Products page to load
     */
    public void waitForProductsPageToLoad() {
        page.locator(productsPageTitle).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE)
            .setTimeout(10000));
        page.locator(inventoryContainer).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
    }
    /**
     * Check if Products page is displayed
     * @return true if Products page is displayed
     */
    public boolean isProductsPageDisplayed() {
        return page.locator(productsPageTitle).isVisible() && 
               page.locator(inventoryContainer).isVisible();
    }
    /**
     * Add Sauce Labs Backpack to cart
     */
    public void addSauceLabsBackpackToCart() {
        page.locator(addToCartBackpack).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(addToCartBackpack).click();
    }
    /**
     * Check if Remove button is displayed for Backpack (indicating item was added)
     * @return true if Remove button is displayed
     */
    public boolean isRemoveButtonDisplayedForBackpack() {
        try {
            return page.locator(removeBackpack).isVisible();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Click on Cart icon in header
     */
    public void clickCartIcon() {
        page.locator(cartIcon).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(cartIcon).click();
        page.waitForLoadState();
    }
    /**
     * Get cart badge count
     * @return The cart badge count as String
     */
    public String getCartBadgeCount() {
        try {
            page.locator(cartBadge).waitFor(new Page.Locator.WaitForOptions()
                .setState(WaitForSelectorState.VISIBLE)
                .setTimeout(5000));
            return page.locator(cartBadge).textContent();
        } catch (Exception e) {
            return "0";
        }
    }
}