package com.swaglab.pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Products Page
 * Contains locators and methods for Products page interactions
 */
public class ProductsPage {
    private Page page;
    // Locators using XPath from provided list
    private final String productsPageTitle = "//span[text()='Products']";
    private final String addToCartSauceLabsBackpack = "//button[@id='add-to-cart-sauce-labs-backpack']";
    private final String removeSauceLabsBackpack = "//button[@id='remove-sauce-labs-backpack']";
    private final String shoppingCartContainer = "//div[@id='shopping_cart_container']";
    private final String cartBadge = "//span[@class='shopping_cart_badge']";
    private final String inventoryContainer = "//div[@id='inventory_container']";
    /**
     * Constructor for ProductsPage
     * @param page Playwright Page instance
     */
    public ProductsPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Products page to load
     */
    public void waitForProductsPageToLoad() {
        page.waitForSelector(productsPageTitle, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.waitForSelector(inventoryContainer, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        System.out.println("Products page loaded successfully");
    }
    /**
     * Verify Products page loads with product listings
     * @return boolean
     */
    public boolean isProductsPageDisplayed() {
        try {
            boolean titleVisible = page.isVisible(productsPageTitle);
            boolean inventoryVisible = page.isVisible(inventoryContainer);
            return titleVisible && inventoryVisible;
        } catch (Exception e) {
            System.err.println("Products page not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Click 'Add to cart' button for Sauce Labs Backpack
     */
    public void clickAddToCartForSauceLabsBackpack() {
        page.click(addToCartSauceLabsBackpack);
        System.out.println("Clicked 'Add to cart' for Sauce Labs Backpack");
    }
    /**
     * Verify 'Add to cart' button changes to 'Remove' after clicking
     * @return boolean
     */
    public boolean isRemoveButtonDisplayed() {
        try {
            page.waitForSelector(removeSauceLabsBackpack, new Page.WaitForSelectorOptions()
                .setTimeout(5000)
                .setState(WaitForSelectorState.VISIBLE));
            return page.isVisible(removeSauceLabsBackpack);
        } catch (Exception e) {
            System.err.println("Remove button not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Get cart badge count
     * @return String
     */
    public String getCartBadgeCount() {
        try {
            page.waitForSelector(cartBadge, new Page.WaitForSelectorOptions()
                .setTimeout(5000)
                .setState(WaitForSelectorState.VISIBLE));
            return page.textContent(cartBadge);
        } catch (Exception e) {
            System.err.println("Cart badge not found: " + e.getMessage());
            return "0";
        }
    }
    /**
     * Verify cart icon shows badge with expected count
     * @param expectedCount Expected badge count
     * @return boolean
     */
    public boolean isCartBadgeCountCorrect(String expectedCount) {
        String actualCount = getCartBadgeCount();
        System.out.println("Cart badge count - Expected: " + expectedCount + ", Actual: " + actualCount);
        return actualCount.equals(expectedCount);
    }
    /**
     * Click on Cart icon to navigate to Cart page
     */
    public void clickCartIcon() {
        page.click(shoppingCartContainer);
        System.out.println("Clicked Cart icon");
    }
}