package com.swaglab.pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Cart Page
 * Contains locators and methods for Cart page interactions
 */
public class CartPage {
    private Page page;
    // Locators using XPath from provided list
    private final String cartContentsContainer = "//div[@id='cart_contents_container']";
    private final String checkoutButton = "//button[@id='checkout']";
    private final String continueShoppingButton = "//button[@id='continue-shopping']";
    private final String removeSauceLabsBackpack = "//button[@id='remove-sauce-labs-backpack']";
    private final String inventoryItemName = "//div[@class='inventory_item_name']";
    private final String cartQuantity = "//div[@class='cart_quantity']";
    /**
     * Constructor for CartPage
     * @param page Playwright Page instance
     */
    public CartPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Cart page to load
     */
    public void waitForCartPageToLoad() {
        page.waitForSelector(cartContentsContainer, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        System.out.println("Cart page loaded successfully");
    }
    /**
     * Verify Cart page displays with correct product name
     * @param productName Expected product name (from {{product_name}} placeholder)
     * @return boolean
     */
    public boolean isProductDisplayedInCart(String productName) {
        try {
            page.waitForSelector(inventoryItemName, new Page.WaitForSelectorOptions()
                .setState(WaitForSelectorState.VISIBLE));
            String actualProductName = page.textContent(inventoryItemName);
            System.out.println("Product in cart - Expected: " + productName + ", Actual: " + actualProductName);
            return actualProductName.contains(productName);
        } catch (Exception e) {
            System.err.println("Product not displayed in cart: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify product appears in cart with quantity 1
     * @param productName Expected product name (from {{product_name}} placeholder)
     * @return boolean
     */
    public boolean verifyProductWithQuantityInCart(String productName) {
        try {
            String quantityXPath = "//div[@class='cart_quantity' and text()='1']";
            page.waitForSelector(quantityXPath, new Page.WaitForSelectorOptions()
                .setState(WaitForSelectorState.VISIBLE));
            boolean quantityVisible = page.isVisible(quantityXPath);
            boolean productVisible = isProductDisplayedInCart(productName);
            return quantityVisible && productVisible;
        } catch (Exception e) {
            System.err.println("Product with quantity not verified: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify Checkout button is visible and clickable
     * @return boolean
     */
    public boolean isCheckoutButtonVisibleAndClickable() {
        try {
            boolean isVisible = page.isVisible(checkoutButton);
            boolean isEnabled = page.isEnabled(checkoutButton);
            return isVisible && isEnabled;
        } catch (Exception e) {
            System.err.println("Checkout button not visible/clickable: " + e.getMessage());
            return false;
        }
    }
    /**
     * Click Checkout button to proceed to checkout
     */
    public void clickCheckoutButton() {
        page.click(checkoutButton);
        System.out.println("Clicked Checkout button");
    }
}