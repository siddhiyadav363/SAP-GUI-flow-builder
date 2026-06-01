package com.swaglab.pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Checkout: Overview Page
 * Contains locators and methods for Checkout Overview page interactions
 */
public class CheckoutOverviewPage {
    private Page page;
    // Locators using XPath from provided list
    private final String checkoutSummaryContainer = "//div[@id='checkout_summary_container']";
    private final String finishButton = "//button[@id='finish']";
    private final String cancelButton = "//button[@id='cancel']";
    private final String checkoutOverviewTitle = "//span[@class='title' and text()='Checkout: Overview']";
    private final String hamburgerMenuButton = "//button[@id='react-burger-menu-btn']";
    private final String appLogo = "//div[@class='app_logo']";
    private final String shoppingCartContainer = "//div[@id='shopping_cart_container']";
    private final String cartItem = "//div[@class='cart_item']";
    private final String inventoryItemName = "//div[@class='inventory_item_name']";
    private final String summarySubtotalLabel = "//div[@class='summary_subtotal_label']";
    private final String summaryTaxLabel = "//div[@class='summary_tax_label']";
    private final String summaryTotalLabel = "//div[@class='summary_total_label']";
    private final String paymentInfoLabel = "//div[contains(text(), 'Payment Information')]";
    private final String shippingInfoLabel = "//div[contains(text(), 'Shipping Information')]";
    /**
     * Constructor for CheckoutOverviewPage
     * @param page Playwright Page instance
     */
    public CheckoutOverviewPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Checkout: Overview page to load
     */
    public void waitForCheckoutOverviewPageToLoad() {
        page.waitForSelector(checkoutSummaryContainer, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.waitForSelector(finishButton, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        System.out.println("Checkout: Overview page loaded successfully");
    }
    /**
     * Verify 'Checkout: Overview' page displays with correct header
     * @return boolean
     */
    public boolean isCheckoutOverviewHeaderDisplayed() {
        try {
            boolean hamburgerVisible = page.isVisible(hamburgerMenuButton);
            boolean logoVisible = page.isVisible(appLogo);
            boolean cartVisible = page.isVisible(shoppingCartContainer);
            return hamburgerVisible && logoVisible && cartVisible;
        } catch (Exception e) {
            System.err.println("Checkout Overview header not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify product details display correctly in overview
     * @param productName Expected product name (from {{product_name}} placeholder)
     * @return boolean
     */
    public boolean isProductDisplayedInOverview(String productName) {
        try {
            page.waitForSelector(cartItem, new Page.WaitForSelectorOptions()
                .setState(WaitForSelectorState.VISIBLE));
            String actualProductName = page.textContent(inventoryItemName);
            System.out.println("Product in overview - Expected: " + productName + ", Actual: " + actualProductName);
            return actualProductName.contains(productName);
        } catch (Exception e) {
            System.err.println("Product not displayed in overview: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify product table shows quantity and description correctly
     * @param productName Expected product name (from {{product_name}} placeholder)
     * @return boolean
     */
    public boolean isProductTableDisplayedCorrectly(String productName) {
        try {
            String cartQuantityXPath = "//div[@class='cart_quantity' and text()='1']";
            page.waitForSelector(cartQuantityXPath, new Page.WaitForSelectorOptions()
                .setState(WaitForSelectorState.VISIBLE));
            boolean quantityVisible = page.isVisible(cartQuantityXPath);
            boolean productVisible = isProductDisplayedInOverview(productName);
            return quantityVisible && productVisible;
        } catch (Exception e) {
            System.err.println("Product table not displayed correctly: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify Payment Information and Shipping Information sections display
     * @return boolean
     */
    public boolean arePaymentAndShippingInfoDisplayed() {
        try {
            boolean paymentVisible = page.isVisible(paymentInfoLabel);
            boolean shippingVisible = page.isVisible(shippingInfoLabel);
            return paymentVisible && shippingVisible;
        } catch (Exception e) {
            System.err.println("Payment/Shipping info not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify Item Total, Tax, and Total amounts are displayed
     * @return boolean
     */
    public boolean arePriceCalculationsDisplayed() {
        try {
            boolean subtotalVisible = page.isVisible(summarySubtotalLabel);
            boolean taxVisible = page.isVisible(summaryTaxLabel);
            boolean totalVisible = page.isVisible(summaryTotalLabel);
            return subtotalVisible && taxVisible && totalVisible;
        } catch (Exception e) {
            System.err.println("Price calculations not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Get Item Total value
     * @return String
     */
    public String getItemTotal() {
        String text = page.textContent(summarySubtotalLabel);
        System.out.println("Item Total: " + text);
        return text;
    }
    /**
     * Get Tax value
     * @return String
     */
    public String getTax() {
        String text = page.textContent(summaryTaxLabel);
        System.out.println("Tax: " + text);
        return text;
    }
    /**
     * Get Total value
     * @return String
     */
    public String getTotal() {
        String text = page.textContent(summaryTotalLabel);
        System.out.println("Total: " + text);
        return text;
    }
    /**
     * Click Finish button to complete the order
     */
    public void clickFinishButton() {
        page.click(finishButton);
        System.out.println("Clicked Finish button");
    }
}