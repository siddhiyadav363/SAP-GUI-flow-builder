/**
 * Page Object Model for Checkout: Overview Page
 * Contains locators and methods for Checkout Overview page interactions
 */
import { WebDriver, By } from 'selenium-webdriver';
import { BasePage } from './BasePage';
export class CheckoutOverviewPage extends BasePage {
    // Locators using XPath from provided list
    private checkoutSummaryContainer = By.xpath("//div[@id='checkout_summary_container']");
    private finishButton = By.xpath("//button[@id='finish']");
    private cancelButton = By.xpath("//button[@id='cancel']");
    private checkoutOverviewTitle = By.xpath("//span[@class='title' and text()='Checkout: Overview']");
    private hamburgerMenuButton = By.xpath("//button[@id='react-burger-menu-btn']");
    private appLogo = By.xpath("//div[@class='app_logo']");
    private shoppingCartContainer = By.xpath("//div[@id='shopping_cart_container']");
    private cartItem = By.xpath("//div[@class='cart_item']");
    private inventoryItemName = By.xpath("//div[@class='inventory_item_name']");
    private summarySubtotalLabel = By.xpath("//div[@class='summary_subtotal_label']");
    private summaryTaxLabel = By.xpath("//div[@class='summary_tax_label']");
    private summaryTotalLabel = By.xpath("//div[@class='summary_total_label']");
    private paymentInfoLabel = By.xpath("//div[contains(text(), 'Payment Information')]");
    private shippingInfoLabel = By.xpath("//div[contains(text(), 'Shipping Information')]");
    constructor(driver: WebDriver) {
        super(driver);
    }
    /**
     * Wait for Checkout: Overview page to load
     */
    async waitForCheckoutOverviewPageToLoad(): Promise<void> {
        await this.waitForElementVisible(this.checkoutSummaryContainer);
        await this.waitForElementVisible(this.finishButton);
        console.log('Checkout: Overview page loaded successfully');
    }
    /**
     * Verify 'Checkout: Overview' page displays with correct header
     */
    async verifyCheckoutOverviewHeaderDisplayed(): Promise<void> {
        const isHamburgerVisible = await this.isElementVisible(this.hamburgerMenuButton);
        const isLogoVisible = await this.isElementVisible(this.appLogo);
        const isCartVisible = await this.isElementVisible(this.shoppingCartContainer);
        if (!isHamburgerVisible || !isLogoVisible || !isCartVisible) {
            throw new Error('Checkout Overview header not displayed correctly');
        }
        this.logVerification("'Checkout: Overview' page displays with correct header (hamburger menu, SWAGLABS logo, cart icon)");
    }
    /**
     * Verify product details display correctly in overview
     * @param productName - Expected product name (from {{product_name}} placeholder)
     */
    async verifyProductDisplayedInOverview(productName: string): Promise<void> {
        await this.waitForElementVisible(this.cartItem);
        await this.waitForElementVisible(this.inventoryItemName);
        const actualProductName = await this.getText(this.inventoryItemName);
        if (!actualProductName.includes(productName)) {
            throw new Error(`Product name mismatch. Expected: ${productName}, Actual: ${actualProductName}`);
        }
        console.log(`Product in overview - Expected: ${productName}, Actual: ${actualProductName}`);
    }
    /**
     * Verify product table shows quantity and description correctly
     * @param productName - Expected product name (from {{product_name}} placeholder)
     */
    async verifyProductTableDisplayedCorrectly(productName: string): Promise<void> {
        const cartQuantityXPath = By.xpath("//div[@class='cart_quantity' and text()='1']");
        await this.waitForElementVisible(cartQuantityXPath);
        const isQuantityVisible = await this.isElementVisible(cartQuantityXPath);
        if (!isQuantityVisible) {
            throw new Error('Quantity 1 not displayed');
        }
        await this.verifyProductDisplayedInOverview(productName);
        this.logVerification('Product table shows quantity and description correctly');
    }
    /**
     * Verify Payment Information and Shipping Information sections display
     */
    async verifyPaymentAndShippingInfoDisplayed(): Promise<void> {
        const isPaymentVisible = await this.isElementVisible(this.paymentInfoLabel);
        const isShippingVisible = await this.isElementVisible(this.shippingInfoLabel);
        if (!isPaymentVisible || !isShippingVisible) {
            throw new Error('Payment or Shipping information not displayed');
        }
        this.logVerification('Payment Information and Shipping Information sections display below product list');
    }
    /**
     * Verify Item Total, Tax, and Total amounts are displayed
     */
    async verifyPriceCalculationsDisplayed(): Promise<void> {
        const isSubtotalVisible = await this.isElementVisible(this.summarySubtotalLabel);
        const isTaxVisible = await this.isElementVisible(this.summaryTaxLabel);
        const isTotalVisible = await this.isElementVisible(this.summaryTotalLabel);
        if (!isSubtotalVisible || !isTaxVisible || !isTotalVisible) {
            throw new Error('Price calculations not displayed');
        }
        this.logVerification('Item Total, Tax, and Total are displayed with correct calculations');
    }
    /**
     * Get Item Total value
     */
    async getItemTotal(): Promise<string> {
        const text = await this.getText(this.summarySubtotalLabel);
        console.log(`Item Total: ${text}`);
        return text;
    }
    /**
     * Get Tax value
     */
    async getTax(): Promise<string> {
        const text = await this.getText(this.summaryTaxLabel);
        console.log(`Tax: ${text}`);
        return text;
    }
    /**
     * Get Total value
     */
    async getTotal(): Promise<string> {
        const text = await this.getText(this.summaryTotalLabel);
        console.log(`Total: ${text}`);
        return text;
    }
    /**
     * Click Finish button to complete the order
     */
    async clickFinishButton(): Promise<void> {
        await this.click(this.finishButton);
        console.log('Clicked Finish button');
    }
}