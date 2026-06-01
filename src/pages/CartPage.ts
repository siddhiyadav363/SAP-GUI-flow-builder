/**
 * Page Object Model for Cart Page
 * Contains locators and methods for Cart page interactions
 */
import { WebDriver, By } from 'selenium-webdriver';
import { BasePage } from './BasePage';
export class CartPage extends BasePage {
    // Locators using XPath from provided list
    private cartContentsContainer = By.xpath("//div[@id='cart_contents_container']");
    private checkoutButton = By.xpath("//button[@id='checkout']");
    private continueShoppingButton = By.xpath("//button[@id='continue-shopping']");
    private removeSauceLabsBackpack = By.xpath("//button[@id='remove-sauce-labs-backpack']");
    private inventoryItemName = By.xpath("//div[@class='inventory_item_name']");
    private cartQuantity = By.xpath("//div[@class='cart_quantity']");
    constructor(driver: WebDriver) {
        super(driver);
    }
    /**
     * Wait for Cart page to load
     */
    async waitForCartPageToLoad(): Promise<void> {
        await this.waitForElementVisible(this.cartContentsContainer);
        console.log('Cart page loaded successfully');
    }
    /**
     * Verify Cart page displays with correct product name
     * @param productName - Expected product name (from {{product_name}} placeholder)
     */
    async verifyProductDisplayedInCart(productName: string): Promise<void> {
        await this.waitForElementVisible(this.inventoryItemName);
        const actualProductName = await this.getText(this.inventoryItemName);
        if (!actualProductName.includes(productName)) {
            throw new Error(`Product name mismatch. Expected: ${productName}, Actual: ${actualProductName}`);
        }
        console.log(`Product in cart - Expected: ${productName}, Actual: ${actualProductName}`);
    }
    /**
     * Verify product appears in cart with quantity 1
     * @param productName - Expected product name (from {{product_name}} placeholder)
     */
    async verifyProductWithQuantityInCart(productName: string): Promise<void> {
        const quantityXPath = By.xpath("//div[@class='cart_quantity' and text()='1']");
        await this.waitForElementVisible(quantityXPath);
        const isQuantityVisible = await this.isElementVisible(quantityXPath);
        if (!isQuantityVisible) {
            throw new Error('Quantity 1 not displayed in cart');
        }
        await this.verifyProductDisplayedInCart(productName);
        this.logVerification(`Cart page displays with correct product name '${productName}' and quantity 1`);
    }
    /**
     * Verify Checkout button is visible and clickable
     */
    async verifyCheckoutButtonVisibleAndClickable(): Promise<void> {
        const isVisible = await this.isElementVisible(this.checkoutButton);
        const isEnabled = await this.isElementEnabled(this.checkoutButton);
        if (!isVisible || !isEnabled) {
            throw new Error('Checkout button is not visible or clickable');
        }
        this.logVerification('Checkout button is visible and clickable');
    }
    /**
     * Click Checkout button to proceed to checkout
     */
    async clickCheckoutButton(): Promise<void> {
        await this.click(this.checkoutButton);
        console.log('Clicked Checkout button');
    }
}