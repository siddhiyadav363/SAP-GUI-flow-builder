/**
 * Page Object Model for Checkout: Overview Page
 * Contains locators and methods for Checkout Overview page interactions
 */
class CheckoutOverviewPage {
    // Locators using CSS selectors
    elements = {
        checkoutSummaryContainer: () => cy.get('#checkout_summary_container'),
        finishButton: () => cy.get('#finish'),
        cancelButton: () => cy.get('#cancel'),
        checkoutOverviewTitle: () => cy.get('.title').contains('Checkout: Overview'),
        hamburgerMenuButton: () => cy.get('#react-burger-menu-btn'),
        appLogo: () => cy.get('.app_logo'),
        shoppingCartContainer: () => cy.get('#shopping_cart_container'),
        cartItem: () => cy.get('.cart_item'),
        inventoryItemName: () => cy.get('.inventory_item_name'),
        summarySubtotalLabel: () => cy.get('.summary_subtotal_label'),
        summaryTaxLabel: () => cy.get('.summary_tax_label'),
        summaryTotalLabel: () => cy.get('.summary_total_label'),
        paymentInfoLabel: () => cy.contains('Payment Information'),
        shippingInfoLabel: () => cy.contains('Shipping Information')
    };
    /**
     * Wait for Checkout: Overview page to load
     */
    waitForCheckoutOverviewPageToLoad() {
        this.elements.checkoutSummaryContainer().should('be.visible');
        this.elements.finishButton().should('be.visible');
        cy.log('Checkout: Overview page loaded successfully');
    }
    /**
     * Verify 'Checkout: Overview' page displays with correct header
     */
    verifyCheckoutOverviewHeaderDisplayed() {
        this.elements.hamburgerMenuButton().should('be.visible');
        this.elements.appLogo().should('be.visible');
        this.elements.shoppingCartContainer().should('be.visible');
        cy.log("✓ Verified: 'Checkout: Overview' page displays with correct header (hamburger menu, SWAGLABS logo, cart icon)");
    }
    /**
     * Verify product details display correctly in overview
     * @param {string} productName - Expected product name (from {{product_name}} placeholder)
     */
    verifyProductDisplayedInOverview(productName) {
        this.elements.cartItem().should('be.visible');
        this.elements.inventoryItemName().should('contain.text', productName);
        cy.log(`Product in overview verified: ${productName}`);
    }
    /**
     * Verify product table shows quantity and description correctly
     * @param {string} productName - Expected product name (from {{product_name}} placeholder)
     */
    verifyProductTableDisplayedCorrectly(productName) {
        cy.get('.cart_quantity').contains('1').should('be.visible');
        this.verifyProductDisplayedInOverview(productName);
        cy.log('✓ Verified: Product table shows quantity and description correctly');
    }
    /**
     * Verify Payment Information and Shipping Information sections display
     */
    verifyPaymentAndShippingInfoDisplayed() {
        this.elements.paymentInfoLabel().should('be.visible');
        this.elements.shippingInfoLabel().should('be.visible');
        cy.log('✓ Verified: Payment Information and Shipping Information sections display below product list');
    }
    /**
     * Verify Item Total, Tax, and Total amounts are displayed
     */
    verifyPriceCalculationsDisplayed() {
        this.elements.summarySubtotalLabel().should('be.visible');
        this.elements.summaryTaxLabel().should('be.visible');
        this.elements.summaryTotalLabel().should('be.visible');
        cy.log('✓ Verified: Item Total, Tax, and Total are displayed with correct calculations');
    }
    /**
     * Get Item Total value
     * @returns {Cypress.Chainable<string>}
     */
    getItemTotal() {
        return this.elements.summarySubtotalLabel().invoke('text').then(text => {
            cy.log('Item Total: ' + text);
            return text;
        });
    }
    /**
     * Get Tax value
     * @returns {Cypress.Chainable<string>}
     */
    getTax() {
        return this.elements.summaryTaxLabel().invoke('text').then(text => {
            cy.log('Tax: ' + text);
            return text;
        });
    }
    /**
     * Get Total value
     * @returns {Cypress.Chainable<string>}
     */
    getTotal() {
        return this.elements.summaryTotalLabel().invoke('text').then(text => {
            cy.log('Total: ' + text);
            return text;
        });
    }
    /**
     * Click Finish button to complete the order
     */
    clickFinishButton() {
        this.elements.finishButton().click();
        cy.log('Clicked Finish button');
    }
}
export default CheckoutOverviewPage;