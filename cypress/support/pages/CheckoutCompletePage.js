/**
 * Page Object Model for Checkout Complete (Finish) Page
 * Contains locators and methods for Checkout Complete page interactions
 */
class CheckoutCompletePage {
    // Locators using CSS selectors
    elements = {
        checkoutCompleteContainer: () => cy.get('#checkout_complete_container'),
        thankYouMessage: () => cy.get('h2'),
        backHomeButton: () => cy.get('#back-to-products'),
        ponyExpressImage: () => cy.get('.pony_express')
    };
    /**
     * Wait for Finish page to load
     */
    waitForFinishPageToLoad() {
        this.elements.checkoutCompleteContainer().should('be.visible');
        this.elements.thankYouMessage().should('be.visible');
        cy.log('Finish page loaded successfully');
    }
    /**
     * Verify Finish page loads successfully
     */
    verifyFinishPageDisplayed() {
        this.elements.checkoutCompleteContainer().should('be.visible');
        cy.log('✓ Verified: Finish page loads successfully');
    }
    /**
     * Get the thank you message text
     * @returns {Cypress.Chainable<string>}
     */
    getThankYouMessage() {
        return this.elements.thankYouMessage().invoke('text').then(text => {
            cy.log('Thank you message: ' + text);
            return text;
        });
    }
    /**
     * Verify 'Thank you for your order!' message displays
     */
    verifyThankYouMessageDisplayed() {
        this.elements.thankYouMessage().should('contain.text', 'Thank you for your order!');
        cy.log("✓ Verified: 'Thank you for your order!' message displays");
    }
    /**
     * Verify Pony Express Sauce Labs logo displays
     */
    verifyPonyExpressLogoDisplayed() {
        this.elements.ponyExpressImage().should('be.visible');
        cy.log('✓ Verified: Pony Express Sauce Labs logo displays');
    }
    /**
     * Verify success message and logo display
     */
    verifyOrderCompletionConfirmed() {
        this.verifyThankYouMessageDisplayed();
        this.verifyPonyExpressLogoDisplayed();
        cy.log('Order completion confirmed - Message and Logo displayed');
    }
}
export default CheckoutCompletePage;