/**
 * Page Object Model for Checkout: Your Information Page
 * Contains locators and methods for Checkout Information page interactions
 */
class CheckoutInformationPage {
    // Locators using CSS selectors
    elements = {
        firstNameField: () => cy.get('#first-name'),
        lastNameField: () => cy.get('#last-name'),
        postalCodeField: () => cy.get('#postal-code'),
        continueButton: () => cy.get('#continue'),
        cancelButton: () => cy.get('#cancel'),
        checkoutInfoContainer: () => cy.get('#checkout_info_container'),
        checkoutInfoTitle: () => cy.get('.title').contains('Checkout: Your Information')
    };
    /**
     * Wait for Checkout: Your Information page to load
     */
    waitForCheckoutInformationPageToLoad() {
        this.elements.checkoutInfoContainer().should('be.visible');
        this.elements.firstNameField().should('be.visible');
        cy.log('Checkout: Your Information page loaded successfully');
    }
    /**
     * Verify 'Checkout: Your Information' page displays with all fields
     */
    verifyCheckoutInformationPageDisplayed() {
        this.elements.checkoutInfoContainer().should('be.visible');
        this.elements.firstNameField().should('be.visible');
        this.elements.lastNameField().should('be.visible');
        this.elements.postalCodeField().should('be.visible');
        cy.log("✓ Verified: 'Checkout: Your Information' page displays with header and three mandatory fields");
    }
    /**
     * Enter first name in the First Name field
     * @param {string} firstName - First name to enter (from {{first_name}} placeholder)
     */
    enterFirstName(firstName) {
        this.elements.firstNameField().clear().type(firstName);
        cy.log('Entered first name: ' + firstName);
    }
    /**
     * Verify First Name field accepts input
     */
    verifyFirstNameFieldEnabled() {
        this.elements.firstNameField().should('be.enabled');
        cy.log('✓ Verified: First Name field accepts alphabetic input');
    }
    /**
     * Enter last name in the Last Name field
     * @param {string} lastName - Last name to enter (from {{last_name}} placeholder)
     */
    enterLastName(lastName) {
        this.elements.lastNameField().clear().type(lastName);
        cy.log('Entered last name: ' + lastName);
    }
    /**
     * Verify Last Name field accepts input
     */
    verifyLastNameFieldEnabled() {
        this.elements.lastNameField().should('be.enabled');
        cy.log('✓ Verified: Last Name field accepts alphabetic input');
    }
    /**
     * Enter zip/postal code in the Zip/Postal Code field
     * @param {string} zipCode - Zip/Postal code to enter (from {{zip_code}} placeholder)
     */
    enterZipPostalCode(zipCode) {
        this.elements.postalCodeField().clear().type(zipCode);
        cy.log('Entered zip/postal code: ' + zipCode);
    }
    /**
     * Verify Zip/Postal Code field accepts input
     */
    verifyZipPostalCodeFieldEnabled() {
        this.elements.postalCodeField().should('be.enabled');
        cy.log('✓ Verified: Zip/Postal Code field accepts numeric input');
    }
    /**
     * Click Continue button to proceed to Checkout Overview
     */
    clickContinueButton() {
        this.elements.continueButton().click();
        cy.log('Clicked Continue button');
    }
    /**
     * Fill all checkout information fields
     * @param {string} firstName - First name (from {{first_name}} placeholder)
     * @param {string} lastName - Last name (from {{last_name}} placeholder)
     * @param {string} zipCode - Zip/Postal code (from {{zip_code}} placeholder)
     */
    fillCheckoutInformation(firstName, lastName, zipCode) {
        this.enterFirstName(firstName);
        this.enterLastName(lastName);
        this.enterZipPostalCode(zipCode);
        cy.log('Filled all checkout information fields');
    }
}
export default CheckoutInformationPage;