/**
 * Finish page POM class
 */
/**
 * CheckoutYourInformationPage page object
 */
class CheckoutYourInformationPage {
  constructor() {
    this.checkoutInfoContainer = "//div[@id='checkout_info_container']";
    this.firstNameInput = "//input[@id='first-name']";
    this.lastNameInput = "//input[@id='last-name']";
    this.postalCodeInput = "//input[@id='postal-code']";
    this.continueButton = "//input[@id='continue']";
  }
  /**
   * enterFirstName
   * @param {string} firstName - uses placeholder in tests
   */
  enterFirstName(firstName) {
    cy.xpath(this.firstNameInput).should('be.visible').then(($el) => {
      try {
        cy.wrap($el).clear().type(firstName);
      } catch (e) {
        cy.log('Error entering first name: ' + e);
        throw e;
      }
    });
  }
  /**
   * enterLastName
   * @param {string} lastName - uses placeholder in tests
   */
  enterLastName(lastName) {
    cy.xpath(this.lastNameInput).should('be.visible').then(($el) => {
      try {
        cy.wrap($el).clear().type(lastName);
      } catch (e) {
        cy.log('Error entering last name: ' + e);
        throw e;
      }
    });
  }
  /**
   * enterPostalCode
   * @param {string} postalCode - uses placeholder in tests
   */
  enterPostalCode(postalCode) {
    cy.xpath(this.postalCodeInput).should('be.visible').then(($el) => {
      try {
        cy.wrap($el).clear().type(postalCode);
      } catch (e) {
        cy.log('Error entering postal code: ' + e);
        throw e;
      }
    });
  }
  /**
   * clickContinue
   */
  clickContinue() {
    cy.xpath(this.continueButton).should('be.visible').click();
  }
}
export default CheckoutYourInformationPage;