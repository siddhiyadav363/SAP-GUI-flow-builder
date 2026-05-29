/**
 * Finish page POM class
 */
/**
 * LoginPage page object for Swag Labs
 */
class LoginPage {
  constructor() {
    // XPaths extracted from provided locators
    this.root = "//div[@id='root']";
    this.usernameInput = "//input[@id='user-name']";
    this.passwordInput = "//input[@id='password']";
    this.loginButton = "//input[@id='login-button']";
  }
  /**
   * enterUsername
   * @param {string} username - uses data-driven placeholder in tests
   */
  enterUsername(username) {
    // Cypress automatically retries; include defensive then/catch for logging
    cy.xpath(this.usernameInput).should('be.visible').then(($el) => {
      try {
        cy.wrap($el).clear().type(username);
      } catch (e) {
        cy.log('Error typing username: ' + e);
        throw e;
      }
    });
  }
  /**
   * enterPassword
   * @param {string} password - uses data-driven placeholder in tests
   */
  enterPassword(password) {
    cy.xpath(this.passwordInput).should('be.visible').then(($el) => {
      try {
        cy.wrap($el).clear().type(password);
      } catch (e) {
        cy.log('Error typing password: ' + e);
        throw e;
      }
    });
  }
  /**
   * clickLogin
   */
  clickLogin() {
    cy.xpath(this.loginButton).should('be.visible').click();
  }
}
export default LoginPage;