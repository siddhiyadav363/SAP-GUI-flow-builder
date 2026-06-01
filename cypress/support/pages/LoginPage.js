/**
 * Page Object Model for Login Page
 * Contains locators and methods for Login page interactions
 */
class LoginPage {
    // Locators using CSS selectors (converted from XPath for Cypress best practices)
    elements = {
        usernameField: () => cy.get('#user-name'),
        passwordField: () => cy.get('#password'),
        loginButton: () => cy.get('#login-button'),
        loginCredentialsDiv: () => cy.get('#login_credentials'),
        loginButtonContainer: () => cy.get('#login_button_container')
    };
    /**
     * Navigate to the login page
     * @param {string} url - Base URL from test data using {{base_url}} placeholder
     */
    navigateToLoginPage(url) {
        cy.visit(url);
        cy.log('Navigated to: ' + url);
    }
    /**
     * Verify that login page has loaded successfully
     */
    verifyLoginPageDisplayed() {
        this.elements.usernameField().should('be.visible');
        this.elements.passwordField().should('be.visible');
        this.elements.loginButton().should('be.visible');
        cy.log('✓ Verified: Login page loads successfully');
    }
    /**
     * Enter username in the username field
     * @param {string} username - Username to enter (from {{username}} placeholder)
     */
    enterUsername(username) {
        this.elements.usernameField().clear().type(username);
        cy.log('Entered username: ' + username);
    }
    /**
     * Verify username field accepts input
     */
    verifyUsernameFieldEnabled() {
        this.elements.usernameField().should('be.enabled');
        cy.log('✓ Verified: Username field accepts input');
    }
    /**
     * Get the current value in username field
     * @returns {Cypress.Chainable<string>}
     */
    getUsernameValue() {
        return this.elements.usernameField().invoke('val');
    }
    /**
     * Enter password in the password field
     * @param {string} password - Password to enter (from {{password}} placeholder)
     */
    enterPassword(password) {
        this.elements.passwordField().clear().type(password);
        const maskedPassword = '*'.repeat(password.length);
        cy.log('Entered password: ' + maskedPassword);
    }
    /**
     * Verify password field accepts input and masks characters
     */
    verifyPasswordFieldMasked() {
        this.elements.passwordField().should('have.attr', 'type', 'password');
        cy.log('✓ Verified: Password field accepts input and masks characters');
    }
    /**
     * Click the Login button
     */
    clickLoginButton() {
        this.elements.loginButton().click();
        cy.log('Clicked Login button');
    }
    /**
     * Verify login button is clickable
     */
    verifyLoginButtonClickable() {
        this.elements.loginButton().should('be.enabled').and('be.visible');
        cy.log('✓ Verified: Login button is clickable');
    }
    /**
     * Perform complete login action
     * @param {string} username - Username for login (from {{username}} placeholder)
     * @param {string} password - Password for login (from {{password}} placeholder)
     */
    login(username, password) {
        this.enterUsername(username);
        this.enterPassword(password);
        this.clickLoginButton();
        cy.log('Login action completed');
    }
}
export default LoginPage;