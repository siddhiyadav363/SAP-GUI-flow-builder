package com.swaglab.pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Login Page
 * Contains locators and methods for Login page interactions
 */
public class LoginPage {
    private Page page;
    // Locators using XPath from provided list
    private final String usernameField = "//input[@id='user-name']";
    private final String passwordField = "//input[@id='password']";
    private final String loginButton = "//input[@id='login-button']";
    private final String loginCredentialsDiv = "//div[@id='login_credentials']";
    private final String loginButtonContainer = "//div[@id='login_button_container']";
    /**
     * Constructor for LoginPage
     * @param page Playwright Page instance
     */
    public LoginPage(Page page) {
        this.page = page;
    }
    /**
     * Navigate to the login page
     * @param url Base URL from test data using {{base_url}} placeholder
     */
    public void navigateToLoginPage(String url) {
        page.navigate(url);
        System.out.println("Navigated to: " + url);
    }
    /**
     * Verify that login page has loaded successfully
     * @return boolean
     */
    public boolean isLoginPageDisplayed() {
        try {
            page.waitForSelector(usernameField, new Page.WaitForSelectorOptions()
                .setState(WaitForSelectorState.VISIBLE));
            boolean usernameVisible = page.isVisible(usernameField);
            boolean passwordVisible = page.isVisible(passwordField);
            boolean loginBtnVisible = page.isVisible(loginButton);
            return usernameVisible && passwordVisible && loginBtnVisible;
        } catch (Exception e) {
            System.err.println("Login page not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Enter username in the username field
     * @param username Username to enter (from {{username}} placeholder)
     */
    public void enterUsername(String username) {
        page.fill(usernameField, username);
        System.out.println("Entered username: " + username);
    }
    /**
     * Verify username field accepts input
     * @return boolean
     */
    public boolean isUsernameFieldEnabled() {
        return page.isEnabled(usernameField);
    }
    /**
     * Get the current value in username field
     * @return String
     */
    public String getUsernameValue() {
        return page.inputValue(usernameField);
    }
    /**
     * Enter password in the password field
     * @param password Password to enter (from {{password}} placeholder)
     */
    public void enterPassword(String password) {
        page.fill(passwordField, password);
        String maskedPassword = "*".repeat(password.length());
        System.out.println("Entered password: " + maskedPassword);
    }
    /**
     * Verify password field accepts input and masks characters
     * @return boolean
     */
    public boolean isPasswordFieldMasked() {
        String fieldType = page.getAttribute(passwordField, "type");
        return "password".equals(fieldType);
    }
    /**
     * Click the Login button
     */
    public void clickLoginButton() {
        page.click(loginButton);
        System.out.println("Clicked Login button");
    }
    /**
     * Verify login button is clickable
     * @return boolean
     */
    public boolean isLoginButtonClickable() {
        return page.isEnabled(loginButton) && page.isVisible(loginButton);
    }
    /**
     * Perform complete login action
     * @param username Username for login (from {{username}} placeholder)
     * @param password Password for login (from {{password}} placeholder)
     */
    public void login(String username, String password) {
        enterUsername(username);
        enterPassword(password);
        clickLoginButton();
        System.out.println("Login action completed");
    }
}