package pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Swag Labs Login Page
 * URL: https://www.saucedemo.com/
 */
public class LoginPage {
    private final Page page;
    // Locators using provided XPaths
    private final String usernameInput = "//input[@id='user-name']";
    private final String passwordInput = "//input[@id='password']";
    private final String loginButton = "//input[@id='login-button']";
    private final String loginButtonContainer = "//div[@id='login_button_container']";
    private final String loginCredentials = "//div[@id='login_credentials']";
    public LoginPage(Page page) {
        this.page = page;
    }
    /**
     * Navigate to the application URL
     * @param url The application URL
     */
    public void navigateTo(String url) {
        page.navigate(url);
        page.waitForLoadState();
    }
    /**
     * Check if login page is displayed
     * @return true if login page is displayed
     */
    public boolean isLoginPageDisplayed() {
        return page.locator(loginButtonContainer).isVisible() && 
               page.locator(usernameInput).isVisible() && 
               page.locator(passwordInput).isVisible();
    }
    /**
     * Enter username in the username field
     * @param username The username to enter
     */
    public void enterUsername(String username) {
        page.locator(usernameInput).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(usernameInput).clear();
        page.locator(usernameInput).fill(username);
    }
    /**
     * Enter password in the password field
     * @param password The password to enter
     */
    public void enterPassword(String password) {
        page.locator(passwordInput).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(passwordInput).clear();
        page.locator(passwordInput).fill(password);
    }
    /**
     * Click the Login button
     */
    public void clickLoginButton() {
        page.locator(loginButton).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(loginButton).click();
        page.waitForLoadState();
    }
}