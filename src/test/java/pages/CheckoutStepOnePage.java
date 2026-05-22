package pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Swag Labs Checkout Step One Page (Your Information)
 * URL: https://www.saucedemo.com/checkout-step-one.html
 */
public class CheckoutStepOnePage {
    private final Page page;
    // Locators using provided XPaths
    private final String checkoutInfoContainer = "//div[@id='checkout_info_container']";
    private final String firstNameInput = "//input[@id='first-name']";
    private final String lastNameInput = "//input[@id='last-name']";
    private final String postalCodeInput = "//input[@id='postal-code']";
    private final String continueButton = "//input[@id='continue']";
    private final String cancelButton = "//button[@id='cancel']";
    private final String checkoutStepOneTitle = "//span[text()='Checkout: Your Information']";
    public CheckoutStepOnePage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Checkout Step One page to load
     */
    public void waitForCheckoutStepOnePageToLoad() {
        page.locator(checkoutStepOneTitle).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE)
            .setTimeout(10000));
        page.locator(checkoutInfoContainer).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
    }
    /**
     * Check if Checkout Step One page is displayed
     * @return true if Checkout Step One page is displayed
     */
    public boolean isCheckoutStepOnePageDisplayed() {
        return page.locator(checkoutStepOneTitle).isVisible() && 
               page.locator(firstNameInput).isVisible() && 
               page.locator(lastNameInput).isVisible() && 
               page.locator(postalCodeInput).isVisible();
    }
    /**
     * Enter first name
     * @param firstName The first name to enter
     */
    public void enterFirstName(String firstName) {
        page.locator(firstNameInput).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(firstNameInput).clear();
        page.locator(firstNameInput).fill(firstName);
    }
    /**
     * Enter last name
     * @param lastName The last name to enter
     */
    public void enterLastName(String lastName) {
        page.locator(lastNameInput).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(lastNameInput).clear();
        page.locator(lastNameInput).fill(lastName);
    }
    /**
     * Enter zip/postal code
     * @param zipCode The zip code to enter
     */
    public void enterZipCode(String zipCode) {
        page.locator(postalCodeInput).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(postalCodeInput).clear();
        page.locator(postalCodeInput).fill(zipCode);
    }
    /**
     * Check if all fields are accepting input (no errors)
     * @return true if all fields are enabled and visible
     */
    public boolean areAllFieldsAcceptingInput() {
        return page.locator(firstNameInput).isEnabled() && 
               page.locator(lastNameInput).isEnabled() && 
               page.locator(postalCodeInput).isEnabled();
    }
    /**
     * Check if Continue button is clickable
     * @return true if Continue button is enabled
     */
    public boolean isContinueButtonClickable() {
        return page.locator(continueButton).isEnabled();
    }
    /**
     * Click Continue button
     */
    public void clickContinueButton() {
        page.locator(continueButton).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(continueButton).click();
        page.waitForLoadState();
    }
}