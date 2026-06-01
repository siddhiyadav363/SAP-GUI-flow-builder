package com.swaglab.pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Checkout: Your Information Page
 * Contains locators and methods for Checkout Information page interactions
 */
public class CheckoutInformationPage {
    private Page page;
    // Locators using XPath from provided list
    private final String firstNameField = "//input[@id='first-name']";
    private final String lastNameField = "//input[@id='last-name']";
    private final String postalCodeField = "//input[@id='postal-code']";
    private final String continueButton = "//input[@id='continue']";
    private final String cancelButton = "//button[@id='cancel']";
    private final String checkoutInfoContainer = "//div[@id='checkout_info_container']";
    private final String checkoutInfoTitle = "//span[@class='title' and text()='Checkout: Your Information']";
    /**
     * Constructor for CheckoutInformationPage
     * @param page Playwright Page instance
     */
    public CheckoutInformationPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Checkout: Your Information page to load
     */
    public void waitForCheckoutInformationPageToLoad() {
        page.waitForSelector(checkoutInfoContainer, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.waitForSelector(firstNameField, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        System.out.println("Checkout: Your Information page loaded successfully");
    }
    /**
     * Verify 'Checkout: Your Information' page displays with all fields
     * @return boolean
     */
    public boolean isCheckoutInformationPageDisplayed() {
        try {
            boolean containerVisible = page.isVisible(checkoutInfoContainer);
            boolean firstNameVisible = page.isVisible(firstNameField);
            boolean lastNameVisible = page.isVisible(lastNameField);
            boolean postalCodeVisible = page.isVisible(postalCodeField);
            return containerVisible && firstNameVisible && lastNameVisible && postalCodeVisible;
        } catch (Exception e) {
            System.err.println("Checkout Information page not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Enter first name in the First Name field
     * @param firstName First name to enter (from {{first_name}} placeholder)
     */
    public void enterFirstName(String firstName) {
        page.fill(firstNameField, firstName);
        System.out.println("Entered first name: " + firstName);
    }
    /**
     * Verify First Name field accepts input
     * @return boolean
     */
    public boolean isFirstNameFieldEnabled() {
        return page.isEnabled(firstNameField);
    }
    /**
     * Enter last name in the Last Name field
     * @param lastName Last name to enter (from {{last_name}} placeholder)
     */
    public void enterLastName(String lastName) {
        page.fill(lastNameField, lastName);
        System.out.println("Entered last name: " + lastName);
    }
    /**
     * Verify Last Name field accepts input
     * @return boolean
     */
    public boolean isLastNameFieldEnabled() {
        return page.isEnabled(lastNameField);
    }
    /**
     * Enter zip/postal code in the Zip/Postal Code field
     * @param zipCode Zip/Postal code to enter (from {{zip_code}} placeholder)
     */
    public void enterZipPostalCode(String zipCode) {
        page.fill(postalCodeField, zipCode);
        System.out.println("Entered zip/postal code: " + zipCode);
    }
    /**
     * Verify Zip/Postal Code field accepts input
     * @return boolean
     */
    public boolean isZipPostalCodeFieldEnabled() {
        return page.isEnabled(postalCodeField);
    }
    /**
     * Click Continue button to proceed to Checkout Overview
     */
    public void clickContinueButton() {
        page.click(continueButton);
        System.out.println("Clicked Continue button");
    }
    /**
     * Fill all checkout information fields
     * @param firstName First name (from {{first_name}} placeholder)
     * @param lastName Last name (from {{last_name}} placeholder)
     * @param zipCode Zip/Postal code (from {{zip_code}} placeholder)
     */
    public void fillCheckoutInformation(String firstName, String lastName, String zipCode) {
        enterFirstName(firstName);
        enterLastName(lastName);
        enterZipPostalCode(zipCode);
        System.out.println("Filled all checkout information fields");
    }
}