package com.swaglab.pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Checkout Complete (Finish) Page
 * Contains locators and methods for Checkout Complete page interactions
 */
public class CheckoutCompletePage {
    private Page page;
    // Locators using XPath from provided list
    private final String checkoutCompleteContainer = "//div[@id='checkout_complete_container']";
    private final String thankYouMessage = "//h2";
    private final String backHomeButton = "//button[@id='back-to-products']";
    private final String ponyExpressImage = "//img[@class='pony_express']";
    /**
     * Constructor for CheckoutCompletePage
     * @param page Playwright Page instance
     */
    public CheckoutCompletePage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Finish page to load
     */
    public void waitForFinishPageToLoad() {
        page.waitForSelector(checkoutCompleteContainer, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.waitForSelector(thankYouMessage, new Page.WaitForSelectorOptions()
            .setState(WaitForSelectorState.VISIBLE));
        System.out.println("Finish page loaded successfully");
    }
    /**
     * Verify Finish page loads successfully
     * @return boolean
     */
    public boolean isFinishPageDisplayed() {
        try {
            return page.isVisible(checkoutCompleteContainer);
        } catch (Exception e) {
            System.err.println("Finish page not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Get the thank you message text
     * @return String
     */
    public String getThankYouMessage() {
        String text = page.textContent(thankYouMessage);
        System.out.println("Thank you message: " + text);
        return text;
    }
    /**
     * Verify 'Thank you for your order!' message displays
     * @return boolean
     */
    public boolean isThankYouMessageDisplayed() {
        try {
            String message = getThankYouMessage();
            return message.contains("Thank you for your order!");
        } catch (Exception e) {
            System.err.println("Thank you message not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify Pony Express Sauce Labs logo displays
     * @return boolean
     */
    public boolean isPonyExpressLogoDisplayed() {
        try {
            return page.isVisible(ponyExpressImage);
        } catch (Exception e) {
            System.err.println("Pony Express logo not displayed: " + e.getMessage());
            return false;
        }
    }
    /**
     * Verify success message and logo display
     * @return boolean
     */
    public boolean isOrderCompletionConfirmed() {
        boolean messageDisplayed = isThankYouMessageDisplayed();
        boolean logoDisplayed = isPonyExpressLogoDisplayed();
        System.out.println("Order completion confirmed - Message: " + messageDisplayed + ", Logo: " + logoDisplayed);
        return messageDisplayed && logoDisplayed;
    }
}