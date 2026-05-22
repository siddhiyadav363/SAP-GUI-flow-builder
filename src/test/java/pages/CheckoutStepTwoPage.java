package pages;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.options.WaitForSelectorState;
/**
 * Page Object Model for Swag Labs Checkout Step Two Page (Overview)
 * URL: https://www.saucedemo.com/checkout-step-two.html
 */
public class CheckoutStepTwoPage {
    private final Page page;
    // Locators using provided XPaths
    private final String checkoutSummaryContainer = "//div[@id='checkout_summary_container']";
    private final String finishButton = "//button[@id='finish']";
    private final String cancelButton = "//button[@id='cancel']";
    private final String checkoutStepTwoTitle = "//span[text()='Checkout: Overview']";
    private final String itemTotal = "//div[@class='summary_subtotal_label']";
    private final String tax = "//div[@class='summary_tax_label']";
    private final String total = "//div[@class='summary_total_label']";
    private final String paymentInformation = "//div[@class='summary_info' and contains(., 'Payment Information')]";
    private final String shippingInformation = "//div[@class='summary_info' and contains(., 'Shipping Information')]";
    private final String cartItem = "//div[@class='cart_item']";
    public CheckoutStepTwoPage(Page page) {
        this.page = page;
    }
    /**
     * Wait for Checkout Step Two page to load
     */
    public void waitForCheckoutStepTwoPageToLoad() {
        page.locator(checkoutStepTwoTitle).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE)
            .setTimeout(10000));
        page.locator(checkoutSummaryContainer).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
    }
    /**
     * Check if Checkout Step Two page is displayed
     * @return true if Checkout Step Two page is displayed
     */
    public boolean isCheckoutStepTwoPageDisplayed() {
        return page.locator(checkoutStepTwoTitle).isVisible() && 
               page.locator(checkoutSummaryContainer).isVisible();
    }
    /**
     * Check if product details are displayed (quantity and description)
     * @return true if product details are displayed
     */
    public boolean isProductDetailsDisplayed() {
        try {
            return page.locator(cartItem).isVisible() && 
                   page.locator("//div[@class='cart_quantity']").isVisible() && 
                   page.locator("//div[@class='inventory_item_name']").isVisible();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Check if payment information is displayed
     * @return true if payment information is displayed
     */
    public boolean isPaymentInformationDisplayed() {
        try {
            return page.locator(paymentInformation).isVisible();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Check if shipping information is displayed
     * @return true if shipping information is displayed
     */
    public boolean isShippingInformationDisplayed() {
        try {
            return page.locator(shippingInformation).isVisible();
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Check if pricing details are displayed (Item Total, Tax, Total)
     * @return true if all pricing details are displayed
     */
    public boolean arePricingDetailsDisplayed() {
        return page.locator(itemTotal).isVisible() && 
               page.locator(tax).isVisible() && 
               page.locator(total).isVisible();
    }
    /**
     * Check if Finish button is visible
     * @return true if Finish button is visible
     */
    public boolean isFinishButtonVisible() {
        return page.locator(finishButton).isVisible();
    }
    /**
     * Check if Finish button is clickable
     * @return true if Finish button is enabled
     */
    public boolean isFinishButtonClickable() {
        return page.locator(finishButton).isEnabled();
    }
    /**
     * Click Finish button
     */
    public void clickFinishButton() {
        page.locator(finishButton).waitFor(new Page.Locator.WaitForOptions()
            .setState(WaitForSelectorState.VISIBLE));
        page.locator(finishButton).click();
        page.waitForLoadState();
    }
}