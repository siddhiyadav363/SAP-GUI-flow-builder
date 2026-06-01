package com.swaglab.base;
import com.microsoft.playwright.*;
import org.testng.annotations.*;
import java.nio.file.Paths;
/**
 * Base test class for Playwright test automation
 * Provides browser initialization, configuration, and cleanup
 */
public class TestBase {
    protected Browser browser;
    protected BrowserContext context;
    protected Page page;
    protected Playwright playwright;
    // Test data with {{placeholders}} for data-driven execution
    protected String baseUrl = System.getProperty("base_url", "{{base_url}}");
    protected String username = System.getProperty("username", "{{username}}");
    protected String password = System.getProperty("password", "{{password}}");
    protected String firstName = System.getProperty("first_name", "{{first_name}}");
    protected String lastName = System.getProperty("last_name", "{{last_name}}");
    protected String zipCode = System.getProperty("zip_code", "{{zip_code}}");
    protected String productName = System.getProperty("product_name", "{{product_name}}");
    /**
     * Setup method to initialize Playwright and Browser before each test
     */
    @BeforeMethod
    public void setUp() {
        try {
            // Initialize Playwright
            playwright = Playwright.create();
            // Launch browser with options
            browser = playwright.chromium().launch(new BrowserType.LaunchOptions()
                .setHeadless(false)
                .setSlowMo(50));
            // Create browser context with viewport
            context = browser.newContext(new Browser.NewContextOptions()
                .setViewportSize(1920, 1080)
                .setAcceptDownloads(true));
            // Set default timeout
            context.setDefaultTimeout(30000);
            // Create new page
            page = context.newPage();
            System.out.println("Browser initialized successfully: Chromium");
        } catch (Exception e) {
            System.err.println("Error initializing browser: " + e.getMessage());
            throw e;
        }
    }
    /**
     * Teardown method to close browser after each test
     */
    @AfterMethod
    public void tearDown() {
        try {
            if (page != null) {
                page.close();
            }
            if (context != null) {
                context.close();
            }
            if (browser != null) {
                browser.close();
            }
            if (playwright != null) {
                playwright.close();
            }
            System.out.println("Browser closed successfully");
        } catch (Exception e) {
            System.err.println("Error closing browser: " + e.getMessage());
        }
    }
    /**
     * Helper method to wait for element to be visible
     * @param selector Element selector (XPath or CSS)
     * @param timeout Wait timeout in milliseconds
     */
    protected void waitForElementVisible(String selector, double timeout) {
        page.waitForSelector(selector, new Page.WaitForSelectorOptions()
            .setTimeout(timeout)
            .setState(WaitForSelectorState.VISIBLE));
    }
    /**
     * Helper method to check if element is visible
     * @param selector Element selector
     * @return boolean
     */
    protected boolean isElementVisible(String selector) {
        try {
            return page.isVisible(selector);
        } catch (Exception e) {
            return false;
        }
    }
    /**
     * Helper method to log test steps
     * @param stepNumber Step number
     * @param description Step description
     */
    protected void logStep(int stepNumber, String description) {
        System.out.println("\nStep " + stepNumber + ": " + description);
    }
    /**
     * Helper method to log verification results
     * @param description Verification description
     */
    protected void logVerification(String description) {
        System.out.println("✓ Verified: " + description);
    }
}