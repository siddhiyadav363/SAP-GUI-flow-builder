const { defineConfig } = require('cypress');
module.exports = defineConfig({
  e2e: {
    // Base URL for the application (can be overridden by environment variable)
    baseUrl: process.env.BASE_URL || '{{base_url}}',
    // Viewport settings
    viewportWidth: 1920,
    viewportHeight: 1080,
    // Timeout settings
    defaultCommandTimeout: 10000,
    pageLoadTimeout: 30000,
    requestTimeout: 10000,
    responseTimeout: 10000,
    // Video and screenshot settings
    video: true,
    videoCompression: 32,
    screenshotOnRunFailure: true,
    // Test isolation
    testIsolation: true,
    // Retry configuration
    retries: {
      runMode: 2,
      openMode: 0
    },
    // Spec pattern
    specPattern: 'cypress/e2e/**/*.cy.{js,jsx,ts,tsx}',
    setupNodeEvents(on, config) {
      // Implement node event listeners here
      // Custom task for logging
      on('task', {
        log(message) {
          console.log(message);
          return null;
        }
      });
      // Override config with environment variables if provided
      config.env.base_url = process.env.BASE_URL || config.env.base_url || '{{base_url}}';
      config.env.username = process.env.USERNAME || config.env.username || '{{username}}';
      config.env.password = process.env.PASSWORD || config.env.password || '{{password}}';
      config.env.first_name = process.env.FIRST_NAME || config.env.first_name || '{{first_name}}';
      config.env.last_name = process.env.LAST_NAME || config.env.last_name || '{{last_name}}';
      config.env.zip_code = process.env.ZIP_CODE || config.env.zip_code || '{{zip_code}}';
      config.env.product_name = process.env.PRODUCT_NAME || config.env.product_name || '{{product_name}}';
      return config;
    },
  },
});