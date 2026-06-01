/**
 * Cypress support file for e2e tests
 * Import commands, custom behaviors, and global configurations
 */
// Import custom commands
import './commands';
// Global before hook
before(() => {
    cy.log('=== Starting Cypress Test Suite ===');
});
// Global after hook
after(() => {
    cy.log('=== Completed Cypress Test Suite ===');
});