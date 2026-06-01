/**
 * Custom Cypress commands for enhanced test capabilities
 */
/**
 * Custom command to log test steps
 * @param {number} stepNumber - Step number
 * @param {string} description - Step description
 */
Cypress.Commands.add('logStep', (stepNumber, description) => {
    cy.log(`\n**Step ${stepNumber}:** ${description}`);
    cy.task('log', `\nStep ${stepNumber}: ${description}`);
});
/**
 * Custom command to log verification results
 * @param {string} description - Verification description
 */
Cypress.Commands.add('logVerification', (description) => {
    cy.log(`✓ **Verified:** ${description}`);
    cy.task('log', `✓ Verified: ${description}`);
});