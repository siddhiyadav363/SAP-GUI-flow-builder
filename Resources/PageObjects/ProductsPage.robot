*** Settings ***
Documentation    Page Object Model for Products Page
...              Contains locators and keywords for Products page interactions
Library          SeleniumLibrary
*** Variables ***
# Locators using XPath from provided list
${PRODUCTS_PAGE_TITLE}                xpath=//span[text()='Products']
${ADD_TO_CART_SAUCE_LABS_BACKPACK}    xpath=//button[@id='add-to-cart-sauce-labs-backpack']
${REMOVE_SAUCE_LABS_BACKPACK}         xpath=//button[@id='remove-sauce-labs-backpack']
${SHOPPING_CART_CONTAINER}            xpath=//div[@id='shopping_cart_container']
${CART_BADGE}                         xpath=//span[@class='shopping_cart_badge']
${INVENTORY_CONTAINER}                xpath=//div[@id='inventory_container']
*** Keywords ***
Wait For Products Page To Load
    [Documentation]    Wait for Products page to load
    Wait Until Element Is Visible    ${PRODUCTS_PAGE_TITLE}    timeout=10s
    Wait Until Element Is Visible    ${INVENTORY_CONTAINER}    timeout=10s
    Log    Products page loaded successfully
Verify Products Page Is Displayed
    [Documentation]    Verify Products page loads with product listings
    Element Should Be Visible    ${PRODUCTS_PAGE_TITLE}
    Element Should Be Visible    ${INVENTORY_CONTAINER}
    Log    ✓ Verified: Products page loads with product listings
Click Add To Cart For Sauce Labs Backpack
    [Documentation]    Click 'Add to cart' button for Sauce Labs Backpack
    Wait Until Element Is Visible    ${ADD_TO_CART_SAUCE_LABS_BACKPACK}    timeout=10s
    Click Element    ${ADD_TO_CART_SAUCE_LABS_BACKPACK}
    Log    Clicked 'Add to cart' for Sauce Labs Backpack
Verify Remove Button Is Displayed
    [Documentation]    Verify 'Add to cart' button changes to 'Remove' after clicking
    Wait Until Element Is Visible    ${REMOVE_SAUCE_LABS_BACKPACK}    timeout=10s
    Element Should Be Visible    ${REMOVE_SAUCE_LABS_BACKPACK}
    Log    ✓ Verified: 'Add to cart' button changes to 'Remove' after clicking
Get Cart Badge Count
    [Documentation]    Get cart badge count
    Wait Until Element Is Visible    ${CART_BADGE}    timeout=10s
    ${count}=    Get Text    ${CART_BADGE}
    RETURN    ${count}
Verify Cart Badge Count
    [Documentation]    Verify cart icon shows badge with expected count
    [Arguments]    ${expected_count}
    ${actual_count}=    Get Cart Badge Count
    Should Be Equal    ${actual_count}    ${expected_count}
    Log    Cart badge count - Expected: ${expected_count}, Actual: ${actual_count}
    Log    ✓ Verified: Cart icon shows badge with '${expected_count}'
Click Cart Icon
    [Documentation]    Click on Cart icon to navigate to Cart page
    Wait Until Element Is Visible    ${SHOPPING_CART_CONTAINER}    timeout=10s
    Click Element    ${SHOPPING_CART_CONTAINER}
    Log    Clicked Cart icon