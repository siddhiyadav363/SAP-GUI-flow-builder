*** Settings ***
Documentation    Page Object Model for Cart Page
...              Contains locators and keywords for Cart page interactions
Library          SeleniumLibrary
*** Variables ***
# Locators using XPath from provided list
${CART_CONTENTS_CONTAINER}      xpath=//div[@id='cart_contents_container']
${CHECKOUT_BUTTON}              xpath=//button[@id='checkout']
${CONTINUE_SHOPPING_BUTTON}     xpath=//button[@id='continue-shopping']
${REMOVE_SAUCE_LABS_BACKPACK}   xpath=//button[@id='remove-sauce-labs-backpack']
${INVENTORY_ITEM_NAME}          xpath=//div[@class='inventory_item_name']
${CART_QUANTITY}                xpath=//div[@class='cart_quantity']
*** Keywords ***
Wait For Cart Page To Load
    [Documentation]    Wait for Cart page to load
    Wait Until Element Is Visible    ${CART_CONTENTS_CONTAINER}    timeout=10s
    Log    Cart page loaded successfully
Verify Product Displayed In Cart
    [Documentation]    Verify Cart page displays with correct product name
    [Arguments]    ${product_name}
    Wait Until Element Is Visible    ${INVENTORY_ITEM_NAME}    timeout=10s
    ${actual_product_name}=    Get Text    ${INVENTORY_ITEM_NAME}
    Should Contain    ${actual_product_name}    ${product_name}
    Log    Product in cart - Expected: ${product_name}, Actual: ${actual_product_name}
Verify Product With Quantity In Cart
    [Documentation]    Verify product appears in cart with quantity 1
    [Arguments]    ${product_name}
    ${quantity_xpath}=    Set Variable    xpath=//div[@class='cart_quantity' and text()='1']
    Wait Until Element Is Visible    ${quantity_xpath}    timeout=10s
    Element Should Be Visible    ${quantity_xpath}
    Verify Product Displayed In Cart    ${product_name}
    Log    ✓ Verified: Cart page displays with correct product name '${product_name}' and quantity 1
Verify Checkout Button Is Visible And Clickable
    [Documentation]    Verify Checkout button is visible and clickable
    Element Should Be Visible    ${CHECKOUT_BUTTON}
    Element Should Be Enabled    ${CHECKOUT_BUTTON}
    Log    ✓ Verified: Checkout button is visible and clickable
Click Checkout Button
    [Documentation]    Click Checkout button to proceed to checkout
    Wait Until Element Is Visible    ${CHECKOUT_BUTTON}    timeout=10s
    Click Element    ${CHECKOUT_BUTTON}
    Log    Clicked Checkout button