*** Settings ***
Documentation    Page Object Model for Checkout: Overview Page
...              Contains locators and keywords for Checkout Overview page interactions
Library          SeleniumLibrary
*** Variables ***
# Locators using XPath from provided list
${CHECKOUT_SUMMARY_CONTAINER}    xpath=//div[@id='checkout_summary_container']
${FINISH_BUTTON}                 xpath=//button[@id='finish']
${CANCEL_BUTTON}                 xpath=//button[@id='cancel']
${CHECKOUT_OVERVIEW_TITLE}       xpath=//span[@class='title' and text()='Checkout: Overview']
${HAMBURGER_MENU_BUTTON}         xpath=//button[@id='react-burger-menu-btn']
${APP_LOGO}                      xpath=//div[@class='app_logo']
${SHOPPING_CART_CONTAINER}       xpath=//div[@id='shopping_cart_container']
${CART_ITEM}                     xpath=//div[@class='cart_item']
${INVENTORY_ITEM_NAME}           xpath=//div[@class='inventory_item_name']
${SUMMARY_SUBTOTAL_LABEL}        xpath=//div[@class='summary_subtotal_label']
${SUMMARY_TAX_LABEL}             xpath=//div[@class='summary_tax_label']
${SUMMARY_TOTAL_LABEL}           xpath=//div[@class='summary_total_label']
${PAYMENT_INFO_LABEL}            xpath=//div[contains(text(), 'Payment Information')]
${SHIPPING_INFO_LABEL}           xpath=//div[contains(text(), 'Shipping Information')]
*** Keywords ***
Wait For Checkout Overview Page To Load
    [Documentation]    Wait for Checkout: Overview page to load
    Wait Until Element Is Visible    ${CHECKOUT_SUMMARY_CONTAINER}    timeout=10s
    Wait Until Element Is Visible    ${FINISH_BUTTON}    timeout=10s
    Log    Checkout: Overview page loaded successfully
Verify Checkout Overview Header Is Displayed
    [Documentation]    Verify 'Checkout: Overview' page displays with correct header
    Element Should Be Visible    ${HAMBURGER_MENU_BUTTON}
    Element Should Be Visible    ${APP_LOGO}
    Element Should Be Visible    ${SHOPPING_CART_CONTAINER}
    Log    ✓ Verified: 'Checkout: Overview' page displays with correct header (hamburger menu, SWAGLABS logo, cart icon)
Verify Product Displayed In Overview
    [Documentation]    Verify product details display correctly in overview
    [Arguments]    ${product_name}
    Wait Until Element Is Visible    ${CART_ITEM}    timeout=10s
    Wait Until Element Is Visible    ${INVENTORY_ITEM_NAME}    timeout=10s
    ${actual_product_name}=    Get Text    ${INVENTORY_ITEM_NAME}
    Should Contain    ${actual_product_name}    ${product_name}
    Log    Product in overview - Expected: ${product_name}, Actual: ${actual_product_name}
Verify Product Table Displayed Correctly
    [Documentation]    Verify product table shows quantity and description correctly
    [Arguments]    ${product_name}
    ${cart_quantity_xpath}=    Set Variable    xpath=//div[@class='cart_quantity' and text()='1']
    Wait Until Element Is Visible    ${cart_quantity_xpath}    timeout=10s
    Element Should Be Visible    ${cart_quantity_xpath}
    Verify Product Displayed In Overview    ${product_name}
    Log    ✓ Verified: Product table shows quantity and description correctly
Verify Payment And Shipping Info Displayed
    [Documentation]    Verify Payment Information and Shipping Information sections display
    Element Should Be Visible    ${PAYMENT_INFO_LABEL}
    Element Should Be Visible    ${SHIPPING_INFO_LABEL}
    Log    ✓ Verified: Payment Information and Shipping Information sections display below product list
Verify Price Calculations Displayed
    [Documentation]    Verify Item Total, Tax, and Total amounts are displayed
    Element Should Be Visible    ${SUMMARY_SUBTOTAL_LABEL}
    Element Should Be Visible    ${SUMMARY_TAX_LABEL}
    Element Should Be Visible    ${SUMMARY_TOTAL_LABEL}
    Log    ✓ Verified: Item Total, Tax, and Total are displayed with correct calculations
Get Item Total
    [Documentation]    Get Item Total value
    ${item_total}=    Get Text    ${SUMMARY_SUBTOTAL_LABEL}
    Log    Item Total: ${item_total}
    RETURN    ${item_total}
Get Tax
    [Documentation]    Get Tax value
    ${tax}=    Get Text    ${SUMMARY_TAX_LABEL}
    Log    Tax: ${tax}
    RETURN    ${tax}
Get Total
    [Documentation]    Get Total value
    ${total}=    Get Text    ${SUMMARY_TOTAL_LABEL}
    Log    Total: ${total}
    RETURN    ${total}
Click Finish Button
    [Documentation]    Click Finish button to complete the order
    Wait Until Element Is Visible    ${FINISH_BUTTON}    timeout=10s
    Click Element    ${FINISH_BUTTON}
    Log    Clicked Finish button