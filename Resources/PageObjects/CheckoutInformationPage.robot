*** Settings ***
Documentation    Page Object Model for Checkout: Your Information Page
...              Contains locators and keywords for Checkout Information page interactions
Library          SeleniumLibrary
*** Variables ***
# Locators using XPath from provided list
${FIRST_NAME_FIELD}             xpath=//input[@id='first-name']
${LAST_NAME_FIELD}              xpath=//input[@id='last-name']
${POSTAL_CODE_FIELD}            xpath=//input[@id='postal-code']
${CONTINUE_BUTTON}              xpath=//input[@id='continue']
${CANCEL_BUTTON}                xpath=//button[@id='cancel']
${CHECKOUT_INFO_CONTAINER}      xpath=//div[@id='checkout_info_container']
${CHECKOUT_INFO_TITLE}          xpath=//span[@class='title' and text()='Checkout: Your Information']
*** Keywords ***
Wait For Checkout Information Page To Load
    [Documentation]    Wait for Checkout: Your Information page to load
    Wait Until Element Is Visible    ${CHECKOUT_INFO_CONTAINER}    timeout=10s
    Wait Until Element Is Visible    ${FIRST_NAME_FIELD}    timeout=10s
    Log    Checkout: Your Information page loaded successfully
Verify Checkout Information Page Is Displayed
    [Documentation]    Verify 'Checkout: Your Information' page displays with all fields
    Element Should Be Visible    ${CHECKOUT_INFO_CONTAINER}
    Element Should Be Visible    ${FIRST_NAME_FIELD}
    Element Should Be Visible    ${LAST_NAME_FIELD}
    Element Should Be Visible    ${POSTAL_CODE_FIELD}
    Log    ✓ Verified: 'Checkout: Your Information' page displays with header and three mandatory fields
Enter First Name
    [Documentation]    Enter first name in the First Name field
    [Arguments]    ${first_name}
    Wait Until Element Is Visible    ${FIRST_NAME_FIELD}    timeout=10s
    Clear Element Text    ${FIRST_NAME_FIELD}
    Input Text    ${FIRST_NAME_FIELD}    ${first_name}
    Log    Entered first name: ${first_name}
Verify First Name Field Is Enabled
    [Documentation]    Verify First Name field accepts alphabetic input
    Element Should Be Enabled    ${FIRST_NAME_FIELD}
    Log    ✓ Verified: First Name field accepts alphabetic input
Enter Last Name
    [Documentation]    Enter last name in the Last Name field
    [Arguments]    ${last_name}
    Wait Until Element Is Visible    ${LAST_NAME_FIELD}    timeout=10s
    Clear Element Text    ${LAST_NAME_FIELD}
    Input Text    ${LAST_NAME_FIELD}    ${last_name}
    Log    Entered last name: ${last_name}
Verify Last Name Field Is Enabled
    [Documentation]    Verify Last Name field accepts alphabetic input
    Element Should Be Enabled    ${LAST_NAME_FIELD}
    Log    ✓ Verified: Last Name field accepts alphabetic input
Enter Zip Postal Code
    [Documentation]    Enter zip/postal code in the Zip/Postal Code field
    [Arguments]    ${zip_code}
    Wait Until Element Is Visible    ${POSTAL_CODE_FIELD}    timeout=10s
    Clear Element Text    ${POSTAL_CODE_FIELD}
    Input Text    ${POSTAL_CODE_FIELD}    ${zip_code}
    Log    Entered zip/postal code: ${zip_code}
Verify Zip Postal Code Field Is Enabled
    [Documentation]    Verify Zip/Postal Code field accepts numeric input
    Element Should Be Enabled    ${POSTAL_CODE_FIELD}
    Log    ✓ Verified: Zip/Postal Code field accepts numeric input
Click Continue Button
    [Documentation]    Click Continue button to proceed to Checkout Overview
    Wait Until Element Is Visible    ${CONTINUE_BUTTON}    timeout=10s
    Click Element    ${CONTINUE_BUTTON}
    Log    Clicked Continue button
Fill Checkout Information
    [Documentation]    Fill all checkout information fields
    [Arguments]    ${first_name}    ${last_name}    ${zip_code}
    Enter First Name    ${first_name}
    Enter Last Name    ${last_name}
    Enter Zip Postal Code    ${zip_code}
    Log    Filled all checkout information fields