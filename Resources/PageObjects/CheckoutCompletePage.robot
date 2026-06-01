*** Settings ***
Documentation    Page Object Model for Checkout Complete (Finish) Page
...              Contains locators and keywords for Checkout Complete page interactions
Library          SeleniumLibrary
*** Variables ***
# Locators using XPath from provided list
${CHECKOUT_COMPLETE_CONTAINER}    xpath=//div[@id='checkout_complete_container']
${THANK_YOU_MESSAGE}              xpath=//h2
${BACK_HOME_BUTTON}               xpath=//button[@id='back-to-products']
${PONY_EXPRESS_IMAGE}             xpath=//img[@class='pony_express']
*** Keywords ***
Wait For Finish Page To Load
    [Documentation]    Wait for Finish page to load
    Wait Until Element Is Visible    ${CHECKOUT_COMPLETE_CONTAINER}    timeout=10s
    Wait Until Element Is Visible    ${THANK_YOU_MESSAGE}    timeout=10s
    Log    Finish page loaded successfully
Verify Finish Page Is Displayed
    [Documentation]    Verify Finish page loads successfully
    Element Should Be Visible    ${CHECKOUT_COMPLETE_CONTAINER}
    Log    ✓ Verified: Finish page loads successfully
Get Thank You Message
    [Documentation]    Get the thank you message text
    Wait Until Element Is Visible    ${THANK_YOU_MESSAGE}    timeout=10s
    ${message}=    Get Text    ${THANK_YOU_MESSAGE}
    Log    Thank you message: ${message}
    RETURN    ${message}
Verify Thank You Message Is Displayed
    [Documentation]    Verify 'Thank you for your order!' message displays
    ${message}=    Get Thank You Message
    Should Contain    ${message}    Thank you for your order!
    Log    ✓ Verified: 'Thank you for your order!' message displays
Verify Pony Express Logo Is Displayed
    [Documentation]    Verify Pony Express Sauce Labs logo displays
    Element Should Be Visible    ${PONY_EXPRESS_IMAGE}
    Log    ✓ Verified: Pony Express Sauce Labs logo displays
Verify Order Completion Confirmed
    [Documentation]    Verify success message and logo display
    Verify Thank You Message Is Displayed
    Verify Pony Express Logo Is Displayed
    Log    Order completion confirmed - Message and Logo displayed