*** Settings ***
Documentation    Page Object Model for Login Page
...              Contains locators and keywords for Login page interactions
Library          SeleniumLibrary
*** Variables ***
# Locators using XPath from provided list
${LOGIN_USERNAME_FIELD}        xpath=//input[@id='user-name']
${LOGIN_PASSWORD_FIELD}        xpath=//input[@id='password']
${LOGIN_BUTTON}                xpath=//input[@id='login-button']
${LOGIN_CREDENTIALS_DIV}       xpath=//div[@id='login_credentials']
${LOGIN_BUTTON_CONTAINER}      xpath=//div[@id='login_button_container']
*** Keywords ***
Navigate To Login Page
    [Documentation]    Navigate to the login page
    [Arguments]    ${url}
    Go To    ${url}
    Log    Navigated to: ${url}
Verify Login Page Is Displayed
    [Documentation]    Verify that login page has loaded successfully
    Wait Until Element Is Visible    ${LOGIN_USERNAME_FIELD}    timeout=10s
    Wait Until Element Is Visible    ${LOGIN_PASSWORD_FIELD}    timeout=10s
    Wait Until Element Is Visible    ${LOGIN_BUTTON}    timeout=10s
    Log    ✓ Verified: Login page loads successfully
Enter Username
    [Documentation]    Enter username in the username field
    [Arguments]    ${username}
    Wait Until Element Is Visible    ${LOGIN_USERNAME_FIELD}    timeout=10s
    Clear Element Text    ${LOGIN_USERNAME_FIELD}
    Input Text    ${LOGIN_USERNAME_FIELD}    ${username}
    Log    Entered username: ${username}
Verify Username Field Accepts Input
    [Documentation]    Verify username field accepts input
    Element Should Be Enabled    ${LOGIN_USERNAME_FIELD}
    Log    ✓ Verified: Username field accepts input
Get Username Value
    [Documentation]    Get the current value in username field
    ${value}=    Get Value    ${LOGIN_USERNAME_FIELD}
    RETURN    ${value}
Enter Password
    [Documentation]    Enter password in the password field
    [Arguments]    ${password}
    Wait Until Element Is Visible    ${LOGIN_PASSWORD_FIELD}    timeout=10s
    Clear Element Text    ${LOGIN_PASSWORD_FIELD}
    Input Text    ${LOGIN_PASSWORD_FIELD}    ${password}
    ${masked_password}=    Evaluate    '*' * len('${password}')
    Log    Entered password: ${masked_password}
Verify Password Field Is Masked
    [Documentation]    Verify password field accepts input and masks characters
    ${field_type}=    Get Element Attribute    ${LOGIN_PASSWORD_FIELD}    type
    Should Be Equal    ${field_type}    password
    Log    ✓ Verified: Password field accepts input and masks characters
Click Login Button
    [Documentation]    Click the Login button
    Wait Until Element Is Visible    ${LOGIN_BUTTON}    timeout=10s
    Click Element    ${LOGIN_BUTTON}
    Log    Clicked Login button
Verify Login Button Is Clickable
    [Documentation]    Verify login button is clickable
    Element Should Be Enabled    ${LOGIN_BUTTON}
    Element Should Be Visible    ${LOGIN_BUTTON}
    Log    ✓ Verified: Login button is clickable
Perform Login
    [Documentation]    Perform complete login action
    [Arguments]    ${username}    ${password}
    Enter Username    ${username}
    Enter Password    ${password}
    Click Login Button
    Log    Login action completed