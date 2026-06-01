*** Settings ***
Documentation    Common keywords and settings for all test cases
Library          SeleniumLibrary
Library          Collections
*** Variables ***
${BROWSER}              Chrome
${TIMEOUT}              10s
${IMPLICIT_WAIT}        10s
*** Keywords ***
Open Browser To Application
    [Documentation]    Open browser and navigate to application URL
    [Arguments]    ${url}    ${browser}=${BROWSER}
    Open Browser    ${url}    ${browser}
    Maximize Browser Window
    Set Selenium Implicit Wait    ${IMPLICIT_WAIT}
    Set Selenium Timeout    ${TIMEOUT}
    Log    Browser initialized: ${browser}
Close Browser Session
    [Documentation]    Close browser and cleanup
    Close Browser
    Log    Browser closed successfully
Log Test Step
    [Documentation]    Log test step with step number and description
    [Arguments]    ${step_number}    ${description}
    Log    ${\n}Step ${step_number}: ${description}    console=True
Log Verification
    [Documentation]    Log verification result
    [Arguments]    ${description}
    Log    ✓ Verified: ${description}    console=True