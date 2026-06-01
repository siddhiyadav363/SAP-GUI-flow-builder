*** Settings ***
Documentation    Test Case: TC_AIGPMM_38_001
...              Test complete checkout flow for a single product from cart to finish page
...              Story: AIGPMM-38
...              Priority: High
...              
...              Objective: Validate that a user can complete the entire checkout process 
...              with valid information for a single product purchase
Library          SeleniumLibrary
Resource         ../Resources/Common.robot
Resource         ../Resources/PageObjects/LoginPage.robot
Resource         ../Resources/PageObjects/ProductsPage.robot
Resource         ../Resources/PageObjects/CartPage.robot
Resource         ../Resources/PageObjects/CheckoutInformationPage.robot
Resource         ../Resources/PageObjects/CheckoutOverviewPage.robot
Resource         ../Resources/PageObjects/CheckoutCompletePage.robot
Suite Setup      Log    === Starting Test Suite: TC_AIGPMM_38_001 - Complete Checkout Flow - Single Product ===    console=True
Suite Teardown   Log    === Completed Test Suite: TC_AIGPMM_38_001 ===    console=True
*** Variables ***
# Test data - Replace these with actual values during execution
${BASE_URL}          https://www.saucedemo.com/
${USERNAME}          standard_user
${PASSWORD}          secret_sauce
${FIRST_NAME}        Sarah
${LAST_NAME}         Johnson
${ZIP_CODE}          78701
${PRODUCT_NAME}      Sauce Labs Backpack
*** Test Cases ***
TC_AIGPMM_38_001 Complete Checkout Flow Single Product
    [Documentation]    Verify successful checkout completion for a single product from cart to finish page
    ...                Test Type: Functional
    ...                Priority: High
    [Tags]    Functional    High    Checkout    SingleProduct    AIGPMM-38
    # Test Setup
    Log    === Starting Test: TC_AIGPMM_38_001 - Complete Checkout Flow - Single Product ===    console=True
    Open Browser To Application    ${BASE_URL}    Chrome
    # Step 1: Navigate to the application URL
    Log Test Step    1    Navigate to application URL
    LoginPage.Navigate To Login Page    ${BASE_URL}
    # Expected Result 1: Login page loads successfully
    LoginPage.Verify Login Page Is Displayed
    # Step 2: Enter username in the Username field
    Log Test Step    2    Enter username in the Username field
    LoginPage.Enter Username    ${USERNAME}
    # Expected Result 2: Username field accepts input
    LoginPage.Verify Username Field Accepts Input
    ${username_value}=    LoginPage.Get Username Value
    Should Be Equal    ${username_value}    ${USERNAME}
    Log Verification    Username value matches input: ${username_value}
    # Step 3: Enter password in the Password field
    Log Test Step    3    Enter password in the Password field
    LoginPage.Enter Password    ${PASSWORD}
    # Expected Result 3: Password field accepts input and masks characters
    LoginPage.Verify Password Field Is Masked
    # Step 4: Click 'Login' button
    Log Test Step    4    Click Login button
    # Expected Result 4: Login button is clickable
    LoginPage.Verify Login Button Is Clickable
    LoginPage.Click Login Button
    # Step 5: Wait for Products page to load
    Log Test Step    5    Wait for Products page to load
    ProductsPage.Wait For Products Page To Load
    # Expected Result 5: Products page loads with product listings
    ProductsPage.Verify Products Page Is Displayed
    # Step 6: Click 'Add to cart' button for Sauce Labs Backpack
    Log Test Step    6    Click 'Add to cart' button for Sauce Labs Backpack
    ProductsPage.Click Add To Cart For Sauce Labs Backpack
    # Expected Result 6: 'Add to cart' button changes to 'Remove' after clicking
    ProductsPage.Verify Remove Button Is Displayed
    # Step 7: Click on 'Cart' icon in the top right corner
    Log Test Step    7    Click on Cart icon
    # Expected Result 7: Cart icon shows badge with '1'
    ProductsPage.Verify Cart Badge Count    1
    ProductsPage.Click Cart Icon
    # Step 8: Verify product appears in cart with quantity 1
    Log Test Step    8    Verify product appears in cart
    CartPage.Wait For Cart Page To Load
    # Expected Result 8: Cart page displays with correct product name and quantity
    CartPage.Verify Product With Quantity In Cart    ${PRODUCT_NAME}
    # Step 9: Click 'Checkout' button
    Log Test Step    9    Click Checkout button
    # Expected Result 9: Checkout button is visible and clickable
    CartPage.Verify Checkout Button Is Visible And Clickable
    CartPage.Click Checkout Button
    # Step 10: Wait for 'Checkout: Your Information' page to load
    Log Test Step    10    Wait for 'Checkout: Your Information' page to load
    CheckoutInformationPage.Wait For Checkout Information Page To Load
    # Expected Result 10: 'Checkout: Your Information' page displays
    CheckoutInformationPage.Verify Checkout Information Page Is Displayed
    # Step 11: Enter first name in First Name field
    Log Test Step    11    Enter first name in First Name field
    CheckoutInformationPage.Enter First Name    ${FIRST_NAME}
    # Expected Result 11: First Name field accepts alphabetic input
    CheckoutInformationPage.Verify First Name Field Is Enabled
    # Step 12: Enter last name in Last Name field
    Log Test Step    12    Enter last name in Last Name field
    CheckoutInformationPage.Enter Last Name    ${LAST_NAME}
    # Expected Result 12: Last Name field accepts alphabetic input
    CheckoutInformationPage.Verify Last Name Field Is Enabled
    # Step 13: Enter zip code in Zip/Postal Code field
    Log Test Step    13    Enter zip code in Zip/Postal Code field
    CheckoutInformationPage.Enter Zip Postal Code    ${ZIP_CODE}
    # Expected Result 13: Zip/Postal Code field accepts numeric input
    CheckoutInformationPage.Verify Zip Postal Code Field Is Enabled
    # Step 14: Click 'Continue' button
    Log Test Step    14    Click Continue button
    CheckoutInformationPage.Click Continue Button
    # Step 15: Wait for 'Checkout: Overview' page to load
    Log Test Step    15    Wait for 'Checkout: Overview' page to load
    CheckoutOverviewPage.Wait For Checkout Overview Page To Load
    # Expected Result 15: 'Checkout: Overview' page displays with correct header
    CheckoutOverviewPage.Verify Checkout Overview Header Is Displayed
    # Step 16: Verify product details, payment information, and shipping information
    Log Test Step    16    Verify product details, payment information, and shipping information
    # Expected Result 16: Product table shows quantity and description correctly
    CheckoutOverviewPage.Verify Product Table Displayed Correctly    ${PRODUCT_NAME}
    # Expected Result 17: Payment Information and Shipping Information sections display
    CheckoutOverviewPage.Verify Payment And Shipping Info Displayed
    # Step 17: Verify Item Total, Tax, and Total amounts
    Log Test Step    17    Verify Item Total, Tax, and Total amounts are calculated and displayed
    # Expected Result 18: Item Total, Tax, and Total are displayed
    CheckoutOverviewPage.Verify Price Calculations Displayed
    ${item_total}=    CheckoutOverviewPage.Get Item Total
    ${tax}=    CheckoutOverviewPage.Get Tax
    ${total}=    CheckoutOverviewPage.Get Total
    Should Not Be Empty    ${item_total}
    Should Not Be Empty    ${tax}
    Should Not Be Empty    ${total}
    Log    Item Total: ${item_total}    console=True
    Log    Tax: ${tax}    console=True
    Log    Total: ${total}    console=True
    # Step 18: Click 'Finish' button
    Log Test Step    18    Click Finish button
    CheckoutOverviewPage.Click Finish Button
    # Step 19: Wait for Finish page to load
    Log Test Step    19    Wait for Finish page to load
    CheckoutCompletePage.Wait For Finish Page To Load
    # Expected Result 19: Finish page loads successfully
    CheckoutCompletePage.Verify Finish Page Is Displayed
    # Step 20: Verify 'Thank you for your order!' message displays
    Log Test Step    20    Verify 'Thank you for your order!' message displays
    # Expected Result 20: Success message and Pony Express Sauce Labs logo display
    CheckoutCompletePage.Verify Thank You Message Is Displayed
    CheckoutCompletePage.Verify Pony Express Logo Is Displayed
    Log    === Test Completed Successfully ===    console=True
    Log    Complete checkout flow for single product verified successfully    console=True
    # Test Teardown
    [Teardown]    Close Browser Session