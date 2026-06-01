@echo off
REM Batch script to run Robot Framework tests on Windows
echo ========================================
echo Swag Labs Robot Framework Test Execution
echo ========================================
echo.
REM Set environment variables (optional - can override variables.py)
REM set BASE_URL=https://www.saucedemo.com/
REM set USERNAME=standard_user
REM set PASSWORD=secret_sauce
echo Running Test: TC_AIGPMM_38_001
echo.
REM Run Robot Framework tests
robot --outputdir Results --variable BASE_URL:https://www.saucedemo.com/ --variable USERNAME:standard_user --variable PASSWORD:secret_sauce --variable FIRST_NAME:Sarah --variable LAST_NAME:Johnson --variable ZIP_CODE:78701 --variable PRODUCT_NAME:"Sauce Labs Backpack" Tests/TC_AIGPMM_38_001_Complete_Checkout_Flow_Single_Product.robot
echo.
echo ========================================
echo Test Execution Completed
echo Results saved in Results/ directory
echo ========================================
pause