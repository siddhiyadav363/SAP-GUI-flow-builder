#!/bin/bash
# Shell script to run Robot Framework tests on Linux/Mac
echo "========================================"
echo "Swag Labs Robot Framework Test Execution"
echo "========================================"
echo ""
# Set environment variables (optional - can override variables.py)
# export BASE_URL="https://www.saucedemo.com/"
# export USERNAME="standard_user"
# export PASSWORD="secret_sauce"
echo "Running Test: TC_AIGPMM_38_001"
echo ""
# Run Robot Framework tests
robot --outputdir Results \
      --variable BASE_URL:"https://www.saucedemo.com/" \
      --variable USERNAME:"standard_user" \
      --variable PASSWORD:"secret_sauce" \
      --variable FIRST_NAME:"Sarah" \
      --variable LAST_NAME:"Johnson" \
      --variable ZIP_CODE:"78701" \
      --variable PRODUCT_NAME:"Sauce Labs Backpack" \
      Tests/TC_AIGPMM_38_001_Complete_Checkout_Flow_Single_Product.robot
echo ""
echo "========================================"
echo "Test Execution Completed"
echo "Results saved in Results/ directory"
echo "========================================"