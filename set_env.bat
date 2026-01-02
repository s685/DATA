@echo off
REM ============================================
REM Snowflake Environment Variables Setup
REM ============================================
REM Edit the values below with your Snowflake credentials
REM ============================================

REM Required: Your Snowflake account (e.g., "xy12345.us-east-1")
set SNOWFLAKE_ACCOUNT=xy12345.us-east-1

REM Required: Your Snowflake warehouse
set SNOWFLAKE_WAREHOUSE=YOUR_WAREHOUSE

REM Required: Your Snowflake database
set SNOWFLAKE_DATABASE=YOUR_DATABASE

REM Required: Your Snowflake schema
set SNOWFLAKE_SCHEMA=YOUR_SCHEMA

REM Required: Authentication method - "externalbrowser" for SSO or "snowflake" for username/password
set SNOWFLAKE_AUTHENTICATOR=externalbrowser

REM Optional: Only needed if SNOWFLAKE_AUTHENTICATOR="snowflake"
REM set SNOWFLAKE_USER=your_username
REM set SNOWFLAKE_PASSWORD=your_password

echo.
echo ============================================
echo Environment variables set!
echo ============================================
echo SNOWFLAKE_ACCOUNT=%SNOWFLAKE_ACCOUNT%
echo SNOWFLAKE_WAREHOUSE=%SNOWFLAKE_WAREHOUSE%
echo SNOWFLAKE_DATABASE=%SNOWFLAKE_DATABASE%
echo SNOWFLAKE_SCHEMA=%SNOWFLAKE_SCHEMA%
echo SNOWFLAKE_AUTHENTICATOR=%SNOWFLAKE_AUTHENTICATOR%
echo ============================================
echo.
echo You can now run your script in this terminal window.
echo.

