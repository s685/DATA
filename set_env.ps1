# ============================================
# Snowflake Environment Variables Setup
# ============================================
# Edit the values below with your Snowflake credentials
# ============================================

# Required: Your Snowflake account (e.g., "xy12345.us-east-1")
$env:SNOWFLAKE_ACCOUNT = "xy12345.us-east-1"

# Required: Your Snowflake warehouse
$env:SNOWFLAKE_WAREHOUSE = "YOUR_WAREHOUSE"

# Required: Your Snowflake database
$env:SNOWFLAKE_DATABASE = "YOUR_DATABASE"

# Required: Your Snowflake schema
$env:SNOWFLAKE_SCHEMA = "YOUR_SCHEMA"

# Required: Authentication method - "externalbrowser" for SSO or "snowflake" for username/password
$env:SNOWFLAKE_AUTHENTICATOR = "externalbrowser"

# Optional: Only needed if SNOWFLAKE_AUTHENTICATOR="snowflake"
# $env:SNOWFLAKE_USER = "your_username"
# $env:SNOWFLAKE_PASSWORD = "your_password"

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "Environment variables set!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "SNOWFLAKE_ACCOUNT=$env:SNOWFLAKE_ACCOUNT"
Write-Host "SNOWFLAKE_WAREHOUSE=$env:SNOWFLAKE_WAREHOUSE"
Write-Host "SNOWFLAKE_DATABASE=$env:SNOWFLAKE_DATABASE"
Write-Host "SNOWFLAKE_SCHEMA=$env:SNOWFLAKE_SCHEMA"
Write-Host "SNOWFLAKE_AUTHENTICATOR=$env:SNOWFLAKE_AUTHENTICATOR"
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Host "You can now run your script in this PowerShell window." -ForegroundColor Yellow
Write-Host ""

