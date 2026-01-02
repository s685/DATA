# How to Set Environment Variables on Windows

## Required Environment Variables

The following environment variables are **required** for the script to work:

- `SNOWFLAKE_ACCOUNT` - Your Snowflake account (e.g., "xy12345.us-east-1")
- `SNOWFLAKE_WAREHOUSE` - Your Snowflake warehouse name
- `SNOWFLAKE_DATABASE` - Your Snowflake database name
- `SNOWFLAKE_SCHEMA` - Your Snowflake schema name
- `SNOWFLAKE_AUTHENTICATOR` - Authentication method: "externalbrowser" (SSO) or "snowflake" (username/password)

**Optional but may be required:**
- `SNOWFLAKE_USER` - Your Snowflake username
  - **Required** if `SNOWFLAKE_AUTHENTICATOR="snowflake"` (username/password)
  - **May be required** for some Snowflake configurations even with `externalbrowser` (SSO)
  - If you get "user is empty" error, set this variable

**Optional** (only if using username/password authentication):
- `SNOWFLAKE_PASSWORD` - Your Snowflake password (required if `SNOWFLAKE_AUTHENTICATOR="snowflake"`)

---

## Method 1: PowerShell (Current Session Only)

Open PowerShell and run these commands:

### For SSO (External Browser) Authentication:
```powershell
$env:SNOWFLAKE_ACCOUNT = "xy12345.us-east-1"
$env:SNOWFLAKE_WAREHOUSE = "YOUR_WAREHOUSE"
$env:SNOWFLAKE_DATABASE = "YOUR_DATABASE"
$env:SNOWFLAKE_SCHEMA = "YOUR_SCHEMA"
$env:SNOWFLAKE_AUTHENTICATOR = "externalbrowser"
```

### For Username/Password Authentication:
```powershell
$env:SNOWFLAKE_ACCOUNT = "xy12345.us-east-1"
$env:SNOWFLAKE_WAREHOUSE = "YOUR_WAREHOUSE"
$env:SNOWFLAKE_DATABASE = "YOUR_DATABASE"
$env:SNOWFLAKE_SCHEMA = "YOUR_SCHEMA"
$env:SNOWFLAKE_AUTHENTICATOR = "snowflake"
$env:SNOWFLAKE_USER = "your_username"
$env:SNOWFLAKE_PASSWORD = "your_password"
```

**Note:** These settings are only valid for the current PowerShell session. Close the window and they're gone.

---

## Method 2: Command Prompt (Current Session Only)

Open Command Prompt (cmd) and run these commands:

### For SSO (External Browser) Authentication:
```cmd
set SNOWFLAKE_ACCOUNT=xy12345.us-east-1
set SNOWFLAKE_WAREHOUSE=YOUR_WAREHOUSE
set SNOWFLAKE_DATABASE=YOUR_DATABASE
set SNOWFLAKE_SCHEMA=YOUR_SCHEMA
set SNOWFLAKE_AUTHENTICATOR=externalbrowser
```

### For Username/Password Authentication:
```cmd
set SNOWFLAKE_ACCOUNT=xy12345.us-east-1
set SNOWFLAKE_WAREHOUSE=YOUR_WAREHOUSE
set SNOWFLAKE_DATABASE=YOUR_DATABASE
set SNOWFLAKE_SCHEMA=YOUR_SCHEMA
set SNOWFLAKE_AUTHENTICATOR=snowflake
set SNOWFLAKE_USER=your_username
set SNOWFLAKE_PASSWORD=your_password
```

**Note:** These settings are only valid for the current Command Prompt session.

---

## Method 3: Set Permanently (System-Wide)

### Using PowerShell (Run as Administrator):

```powershell
# For SSO Authentication
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_ACCOUNT", "xy12345.us-east-1", "User")
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_WAREHOUSE", "YOUR_WAREHOUSE", "User")
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_DATABASE", "YOUR_DATABASE", "User")
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_SCHEMA", "YOUR_SCHEMA", "User")
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_AUTHENTICATOR", "externalbrowser", "User")

# For Username/Password Authentication (add these too):
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_AUTHENTICATOR", "snowflake", "User")
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_USER", "your_username", "User")
[System.Environment]::SetEnvironmentVariable("SNOWFLAKE_PASSWORD", "your_password", "User")
```

**Note:** After setting permanently, you may need to restart your terminal/PowerShell window for changes to take effect.

### Using Windows GUI:

1. Press `Win + R` to open Run dialog
2. Type `sysdm.cpl` and press Enter
3. Go to the **Advanced** tab
4. Click **Environment Variables**
5. Under **User variables** (or **System variables**), click **New**
6. Add each variable:
   - Variable name: `SNOWFLAKE_ACCOUNT`
   - Variable value: `xy12345.us-east-1`
   - Click **OK**
7. Repeat for all required variables
8. Click **OK** to save
9. **Restart your terminal/PowerShell** for changes to take effect

---

## Method 4: Create a Batch Script (Easiest for Repeated Use)

Create a file called `set_env.bat` with the following content:

```batch
@echo off
set SNOWFLAKE_ACCOUNT=xy12345.us-east-1
set SNOWFLAKE_WAREHOUSE=YOUR_WAREHOUSE
set SNOWFLAKE_DATABASE=YOUR_DATABASE
set SNOWFLAKE_SCHEMA=YOUR_SCHEMA
set SNOWFLAKE_AUTHENTICATOR=externalbrowser
REM Uncomment these if using username/password:
REM set SNOWFLAKE_USER=your_username
REM set SNOWFLAKE_PASSWORD=your_password
echo Environment variables set!
```

Then run it before running your script:
```cmd
set_env.bat
python excel_report_generator.py --config config.yaml --output report.xlsx --report-start-dt 2024-01-01 --report-end-dt 2024-12-31
```

---

## Method 5: Create a PowerShell Script

Create a file called `set_env.ps1` with the following content:

```powershell
$env:SNOWFLAKE_ACCOUNT = "xy12345.us-east-1"
$env:SNOWFLAKE_WAREHOUSE = "YOUR_WAREHOUSE"
$env:SNOWFLAKE_DATABASE = "YOUR_DATABASE"
$env:SNOWFLAKE_SCHEMA = "YOUR_SCHEMA"
$env:SNOWFLAKE_AUTHENTICATOR = "externalbrowser"
# Uncomment these if using username/password:
# $env:SNOWFLAKE_USER = "your_username"
# $env:SNOWFLAKE_PASSWORD = "your_password"
Write-Host "Environment variables set!"
```

Then run it:
```powershell
. .\set_env.ps1
python excel_report_generator.py --config config.yaml --output report.xlsx --report-start-dt 2024-01-01 --report-end-dt 2024-12-31
```

---

## Verify Environment Variables Are Set

### In PowerShell:
```powershell
Get-ChildItem Env: | Where-Object Name -like "SNOWFLAKE_*"
```

### In Command Prompt:
```cmd
set | findstr SNOWFLAKE
```

---

## Quick Test

After setting the variables, test if they're accessible:

### PowerShell:
```powershell
echo $env:SNOWFLAKE_ACCOUNT
echo $env:SNOWFLAKE_DATABASE
```

### Command Prompt:
```cmd
echo %SNOWFLAKE_ACCOUNT%
echo %SNOWFLAKE_DATABASE%
```

---

## Troubleshooting

1. **Variables not found after setting**: 
   - If set temporarily, make sure you're in the same terminal session
   - If set permanently, restart your terminal/PowerShell window

2. **Permission denied**:
   - For permanent system-wide variables, run PowerShell as Administrator

3. **Still getting errors**:
   - Verify variable names are spelled correctly (case-sensitive)
   - Check that values don't have extra spaces
   - Make sure you're using the same terminal session where you set them

---

## Alternative: Use config.yaml Instead

If you prefer not to use environment variables, you can set these values directly in `config.yaml`:

```yaml
snowflake:
  account: xy12345.us-east-1
  warehouse: YOUR_WAREHOUSE
  database: YOUR_DATABASE
  schema: YOUR_SCHEMA
  authenticator: externalbrowser  # or "snowflake"
  user: your_username  # Only if authenticator is "snowflake"
  password: your_password  # Only if authenticator is "snowflake"
```

**Note:** Environment variables take precedence over config.yaml values.

