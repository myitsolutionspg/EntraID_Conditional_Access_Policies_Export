@"
# Entra ID Conditional Access Policies Export

PowerShell script to export Microsoft Entra ID Conditional Access policies
to a dynamic HTML report.

The report includes:
- A policy index with name, state and policy Id
- Expand / collapse sections per policy
- Assignments (users, groups, roles)
- Cloud apps & user actions
- Conditions, grant controls and session controls

## Requirements

- PowerShell 5.1 or 7+
- Microsoft Graph PowerShell SDK with permissions to read CA policies
  (e.g. Policy.Read.All)

Install Graph (once per machine/profile):

    Install-Module Microsoft.Graph -Scope CurrentUser

## Usage

1. Open PowerShell in this folder:

    cd C:\Dev\Git\EntraID_CA_Policies_Export

2. Run the export script:

    .\EntraID_CAPolicies_Export.ps1

3. Sign in when prompted.

4. The script will create a file in the same folder:

    EntraID_CA_Policies_YYYY-MM-DD_HHMM.html

Open that file in a browser to view the report.

## Files

- EntraID_CAPolicies_Export.ps1  
  Main export script.

- sample-output/EntraID_CA_Policies_Sample.html  
  Sample report with fake tenant data, for documentation and screenshots.

- .gitignore  
  Ignores transient exports and workspace clutter.
"@ | Set-Content -Encoding utf8 README.md
