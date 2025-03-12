# Domain Inventory Report Script

This PowerShell script inventories objects for given domains from Exchange Online recipients and Azure AD Security Groups. It exports the results to a CSV file, providing comprehensive insights into user mailbox configurations and group memberships without making any changes.

## Prerequisites

1. **PowerShell**: Ensure PowerShell is installed on your system.
2. **Required Modules**: The script uses `ExchangeOnlineManagement`, `MSOnline`, and `AzureAD` modules. These will be automatically installed if they are not already present.

## Instructions

1. **Edit the Script**:
   - Open the script file in a text editor.
   - Locate and update the `$Domains` variable with the domains you want to inventory. The script expects these domains to be listed in a file named `Domains.csv` in the same directory as the script.

2. **Prepare the CSV File**:
   - Ensure you have a `Domains.csv` file in the same directory as the script. This file should list the domains you wish to inventory:
     ```
     Domain
     example.com
     anotherdomain.com
     ```

3. **Run the Script**:
   - Open PowerShell as an administrator.
   - Navigate to the directory containing the script.
   - Execute the script:
     ```powershell
     .\DomainInventoryReport.ps1
     ```
   - You will be prompted to authenticate with Exchange Online and Azure AD.

4. **Check the Output**:
   - The results will be saved in a CSV file named after the processed domain, e.g., `example_com_Report-YYYYMMDD-HHmm.csv`.

## Troubleshooting

- **Authentication Failure**: Ensure you have administrative credentials for Exchange Online and Azure AD.
- **CSV File Not Found**: Check that `Domains.csv` is present in the correct location and properly formatted.
- **Module Installation Issues**: If automatic module installation fails, you may manually install them using:
  ```powershell
  Install-Module -Name ExchangeOnlineManagement -Force
  Install-Module -Name MSOnline -Force
  Install-Module -Name AzureAD -Force
  
Additional Notes
This script is intended for use by administrators needing to audit and maintain accurate records of recipient configurations and group attributes across both Exchange Online and Azure AD.


### Key Points:
- **Simplified Instructions**: The README simplifies the script usage instructions, making it clear and easy to follow.
- **Troubleshooting Section**: Includes common issues and their solutions to help users solve problems independently.
- **Dynamic Output File Naming**: Explains how the output files are named based on the domain being processed, which helps in managing multiple domain reports.
