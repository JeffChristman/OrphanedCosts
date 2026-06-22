
# Step 1: Log in to Azure commercial environment
# This command will prompt you to enter your Azure credentials
Connect-AzAccount 

# Step 2: Install the AzureResourceInventory module for the current user
# This command installs the AzureResourceInventory module which is necessary for invoking inventory commands
Install-Module -Name AzureResourceInventory -Scope CurrentUser

# Step 3: Login to Azure Government Cloud
# This command logs you into the Azure Government Cloud environment with your tenant ID
Connect-AzAccount -EnvironmentName AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com'

# Logout
# Disconnect-AzAccount

# Step 4: Invoke the Azure Resource Inventory for Azure Government Cloud
# The following command gathers resource inventory information for the specified Azure Government environment and tenant
Invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com'

# Step 5: Set the subscription for the Azure Resource Inventory
# Replace '00dba769-e94d-4697-937a-ebd57800d4be' with your actual subscription ID
Invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com' -SubscriptionID 00dba769-e94d-4697-937a-ebd57800d4be -SecurityCenter

# Step 6: Invoke the Azure Resource Inventory with Security Center data
# You can run the command multiple times as needed to gather various sets of data
Invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com' -SubscriptionID 00dba769-e94d-4697-937a-ebd57800d4be -SecurityCenter
Invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com' -SubscriptionID 00dba769-e94d-4697-937a-ebd57800d4be -SecurityCenter
#####################
# PowerShell 7.4.6

# Step 1: Install Azure module
# This command installs the Azure module, allowing you to manage Azure resources through PowerShell
Install-Module -Name AZ -AllowClobber -Scope CurrentUser

# You will see a prompt indicating that you are installing modules from an untrusted repository
# Type 'Y' or 'A' to proceed with the installation

# Step 2: Log in to Azure account
# This command authenticates you with Azure
Connect-AzAccount

# After running the command above, you will need to select an account to log in with

# Step 3: Select the subscription you want to use
# You will be prompted to select a tenant and subscription from the available list
# Example selection:
# [1] ACS-MAP-INTERNAL  [2] CORE-MAP-INTERNAL
# Type '1' or '2' to choose the appropriate subscription.

# Install AzureResourceInventory Module
# Note: Ensure correct module name to avoid typos (e.g., 'AzureResoourceInventory' is incorrect)
Install-Module -Name AzureResourceInventory -Scope CurrentUser

# Again, you might be prompted about the untrusted repository
# Type 'A' to proceed with the installation

# Step 4: Invoke Azure Resource Inventory
# This command gathers resource inventory data and generates an Excel report
Invoke-ARI

# You'll be asked to authenticate and select an account again if required
# Warning messages can inform you about overridden settings, ignore them for now if not explicitly changing settings

# The script then extracts resources from the tenant and generates the following files:
# - An Excel report: AzureResourceInventory_Report with a timestamp
# - A Draw.io Diagram file: AzureResourceInventory_Diagram with a timestamp

# Example output path:
# C:\AzureResourceInventory\AzureResourceInventory_Report_2024-12-19_06_37.xlsx
# C:\AzureResourceInventory\AzureResourceInventory_Diagram_2024-12-19_06_37.xml