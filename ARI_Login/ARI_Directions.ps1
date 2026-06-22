

#login
connect-azaccount

install-module -Name AzureResourceInventory -Scope CurrentUser

#login to AZ GovCloud
 Connect-AzAccount -EnvironmentName AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com'

 Connect-AzAccount -EnvironmentName AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com'
 invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com'


 #Set for subscription
 invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com' -SubscriptionID 00dba769-e94d-4697-937a-ebd57800d4be -SecurityCenter

 #
 invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com' -SubscriptionID 00dba769-e94d-4697-937a-ebd57800d4be -SecurityCenter
 invoke-ARI -AzureEnvironment AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com' -SubscriptionID 00dba769-e94d-4697-937a-ebd57800d4be -SecurityCenter   