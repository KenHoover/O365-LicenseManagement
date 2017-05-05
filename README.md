# O365-LicenseManagement

Powershell scripts that deal with managing license for users in Office365

## enable-O365ServicePlan.ps1 and disable-O365ServicePlan.ps1
These scripts demonstrate changing the status of a single service plan within a user's license in Office365.

IMPORTANT:  Removing service plans from users can result in data loss.  Be careful!

### Assumptions

These scripts assume that the user has already authenticated using connect-msolservice.

### Parameters

#### userPrincipalName

The UPN of the user to target in Office365

#### targetServicePlan

The identifier of the service plan that is to have its status changed.  Note this is not the same thing as the "Name" of the service plan.  To see the list of service plans within a license, use the command (get-msolaccountsku | where { $_.accountSkuID.equals(<AccountSkuID>) }).servicestatus
