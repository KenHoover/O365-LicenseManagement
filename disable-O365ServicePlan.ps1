# disable-O365serviceplan.ps1
#
# by Ken Hoover <ken.hoover@yale.edu> for Yale University
# March 2017

# Turns "off" the service plan listed in $targetServicePlan for $userPrincipalName if it's currently enabled.

# IMPORTANT ** REMOVING SERVICE PLANS CAN RESULT IN DATA LOSS.  BE CAREFUL. **

# Assumes that we are already connected to the target O365 environment (via connect-msolservice).

# Note this script can only handle cases where the target user has a single license assigned.  It will abort if the target user has
# more than one license.

# Provide the UPN of the user to target and the service plan you want to disable on the command line

Param ( [Parameter(Mandatory=$true)][String]$userPrincipalName,
        [Parameter(Mandatory=$true)][String]$targetServicePlan
 )

write-verbose "Attempting to disable $targetServicePlan for $userprincipalname..."

# Retrieve the user object from AAD for the target user
if ($userObject = get-msoluser -UserPrincipalName $userPrincipalName -ErrorAction SilentlyContinue) {
    write-verbose "Retrieved user object for $userPrincipalName"
} else {
    write-warning "$userPrincipalName not found in AzureAD"
    exit
}

# Verify that the user currently has a license
if ($userObject.islicensed -eq $false) {
    write-warning "$userPrincipalName does not have a license.  A license must be present before this script can be used."
    exit
} else {
    write-verbose "Confirmed that $userPrincipalName is licensed."
}

# Check if the user has more than one license assigned - we can't handle that :-(
if ($userObject.licenses.count -gt 1) {
    write-warning "$userPrincipalName has multiple licenses assigned.  Manual intervention required."
    exit
} else {
    write-verbose "Confirmed that $userprincipalname only has one license assigned."
}

$baseLicense = $userObject.licenses[0].accountSkuId  # the user's "base" license

# Confirm that the service plan we're targeting is covered by the user's current license
if ($userObject.licenses.servicestatus | where { $_.serviceplan.servicename.equals($targetServicePlan) })
{
    write-verbose "Confirmed that the license for $userprincipalname includes $targetServicePlan."
} else {
    write-warning "The base license for $userPrincipalName ($baseLicense) does not include $targetServicePlan."
    exit
}

# Verify that the target service plan is currently enabled for this user
if (($userObject.licenses.servicestatus | where { $_.serviceplan.servicename.equals($targetServicePlan) }).provisioningstatus -eq "Success" )
{
    write-verbose "Confirmed that $userprincipalname has $targetServicePlan enabled."
} else {
    write-warning "$targetServicePlan is not enabled for $userPrincipalName."
    exit
}

# Now that we've finished checking  things we can get to work.

# Build a hashtable with the service plan names and statuses for this person's license
write-verbose "Building table of service plans and current statuses..."
$plans = @{}
$userObject.licenses.servicestatus | % { $plans.add($_.serviceplan.servicename, $_.provisioningstatus) }

# Turn the target service plan "off" in the hash table by setting its value to "Disabled" instead of "Success".
# This doesn't do anything to the O365 user
if ($plans.get_item($targetServicePlan) -eq "Success") 
{
    write-verbose "Changing value for service plan $targetServicePlan to `"Disabled`" in local service plan table"
    $plans.set_item($targetServicePlan,"Disabled")
}


# Re-generate the list of disabled plans to be used with the new-MsolLicenseOptions cmdlet by going thru the hash table and
# adding all plans which are set to "Disabled" to the $disabledPlans string
$disabledPlans = @()
$plans.Keys | % { if ($plans.get_item($_) -eq "Disabled") { $disabledPlans += $_ } }

if ($disabledPlans) {
    write-verbose "Creating new msolLicenseOptions object for $userPrincipalName"
    $licenseOptions = new-msolLicenseOptions -AccountSkuId $baseLicense -DisabledPlans $disabledPlans -verbose
} else {
    write-verbose "NOTE: All service plans will now be enabled for this user."
    $licenseOptions = new-msolLicenseOptions -AccountSkuId $baseLicense
}

# Update the user's license to reflect the new set of enabled/disabled services
get-msoluser -UserPrincipalName $userObject.userprincipalname | Set-MsolUserLicense -LicenseOptions $licenseOptions

# write-host "Updated service plans for $userPrincipalName ($baseLicense)" -fore Green

# Verify that the provisioningStatus on the target service plan is no longer "Disabled".
# Note that the status could be something like "PendingInput"

if (((get-msoluser -UserPrincipalName $userPrincipalName).licenses.servicestatus | where { $_.serviceplan.servicename.equals($targetServicePlan) }).provisioningstatus -eq "Disabled" )
{
    write-host "$targetServicePlan has been disabled for $userprincipalname." -ForegroundColor Green
} else {
    write-warning "$targetServicePlan is not disabled for $userPrincipalName!  Something went wrong?"
}

# Done.