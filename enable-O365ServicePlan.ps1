# enable-O365ServicePlan.ps1
#
# by Ken Hoover <ken.hoover@yale.edu> for Yale University
# March 2017

# Turns "on" the service plan listed in $targetServicePlan for $userPrincipalName if it's currently disabled.

# Assumes that we are already connected to the target O365 environment (via connect-msolservice).

# Note this script can only handle cases where the target user has a single license assigned.  It will abort if the target user has
# more than one license.

# Provide the UPN of the user to target and the service plan you want to disable on the command line

Param ( [Parameter(Mandatory=$true)][String]$userPrincipalName,
        [Parameter(Mandatory=$true)][String]$targetServicePlan
 )

write-verbose "Attempting to enable $targetServicePlan for $userprincipalname..."

# Retrieve the user obejct from AAD for the target user
if ($userObject = get-msoluser -UserPrincipalName $userPrincipalName -ErrorAction SilentlyContinue) {
    write-verbose "Retrieved user object for $userPrincipalName"
} else {
    write-warning "$userPrincipalName not found in AzureAD"
    exit
}

# Verify that the user currently has a license
if ($userObject.islicensed -eq $false) {
    write-warning "$userPrincipalName must have a license assigned before this script can be used."
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

$baseLicense = $userObject.licenses.accountSkuId  # the user's "base" license

# Confirm that the service plan we're targeting is covered by the user's current license
if ($userObject.licenses.servicestatus | where { $_.serviceplan.servicename.equals($targetServicePlan) })
{
    write-verbose "Confirmed that $userprincipalname has $targetServicePlan as part of their license."
} else {
    write-warning "The base license for $userPrincipalName ($baseLicense) does not include $targetServicePlan."
    exit
}

# Verify that the target service plan is currently disabled for this user
if (($userObject.licenses.servicestatus | where { $_.serviceplan.servicename.equals($targetServicePlan) }).provisioningstatus -eq "Disabled" )
{
    write-verbose "Confirmed that $userprincipalname has $targetServicePlan disabled in their license."
} else {
    write-warning "$targetServicePlan is already enabled for $userPrincipalName."
    exit
}

# Now that we've finished checkihng things we can get to work.

# Build a hashtable with the service plan names and statuses for this person's license
write-verbose "Building hash table of service plans and current statuses..."
$plans = @{}
$userObject.licenses.servicestatus | % { $plans.add($_.serviceplan.servicename, $_.provisioningstatus) }


# Turn the target service plan "on" in the hash table by setting its value to "Success".
# This doesn't do anything to the user (yet)...
if ($plans.get_item($targetServicePlan) -eq "Disabled") 
{
    write-verbose "Setting value for service plan $targetServicePlan to `"Success`" in hash table"
    $plans.set_item($targetServicePlan,"Success")
}  else {
    write-warning "Service plan $targetServicePlan not found for this user."
    exit
}


# Re-generate the list of disabled plans to be used with the new-MsolLicenseOptions cmdlet by going thru the hash table and
# adding all plans which are set to "Disabled" to the $disabledPlans string
$disabledPlans = @()  # this is a list (array of strings) variable
$plans.Keys | % { if ($plans.get_item($_) -eq "Disabled") { $disabledPlans += $_ } }

# Create the new licenseOptions object if there is anything to disable
if ($disabledPlans) {
    write-verbose "Creating new msolLicenseOptions object for $userPrincipalName"
    $licenseOptions = new-msolLicenseOptions -AccountSkuId $baseLicense -DisabledPlans $disabledPlans -verbose
} else {
    write-verbose "NOTE: All service plans will now be enabled for this user."
    $licenseOptions = new-msolLicenseOptions -AccountSkuId $baseLicense
}

# Update the user's license to reflect the new set of enabled/disabled services
get-msoluser -UserPrincipalName $userObject.userprincipalname | Set-MsolUserLicense -LicenseOptions $licenseOptions
start-sleep -seconds 1

# Verify that the provisioningStatus on the target service plan is no longer "Disabled".
# Note that the status could be something like "PendingInput" so we will consider anything that's not "Disabled" as OK
if (((get-msoluser -UserPrincipalName $userPrincipalName).licenses.servicestatus | where { $_.serviceplan.servicename.equals($targetServicePlan) }).provisioningstatus -ne "Disabled" )
{
    write-host "$targetServicePlan is no longer disabled for $userprincipalname." -ForegroundColor green
} else {
    write-warning "$targetServicePlan is still disabled for $userPrincipalName.  Something went wrong?"
}

# Done.