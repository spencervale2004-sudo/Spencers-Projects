# Powershell to be run as system on the selected users machine #
# Script will need an Entra app created on the Tenant to work correctly #
# Client secret will technically need to be exposed to the machine of the user which is why it only has read only permissions #

  # # # # - - - VARIABLES that need to be set - - - # # # #
  # $URL          | Url to the template file (or files if using $tenantname)
  # $TenantId     | Tenant ID of current logged in users Entra/365 tenant
  # $ClientId     | Client ID of the Entra app on the users tenant
  # $ClientSecret | Client Secret of the Entra app

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Confirm:$false
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -Confirm:$false 
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted


# ----Function---- Checks if a module is installed and then installs if missing #
function Module-Check {
      param (
        [string]$ModuleName
      )
      
      # Check if the module is installed
      $module = Get-InstalledModule -Name "$ModuleName" -ErrorAction SilentlyContinue
      
      # If the module is not installed (count is 0), then install it
      if ($module.Count -eq 0) {
          Write-Host "$ModuleName is not installed. Installing now..."
          Install-Module -Name $ModuleName -Force -AllowClobber -Confirm:$false
      } else {
          Write-Host "$ModuleName is already installed."
      }   
}


# Use function above to download needed graph modules if not already installed
Module-Check -ModuleName "Microsoft.Graph.Users"
Module-Check -ModuleName "Microsoft.Graph.Authentication"
Module-Check -ModuleName "Microsoft.Graph.Identity.DirectoryManagement"

# Get logged in users name
$User = (Get-CimInstance -ClassName Win32_ComputerSystem).Username
# If no user currently logged in cancel script
if (!$User) {
  Write-output "Cancelling script: No logged in user"
  exit 0
}
# Splitting username from the domain name. Before: 'Example\Username'
$User = $User.split("\")[1]

#####-- Details of created Entra Application --#####
    $Tenantid = "$Tenantid"
    $ClientId = "$ClientId" # App will only need User.Read.All and Organization.Read.All
    $ClientSecret = "$ClientSecret" # If doing multiple Tenants you will need to decide how the powershell will get this value (Along with Tenant ID)
    
### URL where template files are to be downloaded from ###   
    $path = "$URL"
    
$SecureClientSecret = ConvertTo-SecureString -String $ClientSecret -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential($ClientId, $SecureClientSecret)

# Try to connect to Tenant's graph with Application Client Secret. If failed assumes no app is present and writes output for admin consent link.
try {
  Connect-MgGraph -ClientSecretCredential $Credential -TenantId $TenantId -NoWelcome
} catch {
  write-output "Connection to graph has failed. If there is no App on this tenant please login using the following url to create the app (Remember to save the client secret) https://login.microsoftonline.com/$TenantId/adminconsent?client_id=784017f8-db47-4764-81ab-a67c74492be0"
  exit 1
}

# Get Tenant name and replace spaces and full stops with hyphens (for folder structure).
$tenant = get-mgorganization
write-output "Connected to tenant: $($tenant.displayName)"
$tenantname = $tenant.displayName -replace " ", "-" -replace "\.", ""
# Get 365 User based on the local user name found at the start
$365user = Get-MgUser -ConsistencyLevel eventual -Count userCount -Filter "startsWith(UserPrincipalName, '$($User)')" -Top 1

### Check if template file already exists at C:\Temp. If not downloads template file ###

$Item = get-item "C:\Temp\Company Default.htm" -ErrorAction SilentlyContinue
if (!$item) {
    Invoke-WebRequest -Uri "$path" -OutFile "C:\Temp\Company Default.htm" 
}

### ----Variables---- Set 365 Users properties to variables ###
$filePath = "C:\Temp\Company Default.htm"
$Name = $365user.givenName
$Surname = $365user.Surname
$Title = $365user.JobTitle
$Phone = $365user.BusinessPhones
$Mobile = $365user.MobilePhone
$UserPrincipalName = $365user.UserPrincipalName
$OfficeLocation = $365user.OfficeLocation
$Mail = $365user.Mail
$StreetAddress = $365user.StreetAddress

### Put variables in table for replacement (Left is text being replaced, right is the text replacing) ###
# Old text | New Text #
$Properties = @{
"Username" = $Name
"UserSurname" = $Surname
"UserTitle" = $Title
"UserPhone" = $Phone
"UserMobile" = $Mobile
"UserPrincipalName" = $UserPrincipalName
"UserMail" = $Mail
"UserLocation" = $OfficeLocation
"UserStreetAddress" = $StreetAddress
}
# Get content of the template signature #
$HTML = Get-Content $filePath 
# Replace template variables (Left in the table) in the template file with the new values (Right) #
foreach ($Value in $Properties.Keys) {
    $HTML = $HTML -replace $Value, $Properties[$Value]
}
# Display properties #
Write-Output "`n##### Entra Properties Below #####"
$Properties

####------------ if $Test variable is equal to 'no' then set registry keys to make new htm file the default signature for new emails and replys -----------####
if ($Test -like "no") {
# Creates new file if doesn't exist or overwrites (put Hyphen in name so it doesn't overwrite the old existing ones)
if(!(Test-Path "C:\Users\$($User)\AppData\Roaming\Microsoft\Signatures")){
    New-Item -Path "C:\Users\$($User)\AppData\Roaming\Microsoft\Signatures" -ItemType Directory
}
Set-Content -Path "C:\Users\$($User)\AppData\Roaming\Microsoft\Signatures\Company-Default.htm" -Value $HTML


### Change Registry Keys to make new htm file the default ###

 $userSID = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.LocalPath -like "C:\Users\$($User)"} | Select-Object -ExpandProperty SID
    
    $Key = "registry::HKEY_USERS\$usersid\Software\Microsoft\Office\16.0\Common\MailSettings"

    if (!(Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }

    Set-ItemProperty -Path "$key" -Name "NewSignature" -Type String -Value "Company-Default"

    Set-ItemProperty -Path "$key" -Name "ReplySignature" -Type String -Value "Company-Default"

####---- if $Test variable is equal to anything other then 'no' then set only make htm file in a new folder in temp named '(Username) Signature Test' e.g. C:\Temp\(Username) Signature Test\Company-Default.htm  ----####
} else {

    if(!(Test-Path "C:\Temp\$($User) Signature Test")){
        New-Item -Path "C:\Temp\$($User) Signature Test" -ItemType Directory
    }
    Set-Content -Path "C:\Temp\$($User) Signature Test\Company-Default.htm" -Value $HTML

}

Disconnect-MgGraph
    
