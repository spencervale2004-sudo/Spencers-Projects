# Signature Automation Script
Script made for pushing out signatures on Entra only machines with no domain connection. The script will still work on Domain and hybrid joined devices but will need Azure AD Sync so that it gets relevent details.

### How it works
Script will use an Entra app with read only permissions to connect to graph and get the properties of the current logged in user.
Steps on how to make the Entra app is in the set up section below.

After the script has gathered the users properties it will download the template file (will recommend a github repo as the path) and replace the template variables in the template .htm file

## Set up
### Entra App
For the script to work you will need to create an Entra app by following the steps below

1. Log in to the Tenant's Entra with at least application administrator permissions
2. Navigate to and select **App registrations** on the left
3. Select **New registration** button
4. Name the app and select either **Single tenant or Multitenant** (preferably no personal accounts)
5. Register the app
6. Select **Api permissions**
7. Select **Add a permission** => **Microsoft Graph** => **Application permissions** and then find and choose permissions **User.Read.All** and **Organization.Read.All**
8. Add the permissions
9. Now if you have the correct rights on the account you have logged in with you can select '**Grant admin consent for _Tenant Name_**'

### Template .htm files

For the html of the signature the script will download the template file named 'Company Default.htm' from the path of your choosing to C:\Temp.\
From there it will read the content and replace the replacement variables outlined below in the template with the Entra values it has found.

| **Template Replacement Text**  | **Entra User property** |
| ------------------------------ | ----------------------- |
| `Username`                     | Given Name              |
| `UserSurname`                  | Surname                 |
| `UserTitle`                    | JobTitle                |
| `UserPhone`                    | BusinessPhones          | 
| `UserMobile`                   | MobilePhone             |
| `UserPrincipalName`            | UserPrincipalName       | 
| `UserMail`                     | UserPrincipalName       |
| `UserLocation`                 | OfficeLocation          | 
| `UserStreetAddress`            | StreetAddress           |

Run the code in the tab below to test replacement variables on your template.

<details>

<summary>Powershell test replacement variables</summary>

**After running the powershell a new .htm file will be created in a folder named 'Signature test' in your C:\temp that will have replaced the variables in your template with Test'_property_' (Can see in code below).**

```ruby
# Set the file path variable to path of your template file #
$filePath = "C:\Temp\Company Default.htm"

### Left is text being replaced, right is the text replacing ###
# Old text | New Text #
$Properties = @{
"Username" = "TestName"
"UserSurname" = "TestSurname"
"UserTitle" = "TestTitle"
"UserPhone" = "TestPhone"
"UserMobile" = "TestMobile"
"UserPrincipalName" = "TestUserPrincipalName"
"UserMail" = "TestMail"
"UserLocation" = "TestOfficeLocation"
"UserStreetAddress" = "TestStreetAddress"
}
# Get content of the template signature #
$HTML = Get-Content $filePath

# Replace template variables (Left in the table) in the template file with the new values (Right) #
foreach ($Value in $Properties.Keys) {
    $HTML = $HTML -replace $Value, $Properties[$Value]
}
# Create htm filepath if not already made
    if(!(Test-Path "C:\Temp\Signature Test")){
        New-Item -Path "C:\Temp\Signature Test" -ItemType Directory
    }
# Create or overwrite new .htm file
Set-Content -Path "C:\Temp\Signature Test\Company-Default.htm" -Value $HTML
```

</details>

### Folder Structure

It is up to the person working with the script to figure out where the template files within the script will be downloaded from. The only important thing is that the file is a .htm file and that it is able to be downloaded through a Powershell Invoke-WebRequest. 

**You can test your url to the .htm file by editing and running the code below in Powershell run as administrator.**

```ruby
# download file to C:\Temp as Company Default.htm
Invoke-WebRequest -Uri "http://Url/to/signaturefiles/Company Default.htm" -OutFile "C:\Temp\Company Default.htm" 
```

<details>

<summary>Tenant name variable for MSP Environments</summary>

### Tenant name variable for Template selection

For msp environments the $tenantname variable is available. The $tenantname variable will replace all spaces in the gathered tenant name with hyphens and will remove full stops to homogenise the naming of the folders (Example folder could be named 'Example Template-Folder.')

You can run the code in the tab below to test the tenant name variable with your folder structure/url.

<details>

<summary>Code for testing</summary>

### Run this first section if you don't have the modules installed already

```ruby
# Download modules if not already installed
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Confirm:$false
Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -Confirm:$false 
Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted

function Module-Check {

param (
[string]$ModuleName
)

# Check if the module is installed
$module = Get-InstalledModule -Name "$ModuleName" -ErrorAction SilentlyContinue

# If the module is not installed (count is 0), then install it
if ($module.Count -eq 0) {
    Write-Host "$ModuleName is not installed. Installing now..."
    Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser -Confirm:$false
} else {
    Write-Host "$ModuleName is already installed."
}

}

Module-Check -ModuleName "Microsoft.Graph.Users"
Module-Check -ModuleName "Microsoft.Graph.Authentication"
Module-Check -ModuleName "Microsoft.Graph.Identity.DirectoryManagement"
```

### Sign in to connect to graph on the tenant you want to test it on and write output for tenant name
```ruby
# Sign in to connect to graph
Connect-MgGraph

# Get tenant details
$tenant = get-mgorganization
$tenantname = $tenant.displayName -replace " ", "-" -replace "\.", ""
Write-output "Tenant Name Output: $tenantname"
```

### Edit URL and run below to test
```ruby
# download file to C:\Temp as Company Default.htm
Invoke-WebRequest -Uri "http://Url/to/signaturefiles/$TenantName/Company Default.htm" -OutFile "C:\Temp\Company Default.htm" 
```

Example folder structure: **start-of-url\Signaturefiles\\_$Tenantname_\Company Default.htm**

</details>

</details>

## Folder path for Signature files used by Entra-Signature.ps1

Signature-Automation/_Tenant Name_ (Replaced " " and "." with "-")/Company Default

