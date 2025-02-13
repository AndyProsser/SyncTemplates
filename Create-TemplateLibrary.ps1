<#====================================#>
<# CONFIGURATION #>
$isTenantAdmin = $true
$sharePointUrl = "https://exectechs.sharepoint.com/"
$templateSite = [PSCustomObject]@{
    Name    = "Company Templates"
    Alias   = "templates"
    Locale  = 1033  # en-US
    Tz      = [PnP.Framework.Enums.TimeZone]::UTCPLUS1000_CANBERRA_MELBOURNE_SYDNEY
    Assets  = "Assets"
    Library = "Templates"
}
<#====================================#>

# Install PorwerShell Libraries
Write-Host "Loading Required Modules"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if (!(Get-Command -Module PnP.Powershell)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Install-Module PnP.PowerShell -Scope CurrentUser
    Register-PnPManagementShellAccess -ShowConsentUrl:(!$isTenantAdmin)
}
else {
    Import-Module -Name PnP.Powershell
}

Write-Host "Creating SharePoint Online Template Site"
#Function to Ensure a SharePoint Online document library
Function Ensure-DocumentLibrary() {
    param
    (
        [Parameter(Mandatory = $true)] [string] $LibraryName,
        [Parameter(Mandatory = $true)] [PnP.PowerShell.Commands.Base.PnPConnection] $connection
    )

    try {
        Write-Host " - Ensuring Library '$LibraryName' Exists..." -NoNewline

        #Check if the Library exist already
        $list = Get-PnPList -Connection $connection | Where-Object { $_.Title -eq $LibraryName } 
        if ($list -eq $null) {
            #Create Document Library
            Write-Host -ForegroundColor Yellow "MISSING" -NoNewline
            Write-Host "..." -NoNewline
            New-PnPList -Title $LibraryName -Template DocumentLibrary -OnQuickLaunch -Connection $connection | Out-Null
            Write-Host -ForegroundColor Green "OK"
        }
        else {
            Write-Host -ForegroundColor Green "OK"
        }
    } catch {
        Write-Host -ForegroundColor Red "ERROR"
        Write-Error $_.Exception.Message
    }
}

Function Ensure-TemplateSite() {
    param
    (
        [Parameter(Mandatory = $true)] [PSCustomObject] $TemplateSite,
        [Parameter(Mandatory = $true)] [PnP.PowerShell.Commands.Base.PnPConnection] $Connection
    )

    try {
        $siteName = $TemplateSite.Name
        Write-Host " - Ensuring Site '$siteName' Exists..." -NoNewline

        #Check if the Template Site exist already
        $site = Get-PnPTenantSite -Connection $Connection | Where-Object { $_.Title -eq $TemplateSite.Name }  
        if ($site -eq $null) {
            #Create Document Library
            Write-Host -ForegroundColor Yellow "MISSING" -NoNewline
            Write-Host "..." -NoNewline
            $url = New-PnPSite -Type TeamSite -Title $TemplateSite.Name -Alias $TemplateSite.Alias -IsPublic -Lcid $TemplateSite.Locale -TimeZone $TemplateSite.Tz -Connection $Connection -Wait
            Write-Host -ForegroundColor Green "OK"
        }
        else {
            $url = $site.Url
            Write-Host -ForegroundColor Green "OK"
        }
    } catch {
        Write-Host -ForegroundColor Red "ERROR"
        Write-Error $_.Exception.Message
    }
    return $url
}

Function Ensure-SitePermissions() {
    param
    (
        [Parameter(Mandatory = $true)] [string] $MemberName,
        [Parameter(Mandatory = $true)] [PnP.PowerShell.Commands.Base.PnPConnection] $Connection
    )

    try {
        Write-Host " - Ensuring '$MemberName' Exists in Visitors Group..." -NoNewline

        #Get the Group to Add - Default Visitors group of the site
        $Group = Get-PnPGroup -AssociatedVisitorGroup -Connection $Connection

        #Check if the Member exist already
        $members = Get-PnPGroupMember -Identity $Group -Connection $siteConnection | Where-Object { $_.Title -eq $MemberName } 
        if ($members -eq $null) {
            #Create Document Library
            Write-Host -ForegroundColor Yellow "MISSING" -NoNewline
            Write-Host "..." -NoNewline
            Add-PnPGroupMember -Identity $Group -LoginName $MemberName -Connection $siteConnection | Out-Null
            Write-Host -ForegroundColor Green "OK"
        }
        else {
            Write-Host -ForegroundColor Green "OK"
        }
    } catch {
        Write-Host -ForegroundColor Red "ERROR"
        Write-Error $_.Exception.Message
    }
}

# Connect to SharePoint Online
Write-Host " - Connecting to SharePoint Online"
$adminConnection = Connect-PnPOnline -Url $sharePointUrl -Interactive -ReturnConnection

# Ensure Template Site Exist
$siteUrl = Ensure-TemplateSite -TemplateSite $templateSite -Connection $adminConnection
$siteConnection = Connect-PnPOnline -Url $siteUrl -Interactive -ReturnConnection

# Ensure Permissions Exist
Ensure-SitePermissions -MemberName "everyone except external users" -Connection $siteConnection

# Ensure Libraries Exist
Ensure-DocumentLibrary -LibraryName $templateSite.Assets -Connection $siteConnection
Ensure-DocumentLibrary -LibraryName $templateSite.Library -Connection $siteConnection

# Get Template Site Details
$tenant = Get-PnPTenantInfo -CurrentTenant -Connection $siteConnection
$site = Get-PnPSite -Connection $siteConnection -Includes ID, Url
$web = Get-PnPWeb -Connection $siteConnection -Includes ID, Url
$list = Get-PnPList -Connection $siteConnection -Includes ID | Where-Object { $_.Title -eq $templateSite.Library }

$params = @{
    orgName   = $tenant.DisplayName
    siteId    = $site.Id
    webId     = $web.Id
    listId    = $list.Id
    webUrl    = $web.Url
    webTitle  = $web.Title
    listTitle = $list.Title
}

Write-Host
Write-Host "SharePoint Online Connection Parameters"
$params | Out-String | Write-Host
