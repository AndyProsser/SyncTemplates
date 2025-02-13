<#====================================#>
<# CONFIGURATION #>
$siteId    = 'd0a2f731-3f1e-40d2-b429-d26eff835ca9'
$webId     = 'b3f1c3fd-24bd-40e2-a8b5-86fb07efb3d9'
$listId    = '34640710-9bcf-435f-86cf-4701cf5583a6'
$listTitle = 'Templates'
$webUrl    = 'https://<tenant>.sharepoint.com/teams/templates'
$webTitle  = 'Company Templates'
$orgName   = 'TheExecTechs'
<#====================================#>

# Sync Template Library
function Sync-SharepointLocation {
    param (
        [guid]$siteId,
        [guid]$webId,
        [guid]$listId,
        [mailaddress]$userEmail,
        [string]$webUrl,
        [string]$webTitle,
        [string]$listTitle,
        [string]$syncPath
    )

    try {
        Add-Type -AssemblyName System.Web
        #Encode site, web, list, url & email
        [string]$siteId = [System.Web.HttpUtility]::UrlEncode($siteId)
        [string]$webId = [System.Web.HttpUtility]::UrlEncode($webId)
        [string]$listId = [System.Web.HttpUtility]::UrlEncode($listId)
        [string]$userEmail = [System.Web.HttpUtility]::UrlEncode($userEmail)
        [string]$webUrl = [System.Web.HttpUtility]::UrlEncode($webUrl)
        #build the URI
        $uri = New-Object System.UriBuilder
        $uri.Scheme = "odopen"
        $uri.Host = "sync"
        $uri.Query = "siteId=$siteId&webId=$webId&listId=$listId&userEmail=$userEmail&webUrl=$webUrl&listTitle=$listTitle&webTitle=$webTitle"
        #launch the process from URI
        Write-Host $uri.ToString()
        start-process -filepath $($uri.ToString())
    } catch {
        $errorMsg = $_.Exception.Message
    }
    if ($errorMsg) {
        Write-Warning "Sync failed."
        Write-Warning $errorMsg
    } else {
        Write-Host "Sync completed."
        while (!(Get-ChildItem -Path $syncPath -ErrorAction SilentlyContinue)) {
            Start-Sleep -Seconds 2
        }
        return $true
    }    
}
#endregion

#region Main Process
try {
    #region Sharepoint Sync
    #[mailaddress]$userUpn = cmd /c "whoami/upn"
    #Get Loggedin User
    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    #Check if AAD account
    $sid = $user.user.value
    [mailaddress]$userUpn = Get-ItemPropertyValue -path "HKLM:\SOFTWARE\Microsoft\IdentityStore\Cache\$sid\IdentityCache\$sid\" -name "UserName"
    Write-Host "User $userUpn detected"

    $syncPath = "$(split-path $env:onedrive)\" + $orgName + "\$($params.webTitle) - $($Params.listTitle)"

    $params = @{
        #replace with data captured from your sharepoint site.
        siteId    = "{$siteId"
        webId     = "{$webId}"
        listId    = "{$listId}"
        userEmail = $userUpn
        syncPath  = $syncPath
        webUrl    = $webUrl
        webTitle  = $webTitle
        listTitle = $listTitle
    }
    
    Write-Host "SharePoint params:"
    $params | Format-Table
    if (!(Test-Path $($params.syncPath))) {
        Write-Host " - Sharepoint folder not found locally, will now sync.." -ForegroundColor Yellow
        $sp = Sync-SharepointLocation @params
        if (!($sp)) {
            Throw " - Sharepoint sync failed."
        }
    }
    else {
        Write-Host " - Location already syncronized: $($params.syncPath)" -ForegroundColor Yellow
    }
    #endregion

    #region Office 365 Workgroup Templates
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\General" -Name "SharedTemplates" -Value $syncPath
    #endregion
} catch {
    $errorMsg = $_.Exception.Message
} finally {
    if ($errorMsg) {
        Write-Warning $errorMsg
        Throw $errorMsg
    } else { Write-Host "Completed successfully.." }
}
#endregion

################################################################################################
