param (
    [Parameter(Mandatory = $true)]
    [string]$SiteListCsvPath
)

Clear-Host

# GLOBAL VARIABLES
$script:permissions = @()
$script:sharingLinks = @()
$everyoneGroups = @("everyone except external users", "everyone", "all users") | ForEach-Object { $_.ToLower() }

$ClientId = ""
$Thumbprint = ""
$Tenant = ""

$dateTime = (Get-Date).ToString("dd-MM-yyyy-hh-mm-ss")
$invocation = (Get-Variable MyInvocation).Value
$logDirectory = Join-Path (Split-Path $invocation.MyCommand.Path) "Logs"
$errorLogPath = Join-Path $logDirectory "Errors-$dateTime.log"

if (!(Test-Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# ────── UTILITY FUNCTIONS ──────

function Log-Error {
    param (
        [string]$Message,
        [string]$Context
    )
    $entry = "[ERROR] [$((Get-Date).ToString())] [$Context] $Message"
    Write-Host $entry -ForegroundColor Red
    Add-Content -Path $errorLogPath -Value $entry
}

function Add-PermissionEntry {
    param(
        $Object, $Type, $RelativeUrl, $SiteUrl, $SiteTitle, $ListTitle,
        $MemberType, $ParentGroup, $MemberName, $MemberLoginName, $Roles, $SensitivityLabel = ''
    )

    $entry = [PSCustomObject]@{
        SiteUrl          = $SiteUrl
        SiteTitle        = $SiteTitle
        ListTitle        = $ListTitle
        SensitivityLabel = $SensitivityLabel
        Type             = if ($Type -eq 1) {"Folder"} elseif ($Type -eq 0) {"File"} else {$Type}
        RelativeUrl      = $RelativeUrl
        MemberType       = $MemberType
        ParentGroup      = $ParentGroup
        MemberName       = $MemberName
        MemberLoginName  = $MemberLoginName
        Roles            = $Roles -join ","
    }
    $script:permissions += $entry
}

function Add-SharingLinkEntry {
    param(
        $SiteUrl, $ListTitle, $ListId, $RelativeUrl, $ItemUniqueId, $Type, $ShareLink
    )

    $entry = [PSCustomObject]@{
        SiteUrl          = $SiteUrl
        ListTitle        = $ListTitle
        ListId           = $ListId
        RelativeUrl      = $RelativeUrl
        UniqueId         = $ItemUniqueId
        ObjectType       = $Type
        ShareId          = $ShareLink.Id
        RoleList         = $ShareLink.Roles -join "|"
        Users            = $ShareLink.GrantedToIdentitiesV2.User.Email -join "|"
        ShareLinkUrl     = $ShareLink.Link.WebUrl
        ShareLinkType    = $ShareLink.Link.Type
        ShareLinkScope   = $ShareLink.Link.Scope
        Expiration       = $ShareLink.ExpirationDateTime
        BlocksDownload   = $ShareLink.Link.PreventsDownload
        RequiresPassword = $ShareLink.HasPassword
    }
    $script:sharingLinks += $entry
}

function Check-NestedGroupForEEEU {
    param($Group, $Object, $Type, $RelativeUrl, $SiteUrl, $SiteTitle, $ListTitle, $RoleBindings)

    try {
        $users = Get-PnPGroupMember -Identity $Group.Title
        foreach ($user in $users) {
            if ($everyoneGroups -contains $user.Title.ToLower()) {
                Add-PermissionEntry -Object $Object -Type $Type -RelativeUrl $RelativeUrl `
                    -SiteUrl $SiteUrl -SiteTitle $SiteTitle -ListTitle $ListTitle `
                    -MemberType $user.GetType().Name -ParentGroup $Group.Title `
                    -MemberName $user.Title -MemberLoginName $user.LoginName -Roles $RoleBindings
            }
        }
    } catch {
        Log-Error -Message $_.Exception.Message -Context "Nested Group: $($Group.Title)"
    }
}

function Get-ListItemsWithUniquePermission {
    param([Microsoft.SharePoint.Client.List]$List)
    $selectFields = "ID,HasUniqueRoleAssignments,FileRef,FileLeafRef,FileSystemObjectType,UniqueId"
    $Url = $siteUrl + '/_api/web/lists/getbytitle(''' + $($list.Title) + ''')/items?$select=' + $($selectFields)
    $nextLink = $Url
    $listItems = @()
    while ($nextLink) {
        try {
            $response = Invoke-PnPSPRestMethod -Url $nextLink -Method Get
            $listItems += $response.value | Where-Object { $_.HasUniqueRoleAssignments -eq $true }
            $nextLink = $response.'odata.nextlink'
        } catch {
            Write-Host "Error: $_. Retrying..." -ForegroundColor Red
            Start-Sleep -Seconds 10
        }
    }
    return $listItems
}

# ────── MAIN PROCESSING FUNCTIONS ──────

function Process-EEEUPermissions {
    param($SiteUrl)

    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
    } catch {
        Log-Error -Message $_.Exception.Message -Context "Connect EEEU: $SiteUrl"
        return
    }

    $web = Get-PnPWeb
    $siteTitle = $web.Title

    # Site-level
    $siteRoles = Get-PnPProperty -ClientObject $web -Property RoleAssignments
    foreach ($role in $siteRoles) {
        Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings, Member
        $member = $role.Member
        if ($everyoneGroups -contains $member.Title.ToLower()) {
            Add-PermissionEntry $web "Site" "" $SiteUrl $siteTitle "" $member.GetType().Name "" $member.Title $member.LoginName $role.RoleDefinitionBindings.Name
        } elseif ($member.PrincipalType -eq "SharePointGroup") {
            Check-NestedGroupForEEEU $member $web "Site" "" $SiteUrl $siteTitle "" $role.RoleDefinitionBindings.Name
        }
    }

    # List + Item-level
    $lists = Get-PnPList -Includes BaseType, Hidden, Title, HasUniqueRoleAssignments, RootFolder | Where-Object { !$_.Hidden }
    foreach ($list in $lists) {
        $listTitle = $list.Title
        $listUrl = $list.RootFolder.ServerRelativeUrl
        $listId = $list.Id

        try {
            if ($list.HasUniqueRoleAssignments) {
                $listRoles = Get-PnPProperty -ClientObject $list -Property RoleAssignments
                foreach ($role in $listRoles) {
                    Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings, Member
                    $member = $role.Member
                    if ($everyoneGroups -contains $member.Title.ToLower()) {
                        Add-PermissionEntry $list "List" $listUrl $SiteUrl $siteTitle $listTitle $member.GetType().Name "" $member.Title $member.LoginName $role.RoleDefinitionBindings.Name
                    } elseif ($member.PrincipalType -eq "SharePointGroup") {
                        Check-NestedGroupForEEEU $member $list "List" $listUrl $SiteUrl $siteTitle $listTitle $role.RoleDefinitionBindings.Name
                    }
                }
            }

            $items = Get-ListItemsWithUniquePermission -List $list
            foreach ($item in $items) {
                $type = $item.FileSystemObjectType
                $relativeUrl = $item.FileRef
                $itemId = $item.Id
                $ItemuniqueId = $item.UniqueId

                try {
                    $listItem = Get-PnPListItem -List $list -Id $itemId
                    $roles = Get-PnPProperty -ClientObject $listItem -Property RoleAssignments
                    foreach ($role in $roles) {
                        Get-PnPProperty -ClientObject $role -Property RoleDefinitionBindings, Member
                        $member = $role.Member
                        if ($everyoneGroups -contains $member.Title.ToLower()) {
                            Add-PermissionEntry $listItem $type $relativeUrl $SiteUrl $siteTitle $listTitle $member.GetType().Name "" $member.Title $member.LoginName $role.RoleDefinitionBindings.Name
                        } elseif ($member.PrincipalType -eq "SharePointGroup") {
                            Check-NestedGroupForEEEU $member $listItem $type $relativeUrl $SiteUrl $siteTitle $listTitle $role.RoleDefinitionBindings.Name
                        }
                    }
                } catch {
                    Log-Error -Message $_.Exception.Message -Context "Item Role: $relativeUrl"
                }
            }

        } catch {
            Log-Error -Message $_.Exception.Message -Context "List: $listTitle"
        }
    }
}

function Process-SharingLinks {
    param($SiteUrl)

    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
    } catch {
        Log-Error -Message $_.Exception.Message -Context "Connect SharingLinks: $SiteUrl"
        return
    }

    $lists = Get-PnPList -Includes Hidden, Title, RootFolder | Where-Object { !$_.Hidden }
    foreach ($list in $lists) {
        $listTitle = $list.Title
        $listUrl = $list.RootFolder.ServerRelativeUrl
        $listId = $list.Id

        try {
            $items = Get-ListItemsWithUniquePermission -List $list
            foreach ($item in $items) {
                $type = $item.FileSystemObjectType
                $relativeUrl = $item.FileRef
                $itemId = $item.Id
                $ItemuniqueId = $item.UniqueId
                If($relativeUrl -like "*/Social/Private*"){continue}

                try {
                    $shareLinks = if ($type -eq 0) {
                        Get-PnPFileSharingLink -Identity $relativeUrl
                    } else {
                        Get-PnPFolderSharingLink -Folder $relativeUrl
                    }

                    foreach ($link in $shareLinks) {
                        Add-SharingLinkEntry $SiteUrl $listTitle $listId $relativeUrl $uniqueId $type $link
                    }
                } catch {
                    Log-Error -Message $_.Exception.Message -Context "Sharing Link: $relativeUrl"
                }
            }
        } catch {
            Log-Error -Message $_.Exception.Message -Context "List Items (Sharing): $listTitle"
        }
    }
}

function Export-Reports {
    $permFile = Join-Path $logDirectory "Permissions-EEEU-$dateTime.csv"
    $shareFile = Join-Path $logDirectory "SharingLinks-$dateTime.csv"

    $script:permissions | Export-Csv -Path $permFile -NoTypeInformation
    $script:sharingLinks | Export-Csv -Path $shareFile -NoTypeInformation

    Write-Host "`n--- REPORT SUMMARY ---" -ForegroundColor Green
    Write-Host "EEEU Permission Entries: $($script:permissions.Count)"
    Write-Host "Sharing Links Found:    $($script:sharingLinks.Count)"
    Write-Host "EEEU Report:            $permFile"
    Write-Host "Sharing Link Report:    $shareFile"
    Write-Host "Error Log (if any):     $errorLogPath"
}

# ────── MAIN EXECUTION ──────

$siteList = Import-Csv -Path $SiteListCsvPath
$p = 0

foreach ($row in $siteList) {
    $p++
    Write-Host "[$p/$($siteList.Count)] Processing: $($row.SiteUrl)" -ForegroundColor Yellow
    try {
        Process-EEEUPermissions -SiteUrl $row.SiteUrl
        Process-SharingLinks -SiteUrl $row.SiteUrl
    } catch {
        Log-Error -Message $_.Exception.Message -Context "Top-Level: $($row.SiteUrl)"
    }
}

Export-Reports
