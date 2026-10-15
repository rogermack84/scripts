<#
.SYNOPSIS
    Creates a new Microsoft 365 / Entra ID user, assigns a license, and adds them to groups.

.DESCRIPTION
    First-project onboarding script. Uses the Microsoft Graph PowerShell SDK.
    Run with -WhatIf first to see what it would do without changing anything.

.EXAMPLE
    .\New-StaffUser.ps1 -FirstName "Jane" -LastName "Doe" -Department "Programs" -JobTitle "Program Officer" -WhatIf

.NOTES
    One-time setup (run once in PowerShell 7):
        Install-Module Microsoft.Graph -Scope CurrentUser
    Find your license SKU names:
        Connect-MgGraph -Scopes Organization.Read.All
        Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)] [string] $FirstName,
    [Parameter(Mandatory)] [string] $LastName,
    [Parameter(Mandatory)] [string] $Department,
    [Parameter(Mandatory)] [string] $JobTitle,
    [string] $ManagerEmail,                        # Optional: manager's sign-in address
    [string[]] $DirectReports = @(),               # Optional: sign-in addresses of people who report to this user

    # ---- Change these defaults to match your organization ----
    [string]   $Domain        = "yourorg.org",
    [string]   $LicenseSku    = "SPE_E3",          # SkuPartNumber from Get-MgSubscribedSku
    [string[]] $Groups        = @("All Staff"),    # Security / Microsoft 365 group display names
    [string[]] $DistributionLists = @(),           # Distribution list names, e.g. @("allstaff@yourorg.org")
    [string]   $UsageLocation = "US"               # Required before a license can be assigned
)

# --- 1. Connect to Microsoft Graph with only the permissions we need ---
Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All", "Organization.Read.All" -NoWelcome

# --- 2. Build the user's names ---
$displayName = "$FirstName $LastName"
$mailNick    = ("$FirstName.$LastName" -replace "[^a-zA-Z0-9.]", "").ToLower()
$upn         = "$mailNick@$Domain"

# Stop if the account already exists
if (Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue) {
    Write-Warning "$upn already exists. Nothing was changed."
    return
}

# --- 3. Generate a temporary password (user must change it at first sign-in) ---
$chars        = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#$%"
$tempPassword = -join (1..16 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })

# --- 4. Create the user ---
if ($PSCmdlet.ShouldProcess($upn, "Create user")) {
    $user = New-MgUser -DisplayName $displayName `
        -GivenName $FirstName -Surname $LastName `
        -UserPrincipalName $upn -MailNickname $mailNick `
        -Department $Department -JobTitle $JobTitle `
        -UsageLocation $UsageLocation -AccountEnabled `
        -PasswordProfile @{ Password = $tempPassword; ForceChangePasswordNextSignIn = $true }

    Write-Host "Created $upn" -ForegroundColor Green
}

# Set the manager (a separate call: New-MgUser has no manager parameter)
if ($ManagerEmail) {
    $manager = Get-MgUser -UserId $ManagerEmail -ErrorAction SilentlyContinue
    if (-not $manager) {
        Write-Warning "Manager '$ManagerEmail' not found. Skipping."
    }
    elseif ($PSCmdlet.ShouldProcess($upn, "Set manager to $ManagerEmail")) {
        Set-MgUserManagerByRef -UserId $user.Id -OdataId "https://graph.microsoft.com/v1.0/users/$($manager.Id)"
        Write-Host "Manager set to $($manager.DisplayName)" -ForegroundColor Green
    }
}

# Set direct reports: make the new user the manager of each person listed.
# (Entra has no "add direct report" call; you set the manager on each report.)
foreach ($reportEmail in $DirectReports) {
    $report = Get-MgUser -UserId $reportEmail -Property Id, DisplayName, Manager -ExpandProperty Manager -ErrorAction SilentlyContinue
    if (-not $report) {
        Write-Warning "Direct report '$reportEmail' not found. Skipping."
        continue
    }

    # Show who they reported to before, so the change is easy to spot or undo
    $previous = $report.Manager.AdditionalProperties.displayName
    $action   = if ($previous) { "Change manager from $previous to $displayName" } else { "Set manager to $displayName" }

    if ($PSCmdlet.ShouldProcess($reportEmail, $action)) {
        Set-MgUserManagerByRef -UserId $report.Id -OdataId "https://graph.microsoft.com/v1.0/users/$($user.Id)"
        Write-Host "$($report.DisplayName) now reports to $displayName" -ForegroundColor Green
    }
}

# --- 5. Assign the license ---
$sku = Get-MgSubscribedSku | Where-Object SkuPartNumber -eq $LicenseSku
if (-not $sku) {
    Write-Warning "License '$LicenseSku' not found. Check Get-MgSubscribedSku."
}
elseif ($PSCmdlet.ShouldProcess($upn, "Assign license $LicenseSku")) {
    Set-MgUserLicense -UserId $user.Id -AddLicenses @(@{ SkuId = $sku.SkuId }) -RemoveLicenses @() | Out-Null
    Write-Host "Assigned $LicenseSku" -ForegroundColor Green
}

# --- 6. Add to groups ---
foreach ($groupName in $Groups) {
    $group = Get-MgGroup -Filter "displayName eq '$groupName'"
    if (-not $group) {
        Write-Warning "Group '$groupName' not found. Skipping."
        continue
    }
    if ($PSCmdlet.ShouldProcess($upn, "Add to group $groupName")) {
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id
        Write-Host "Added to $groupName" -ForegroundColor Green
    }
}

# --- 7. Add to distribution lists (Exchange Online) ---
# Distribution lists can't be managed through Graph, so this uses the
# ExchangeOnlineManagement module. One-time setup:
#     Install-Module ExchangeOnlineManagement -Scope CurrentUser
if ($DistributionLists.Count -gt 0) {
    Connect-ExchangeOnline -ShowBanner:$false

    if (-not $WhatIfPreference) {
        # Exchange needs a few minutes after licensing to create the mailbox.
        # Wait until it exists (up to about 15 minutes) before adding to lists.
        Write-Host "Waiting for the mailbox to be created..." -ForegroundColor Yellow
        $tries = 0
        while (-not (Get-Recipient -Identity $upn -ErrorAction SilentlyContinue) -and $tries -lt 30) {
            Start-Sleep -Seconds 30
            $tries++
        }
        if (-not (Get-Recipient -Identity $upn -ErrorAction SilentlyContinue)) {
            Write-Warning "Mailbox not ready yet. Add $upn to distribution lists later."
            $DistributionLists = @()
        }
    }

    foreach ($list in $DistributionLists) {
        if ($PSCmdlet.ShouldProcess($upn, "Add to distribution list $list")) {
            try {
                Add-DistributionGroupMember -Identity $list -Member $upn -ErrorAction Stop
                Write-Host "Added to $list" -ForegroundColor Green
            }
            catch {
                Write-Warning "Could not add to '$list': $($_.Exception.Message)"
            }
        }
    }

    Disconnect-ExchangeOnline -Confirm:$false
}

# --- 8. Summary ---
if (-not $WhatIfPreference) {
    Write-Host "`nDone. Give the new hire these sign-in details securely:" -ForegroundColor Cyan
    Write-Host "  Username:           $upn"
    Write-Host "  Temporary password: $tempPassword"
}
