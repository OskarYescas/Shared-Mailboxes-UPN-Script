# ==============================================================================
# Exchange Online Shared Mailbox Resolver & Migration Planner CSV Generator
# ==============================================================================

# 1. Connect to Exchange Online if not already connected
if (-not (Get-ConnectionInformation)) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline
}

# 2. Prompt for or specify input CSV file path
$csvPath = Read-Host -Prompt "Enter the path to your input CSV file (e.g., C:\Migration\input.csv)"

# Strip surrounding quotes if user drag-and-dropped file into terminal
$csvPath = $csvPath.Trim('"').Trim("'")

if (-not (Test-Path $csvPath)) {
    Write-Error "File not found at: $csvPath"
    return
}

# Define output file paths in the same directory as input
$workingDir = Split-Path -Path $csvPath -Parent
$cleanCsvPath = Join-Path -Path $workingDir -ChildPath "MigrationPlanner_SharedMailboxes_UPN.csv"
$reportCsvPath = Join-Path -Path $workingDir -ChildPath "ExchangeOnline_Full_Audit_Report.csv"

# 3. Read CSV & Auto-detect Email Column
$inputData = Import-Csv -Path $csvPath
$totalCount = $inputData.Count
Write-Host "Successfully loaded $totalCount entries from CSV." -ForegroundColor Cyan

$firstRow = $inputData[0]
$emailColumn = @("Email Id", "Email", "PrimarySmtpAddress", "UserPrincipalName", "Identity", "EmailAddress") | 
    Where-Object { $_ -in $firstRow.PSObject.Properties.Name } | Select-Object -First 1

if (-not $emailColumn) {
    Write-Error "Could not find a valid email column. Please ensure CSV has a header like 'Email Id', 'Email', or 'PrimarySmtpAddress'."
    return
}

Write-Host "Using column '$emailColumn' for Exchange Online lookup..." -ForegroundColor Yellow
Write-Host "Processing... Please wait.`n" -ForegroundColor Cyan

# 4. Resolve entries against Exchange Online (Bulk In-Memory Lookup)
Write-Host "Pre-fetching Exchange Online mailboxes into memory..." -ForegroundColor Cyan

# Fetch all mailboxes in 1 single API call using EXO V3
$allMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, EmailAddresses

# Build in-memory hashtable mapping UPNs, Primary SMTP, and Aliases -> Mailbox Object
$mailboxLookup = @{}
foreach ($mbx in $allMailboxes) {
    if ($mbx.UserPrincipalName)   { $mailboxLookup[$mbx.UserPrincipalName.ToLower()] = $mbx }
    if ($mbx.PrimarySmtpAddress)  { $mailboxLookup[$mbx.PrimarySmtpAddress.ToLower()] = $mbx }
    
    # Map all proxy/alias addresses (e.g. smtp:alias@domain.com)
    foreach ($addr in $mbx.EmailAddresses) {
        $cleanAddr = $addr -replace '^(?i)smtp:', ''
        if ($cleanAddr -and -not $mailboxLookup.ContainsKey($cleanAddr.ToLower())) {
            $mailboxLookup[$cleanAddr.ToLower()] = $mbx
        }
    }
}

Write-Host "Loaded $($mailboxLookup.Count) email address/alias mappings from Exchange Online.`n" -ForegroundColor Green

# Process CSV against in-memory lookup table
$index = 0
$results = foreach ($row in $inputData) {
    $index++
    $rawEmail = $row.$emailColumn.Trim()
    
    Write-Progress -Activity "Resolving Exchange Online Mailboxes" -Status "Processing $index of $totalCount ($rawEmail)" -PercentComplete (($index / $totalCount) * 100)
    
    if (-not $rawEmail) { continue }
    
    $emailKey = $rawEmail.ToLower()
    
    if ($mailboxLookup.ContainsKey($emailKey)) {
        $mbx = $mailboxLookup[$emailKey]
        [PSCustomObject]@{
            "Input Email"          = $rawEmail
            "UserPrincipalName"   = $mbx.UserPrincipalName
            "PrimarySmtpAddress" = $mbx.PrimarySmtpAddress
            "RecipientType"      = $mbx.RecipientTypeDetails
            "Status"              = "Resolved"
        }
    } else {
        [PSCustomObject]@{
            "Input Email"          = $rawEmail
            "UserPrincipalName"   = "NOT_FOUND"
            "PrimarySmtpAddress"  = "NOT_FOUND"
            "RecipientType"      = "NOT_FOUND"
            "Status"              = "Not Found in Exchange"
        }
    }
}

# 5. Export Full Diagnostic Report
$results | Export-Csv -Path $reportCsvPath -NoTypeInformation

# 6. Filter only Shared Mailboxes with valid UPNs for Migration Planner
$sharedMailboxes = $results | Where-Object { $_.RecipientType -eq "SharedMailbox" -and $_.UserPrincipalName -ne "NOT_FOUND" }
$sharedMailboxes | Select-Object @{N="Email Id"; E={$_.UserPrincipalName}} | Export-Csv -Path $cleanCsvPath -NoTypeInformation

# 7. Print Console Summary
Write-Host "============================================================" -ForegroundColor Green
Write-Host "                    PROCESSING COMPLETE                     " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host " Total Input Rows           : $totalCount"
Write-Host " Successfully Resolved      : $(($results | Where-Object {$_.Status -eq 'Resolved'}).Count)"
Write-Host " Shared Mailboxes Found     : $($sharedMailboxes.Count)"
Write-Host " User Mailboxes / Other     : $(($results | Where-Object {$_.RecipientType -ne 'SharedMailbox' -and $_.Status -eq 'Resolved'}).Count)"
Write-Host " Not Found / Skipped        : $(($results | Where-Object {$_.Status -ne 'Resolved'}).Count)"
Write-Host "============================================================" -ForegroundColor Green
Write-Host "`nOUTPUT FILES GENERATED:" -ForegroundColor Yellow
Write-Host " 1. Migration Planner CSV   : $cleanCsvPath" -ForegroundColor White
Write-Host " 2. Full Audit Report       : $reportCsvPath`n" -ForegroundColor White
