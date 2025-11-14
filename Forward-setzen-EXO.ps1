
$csv = Import-Csv "C:\Temp\ForwardingResolved.csv" -Delimiter ";"

$report = @()

foreach ($row in $csv) {

    # --- Get mailbox in EXO ---
    $mbx = Get-Mailbox -Identity $row.UserSMTP -ErrorAction SilentlyContinue
    if (-not $mbx) {
        Write-Host "Mailbox not found in EXO: $($row.UserSMTP)" -ForegroundColor Yellow
        $report += [PSCustomObject]@{
            UserSMTP                        = $row.UserSMTP
            TargetFromCsv                   = $row.ForwardToSMTP
            RecipientResolved               = $null
            Action                          = "Mailbox not found in EXO"
            BeforeForwardingAddress         = $null
            BeforeForwardingSmtpAddress     = $null
            BeforeDeliverToMailboxAndForward= $null
            AfterForwardingAddress          = $null
            AfterForwardingSmtpAddress      = $null
            AfterDeliverToMailboxAndForward = $null
        }
        continue
    }

    # --- Capture BEFORE state (as SMTP where possible) ---
    $beforeFA  = $null
    if ($mbx.ForwardingAddress) {
        $baRec = Get-Recipient -Identity $mbx.ForwardingAddress -ErrorAction SilentlyContinue
        if ($baRec) { $beforeFA = $baRec.PrimarySmtpAddress.ToString() }
        else        { $beforeFA = $mbx.ForwardingAddress.ToString() }
    }

    $beforeFSA = $null
    if ($mbx.ForwardingSmtpAddress) {
        # ForwardingSmtpAddress is of form "SMTP:user@domain.com"
        $beforeFSA = $mbx.ForwardingSmtpAddress.ToString().Replace("SMTP:","")
    }

    $beforeDeliver = $mbx.DeliverToMailboxAndForward

    # --- Decide what we want to do based on recipient existence ---
    $recipient = $null
    if ($row.ForwardToSMTP) {
        # Try to find a *recipient* in EXO for this SMTP
        $recipient = Get-Recipient -ResultSize 1 -ErrorAction SilentlyContinue -Filter "EmailAddresses -eq 'SMTP:$($row.ForwardToSMTP)'"
    }

    # Parse DeliverAndKeepCopy from CSV into [bool]
    $deliver = $false
    if ($row.DeliverAndKeepCopy) {
        $deliver = [System.Convert]::ToBoolean($row.DeliverAndKeepCopy)
    }

    $action = ""

    if ($recipient) {
        Write-Host "[$($row.UserSMTP)] Directory recipient found for $($row.ForwardToSMTP) -> using ForwardingAddress" -ForegroundColor Green

        # Set ForwardingAddress (directory object), clear ForwardingSmtpAddress
        Set-Mailbox -Identity $mbx.Identity `
            -ForwardingAddress $recipient.Identity `
            -ForwardingSmtpAddress $null `
            -DeliverToMailboxAndForward:$deliver

        $action = "Set ForwardingAddress to $($recipient.PrimarySmtpAddress); cleared ForwardingSmtpAddress"
    }
    elseif ($row.ForwardToSMTP) {
        Write-Host "[$($row.UserSMTP)] No directory recipient for $($row.ForwardToSMTP) -> using ForwardingSmtpAddress" -ForegroundColor Cyan

        # Set ForwardingSmtpAddress (external SMTP), clear ForwardingAddress
        Set-Mailbox -Identity $mbx.Identity `
            -ForwardingAddress $null `
            -ForwardingSmtpAddress $row.ForwardToSMTP `
            -DeliverToMailboxAndForward:$deliver

        $action = "Set ForwardingSmtpAddress to $($row.ForwardToSMTP); cleared ForwardingAddress"
    }
    else {
        Write-Host "[$($row.UserSMTP)] No forwarding target in CSV -> skipped" -ForegroundColor DarkYellow
        $action = "No target in CSV; skipped"
    }

    # --- Capture AFTER state (again, try to surface SMTP) ---
    $mbxAfter = Get-Mailbox -Identity $mbx.Identity

    $afterFA = $null
    if ($mbxAfter.ForwardingAddress) {
        $aaRec = Get-Recipient -Identity $mbxAfter.ForwardingAddress -ErrorAction SilentlyContinue
        if ($aaRec) { $afterFA = $aaRec.PrimarySmtpAddress.ToString() }
        else        { $afterFA = $mbxAfter.ForwardingAddress.ToString() }
    }

    $afterFSA = $null
    if ($mbxAfter.ForwardingSmtpAddress) {
        $afterFSA = $mbxAfter.ForwardingSmtpAddress.ToString().Replace("SMTP:","")
    }

    $afterDeliver = $mbxAfter.DeliverToMailboxAndForward

    # --- Add to report object ---
    $report += [PSCustomObject]@{
        UserSMTP                        = $row.UserSMTP
        TargetFromCsv                   = $row.ForwardToSMTP
        RecipientResolved               = if ($recipient) { $recipient.PrimarySmtpAddress.ToString() } else { $null }
        Action                          = $action
        BeforeForwardingAddress         = $beforeFA
        BeforeForwardingSmtpAddress     = $beforeFSA
        BeforeDeliverToMailboxAndForward= $beforeDeliver
        AfterForwardingAddress          = $afterFA
        AfterForwardingSmtpAddress      = $afterFSA
        AfterDeliverToMailboxAndForward = $afterDeliver
    }
}

# --- Output & export report ---
$report | Format-Table -AutoSize

$report | Export-Csv "c:\temp\ForwardingChangeReport.csv" -NoTypeInformation -Encoding UTF8 -Delimiter ";" -Force
Write-Host "Report written to c:\temp\ForwardingChangeReport.csv" -ForegroundColor Magenta
