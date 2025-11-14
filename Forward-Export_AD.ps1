Get-ADUser -Filter { altRecipient -like "*" } `
    -Properties altRecipient, deliverAndRedirect, mail, proxyAddresses |
    ForEach-Object {

        #
        # Resolve FORWARDING TARGET (altRecipient)
        #
        $target = Get-ADObject $_.altRecipient -Properties mail, proxyAddresses

        $targetSmtp = $target.mail
        if (-not $targetSmtp -and $target.proxyAddresses) {
            $targetSmtp = ($target.proxyAddresses |
                           Where-Object { $_ -like "SMTP:*" } |
                           ForEach-Object { $_.Substring(5) }) |
                           Select-Object -First 1
        }

        #
        # Resolve USER'S OWN SMTP
        #
        $userSmtp = $_.mail
        if (-not $userSmtp -and $_.proxyAddresses) {
            $userSmtp = ($_.proxyAddresses |
                         Where-Object { $_ -like "SMTP:*" } |
                         ForEach-Object { $_.Substring(5) }) |
                         Select-Object -First 1
        }

        [PSCustomObject]@{
            UserSMTP           = $userSmtp
            DisplayName        = $_.DisplayName
            ForwardToSMTP      = $targetSmtp
            DeliverAndKeepCopy = $_.deliverAndRedirect
        }
    } | Export-Csv "c:\temp\ForwardingResolved.csv" -NoTypeInformation -Encoding UTF8
