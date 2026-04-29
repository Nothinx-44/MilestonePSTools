function Get-VmsLicenseSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [scriptblock]$Log
    )

    & $Log $script:T.LI_LogHeader
    try {
        $products = @(Get-VmsLicensedProducts -ErrorAction Stop)

        if ($products.Count -eq 0) {
            & $Log $script:T.LI_LogNone
            return
        }

        foreach ($product in $products) {
            & $Log ($script:T.LI_LogProduct -f $product.DisplayName)

            $expRaw = $product.ExpirationDate
            if ($expRaw -and $expRaw -ne 'N/A') {
                try {
                    $expDate  = [datetime]$expRaw
                    $daysLeft = ($expDate - (Get-Date)).Days
                    $expStr   = $expDate.ToString('dd/MM/yyyy')
                    if ($daysLeft -lt 0) {
                        & $Log ($script:T.LI_LogExpired -f $expStr, [math]::Abs($daysLeft))
                    }
                    elseif ($daysLeft -lt 30) {
                        & $Log ($script:T.LI_LogExpSoon -f $expStr, $daysLeft)
                    }
                    else {
                        & $Log ($script:T.LI_LogExpiry -f $expStr, $daysLeft)
                    }
                }
                catch {
                    & $Log ($script:T.LI_LogExpiryRaw -f $expRaw)
                }
            }
            else {
                & $Log $script:T.LI_LogPerpetual
            }

            if ($product.Slc) {
                & $Log ($script:T.LI_LogSlc -f $product.Slc)
            }

            $licensed = $product.LicensedChannels
            $used     = $product.UsedChannels

            if ($null -ne $licensed) {
                if ($null -ne $used -and $licensed -match '^\d+$' -and [int]$licensed -gt 0) {
                    $pct = [math]::Round(([int]$used / [int]$licensed) * 100, 1)
                    if ($pct -ge 90) {
                        & $Log ($script:T.LI_LogChanWarn -f $used, $licensed, $pct)
                    }
                    else {
                        & $Log ($script:T.LI_LogChan -f $used, $licensed, $pct)
                    }
                }
                elseif ($null -ne $used) {
                    & $Log ($script:T.LI_LogChanSingle -f $used, $licensed)
                }
                else {
                    & $Log ($script:T.LI_LogChanRaw -f $licensed)
                }
            }

            foreach ($careProp in @('CarePlus','CarePremium')) {
                $val = $product.$careProp
                if ($val -and $val -ne 'N/A') {
                    & $Log ($script:T.LI_LogCareProp -f $careProp, $val)
                }
            }
        }
    }
    catch {
        & $Log ($script:T.LI_LogError -f $_)
    }
}
