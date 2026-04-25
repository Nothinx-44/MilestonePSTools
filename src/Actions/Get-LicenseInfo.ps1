<#
.SYNOPSIS
    Affiche les informations de licence du VMS Milestone via Get-VmsLicensedProducts.
#>

function Get-VmsLicenseSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Log
    )

    # Proprietes internes du SDK Milestone a ignorer
    $skipProps = @(
        'DisplayName','ExpirationDate','LicensedChannels','UsedChannels','Slc',
        'ProductDisplayName','Path','ParentPath','ItemCategory','ParentItemPath',
        'ServerId','PluginId'
    )

    & $Log "--- Produits licencies ---"
    try {
        $products = @(Get-VmsLicensedProducts -ErrorAction Stop)

        if ($products.Count -eq 0) {
            & $Log "  Aucun produit licence trouve."
            return
        }

        foreach ($product in $products) {
            & $Log "Produit : $($product.DisplayName)"

            # Expiration — peut valoir "Unrestricted" ou une vraie date
            $expRaw = $product.ExpirationDate
            if ($expRaw -and $expRaw -ne 'N/A') {
                try {
                    $expDate  = [datetime]$expRaw
                    $daysLeft = ($expDate - (Get-Date)).Days
                    $expStr   = $expDate.ToString('dd/MM/yyyy')
                    if ($daysLeft -lt 0) {
                        & $Log "  ERREUR: Licence expiree le $expStr ($([math]::Abs($daysLeft)) jours depasses)"
                    }
                    elseif ($daysLeft -lt 30) {
                        & $Log "  AVERTISSEMENT: Expiration le $expStr (dans $daysLeft jours)"
                    }
                    else {
                        & $Log "  Expiration      : $expStr (dans $daysLeft jours)"
                    }
                }
                catch {
                    # Valeur non-date (ex: "Unrestricted")
                    & $Log "  Expiration      : $expRaw"
                }
            }
            else {
                & $Log "  Expiration      : Aucune (licence perpetuelle)"
            }

            # SLC
            if ($product.Slc) {
                & $Log "  SLC             : $($product.Slc)"
            }

            # Canaux licencies / utilises
            $licensed = $product.LicensedChannels
            $used     = $product.UsedChannels

            if ($null -ne $licensed) {
                $licStr = if ($licensed -is [int] -or $licensed -match '^\d+$') { $licensed } else { $licensed }
                if ($null -ne $used -and $licensed -match '^\d+$' -and [int]$licensed -gt 0) {
                    $pct = [math]::Round(([int]$used / [int]$licensed) * 100, 1)
                    if ($pct -ge 90) {
                        & $Log "  AVERTISSEMENT: Canaux : $used / $licensed ($pct % utilises)"
                    }
                    else {
                        & $Log "  Canaux          : $used / $licensed ($pct % utilises)"
                    }
                }
                elseif ($null -ne $used) {
                    & $Log "  Canaux          : $used utilises / $licStr"
                }
                else {
                    & $Log "  Canaux          : $licStr"
                }
            }

            # Care Plus / Care Premium si pertinent
            foreach ($careProp in @('CarePlus','CarePremium')) {
                $val = $product.$careProp
                if ($val -and $val -ne 'N/A') {
                    & $Log "  $careProp       : $val"
                }
            }
        }
    }
    catch {
        & $Log "ERREUR: Get-VmsLicensedProducts : $_"
    }
}
