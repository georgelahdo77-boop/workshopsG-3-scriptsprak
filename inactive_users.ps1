# inactive_users.ps1
# CSV och datumtolkning

$culture = [System.Globalization.CultureInfo]'sv-SE'
$styles = [System.Globalization.DateTimeStyles]::AssumeLocal
$formats = @(
    'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd',          # ISO
    'MM/dd/yyyy HH:mm:ss', 'M/d/yyyy H:mm:ss',    # US med tid
    'MM/dd/yyyy', 'M/d/yyyy'                      # US utan tid    
)

# Fixa datumtolkning utan warning i flödet
function  Parse-Date {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $sv = [System.Globalization.CultureInfo]'sv-SE'
    $us = [System.Globalization.CultureInfo]'en-US'

    foreach ($fmt in $formats) {
        foreach ($cult in @($sv, $us)) {
            try { return [datetime]:: ParseExact($s, $fmt, $cult, $styles) } catch { }
        }
    }
    foreach ($cult in @($sv, $us)) { try { return [datetime]::Parse($s, $cult) } catch { } }
    return $null
}

function Get-InactiveAccounts {
    [CmdletBinding()]
    param([int]$Days = 30)

    # Läs JSON
    try {
        $json = Get-Content ".\ad_export.json" -Encoding UTF8 -Raw -ErrorAction Stop | ConvertFrom-Json
    }
    catch {
        throw "kunde inte läsa ad_export.json: $($_.Exception.Message)"
    } 
    
    $now = Get-Date
    $th = $now.AddDays(-$Days)

    # Bygg objekt och filtrera i pipeline
    $json.users |
    ForEach-Object {
        $dt = Parse-Date $_.lastLogon
        [pscustomobject]@{
            Namn                 = $_.displayName
            Konto                = $_.samAccountName
            Avdelning            = $_.department
            SenasteInloggad      = if ($dt) { $dt.ToString('yyyy-MM-dd HH:mm', $culture) } else { '' }
            DagarSedanInloggning = if ($dt) { [int]($now - $dt).TotalDays } else { $null }
            Enabled              = $_.enabled
            _LastLogonDT         = $dt
        }
    } |
    Where-Object { $_._LastLogonDT -and $_._LastLogonDT -lt $th } |
    Sort-Object DagarSedanInloggning -Descending
}

# hämta och köra + visa i konsolen
$inactive = Get-InactiveAccounts -Days 30

if ($inactive -and $inactive.Count) {
    "Inaktiva användare (30+ dagar):" | Write-Host
    $inactive |
    Select-Object Namn, Konto, Avdelning,
    @{n = 'SenastInloggning'; e = { $_.SenastInloggad } },
    DagarSedanInloggning, Enabled |
    Format-Table -AutoSize | Out-Host
}
else {
    "Inga inaktiva användare över 30 dagar." | Write-Host
}

# Exportera CSV
$inactive |
Select-Object Namn, Konto, Avdelning, SenasteInloggad, DagarSedanInloggning, Enabled |
Export-Csv -Path ".\inactive_users.csv" -NoTypeInformation -Encoding UTF8

Write-Host "CSV skapad: inactive_users.csv"