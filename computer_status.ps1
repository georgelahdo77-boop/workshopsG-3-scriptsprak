# Computer_status.ps1
# Antal Datorer per site - Lösenordålder (dagar) per användare
# 10 datorer med äldst senaste inloggning

# Datumtolkning med hjälp av Try/catch, ISO + US
$styles = [System.Globalization.DateTimeStyles]::AssumeLocal
$formats = @(
    'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm',
    'yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm', 'yyyy-MM-dd',
    'MM/dd/yyyy HH:mm:ss', 'M/d/yyyy H:mm:ss', 'MM/dd/yyyy HH:mm', 'M/d/yyyy H:mm',
    'MM/dd/yyyy', 'M/d/yyyy' 
)
function Parse-Date {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $sv = [System.Globalization.CultureInfo]'sv-SE'
    $us = [System.Globalization.CultureInfo]'en-US'
    foreach ($fmt in $formats) {
        foreach ($cult in @($sv, $us)) {
            try { return [datetime]::ParseExact($s, $fmt, $cult, $styles) } catch { }
        }
    }
    foreach ($cult in @($sv, $us)) { try { return [datetime]::Parse($s, $cult) } catch { } }
    return $null 
}

# Läs JSON
try {
    $data = Get-Content ".\ad_export.json" -Encoding UTF8 -Raw -ErrorAction Stop | ConvertFrom-Json
}
catch {
    Write-Error "Kunde inte läsa ad_export.json: $($_.Exception.Message)"
    exit 1
}

$now = Get-Date

# Förbered dator objekt med datum och sortnyckel
$computers = $data.computers | ForEach-Object {
    $dt = Parse-Date $_.lastLogon
    [pscustomobject]@{
        name                   = $_.name
        site                   = $_.Site
        operatingSystem        = $_.operatingSystem
        operatingSystemVersion = $_.operatingSystemVersion
        ipAddress              = $_.ipAddress
        enabled                = $_.enabled
        lastLogonDT            = $dt
        LastLogon              = if ($dt) { $dt.ToString('yyyy-MM-dd HH:mm', $culture) } else { '' }
        DaysSinceLastLogon     = if ($dt) { [int]($now - $dt).TotalDays } else { $null }
        sortkey                = if ($dt) { $dt } else { [datetime]::MinValue } # null → äldst
    }
}

# Nu gruppera datorer per site
"--- Datorer per site ----" | Write-Host
$computers | Group-Object -Property site | Sort-Object Name |
ForEach-Object { "{0,-20} {1,3} st" -f $_.Name, $_.Count } | Out-Host
Write-Host ""

# Lösenordsålder per användare (top 5)
"---- Lösenordsålder (dagar) per användare (utdrag) ----" | Write-Host
$data.users | ForEach-Object {
    $dt = Parse-Date $_.passwordLastSet
    [pscustomobject]@{
        Namn  = $_.displayName
        Dagar = if ($dt) { [int]($now - $dt).TotalDays } else { $null }
    }
} | Sort-Object Dagar -Descending | Select-Object -First 5 |
Format-Table -AutoSize | Out-Host
Write-Host ""

# 10 datorer som inte checkat in på längst tid 
"---- 10 datorer med äldst senaste inloggning ----" | Write-Host
$oldest10 = $computers |
Sort-Object sortkey | Select-Object -First 10 |
Select-Object @{n = 'Dator'; e = { $_.name } },
@{n = 'Site'; e = { $_.site } },
@{n = 'SenastSedd'; e = { if ($_.lastLogonDT) { $_.lastLogonDT.ToString('yyyy-MM-dd HH:mm', $culture) } else { 'Aldrig' } } },
@{n = 'DagarSedan'; e = { $_.DaysSinceLastLogon } },
operatingSystem, enabled
$oldest10 | Format-Table -AutoSize | Out-Host

# Exportera full datorstatus till CSV
$computers |
Select-Object name, site, LastLogon, DaysSinceLastLogon,
operatingSystem, operatingSystemVersion, ipAddress, enabled |
Export-Csv -Path ".\computer_status.csv" -NoTypeInformation -Encoding UTF8

Write-Host "CSV skapad: computer_status.csv"
