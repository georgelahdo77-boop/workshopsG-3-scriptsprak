# read-json.ps1
# Skriver input för att få rätt svenska bokstäver
# Hjälpfunktion: Datumtolkning med try/catch

function Parse-Date {
    param([string]$s)

    if ([string]::IsNullOrWhiteSpace($s)) { return $null }

    $sv = [System.Globalization.CultureInfo]'sv-SE'
    $us = [System.Globalization.CultureInfo]'en-US'
    $styles = [System.Globalization.DateTimeStyles]::AssumeLocal

    # Vanliga format i källdatan (ISO + US)
    $formats = @(
        'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd',
        'MM/dd/yyyy HH:mm:ss', 'M/d/yyyy H:mm:ss',
        'MM/dd/yyyy', 'M/d/yyyy'
    )

    foreach ($fmt in $formats) {
        foreach ($cult in @($sv, $us)) {
            try {
                return [datetime]::ParseExact($s, $fmt, $cult, $styles)
            }
            catch { } # prova nästa format/kultur
        }
    }

    # fri parsning i båda kulturerna
    foreach ($cult in @($sv, $us)) {
        try { return [datetime]::Parse($s, $cult) } catch { }
    }

    Write-Warning "Kunde inte tolka datumvärdet: '$s'"
    return $null
}


# Läs JSON
try {
    $data = Get-Content -Path ".\ad_export.json" -Encoding UTF8 -Raw -ErrorAction Stop | ConvertFrom-Json
}
catch {
    Write-Error "Kunde inte läsa ad_export.json: $($_.Exception.Message)"
    exit 1
}

$now = Get-Date
$svCulture = [System.Globalization.CultureInfo]'sv-SE'

# Visa domännamn och exportdatum
$exportDt = Parse-Date $data.export_date
"Domän: {0}" -f $data.domain | Write-Host
if ($exportDt) {
    "Exportdatum: {0}" -f $exportDt.ToString('yyyy-MM-dd HH:mm', $svCulture) | Write-Host
}
else {
    Write-Warning ("Exportdatum kunde inte tolkas (råvärde: '{0}')" -f $data.export_date)
}
Write-Host ""

# Lista alla användare som inte loggat in på 30+ dagar
$threshold = $now.AddDays(-30)
$inactive = $data.users | ForEach-Object {
    $dt = Parse-Date $_.lastLogon
    if ($dt -and $dt -lt $threshold) {
        [pscustomobject]@{
            Namn                 = $_.displayName
            Konto                = $_.samAccountName
            SenastInloggad       = $dt.ToString('yyyy-MM-dd HH:mm', $svCulture)
            DagarSedanInloggning = [int]($now - $dt).TotalDays
            Department           = $_.department
            Enabled              = $_.enabled
        }
    }
}

"Användare inaktiva (30+ dagar):" | Write-Host
if ($inactive) {
    $inactive | Sort-Object DagarSedanInloggning -Descending |
    Format-Table -AutoSize | Out-Host
}
else {
    "Inga inaktiva användare över 30 dagar." | Write-Host
}
Write-Host ""

# Räkna antal användare per avdelning -loop-
$deptCount = @{}
foreach ($u in $data.users) {
    $key = if ($u.department) { $u.department } else { 'Okänd' }
    if ($deptCount.ContainsKey($key)) { $deptCount[$key]++ } else { $deptCount[$key] = 1 }
}

"Antal användare per avdelning:" | Write-Host
$deptCount.GetEnumerator() |
Sort-Object Name |
Format-Table @{n = 'Avdelning'; e = { $_.Name } }, @{n = 'Antal'; e = { $_.Value } } -AutoSize |
Out-Host
