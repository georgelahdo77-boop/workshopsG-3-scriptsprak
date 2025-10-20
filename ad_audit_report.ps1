# ad_audit_report.ps1
# Försöka tysta datumparsning eftersom jag får denna varning "Datum kunde inte tolkas"
# export datum och listor (Count)

# Tyst datumtolkning (ISO + US), returnerar $null vid misslyckande
function Parse-Date {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $sv = [System.Globalization.CultureInfo]'sv-SE'; $us = [System.Globalization.CultureInfo]'en-US'
    $f = 'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-dd', 'MM/dd/yyyy HH:mm:ss', 'M/d/yyyy H:mm:ss', 'MM/dd/yyyy', 'M/d/yyyy'
    foreach ($fmt in $f) { foreach ($c in @($sv, $us)) { try { return [datetime]::ParseExact($s, $fmt, $c, $styles) } catch {} } }
    foreach ($c in @($sv, $us)) { try { return [datetime]::Parse($s, $c) } catch {} }
    return $null
}

# Läs JSON
try { $d = Get-Content ".\ad_export.json" -Encoding UTF8 -Raw -ErrorAction Stop | ConvertFrom-Json }
catch { Write-Error "Kunde inte läsa ad_export.json: $($_.Exception.Message)"; exit 1 }

$now = Get-Date
$expDate = Parse-Date $d.export_date

# bygg som arrayer med @() så .Count alltid fungerar
$usersByDept = $d.users | Group-Object department | Sort-Object Name

$expiringUsers = @(
    $d.users | ForEach-Object {
        $dt = Parse-Date $_.accountExpires
        if ($dt -and $dt -ge $now -and $dt -le $now.AddDays(30)) {
            [pscustomobject]@{displayName = $_.displayName; samAccountName = $_.samAccountName;
                GårUT = $dt.ToString('yyyy-MM-dd', $culture); Dagarkvar = [int][math]::Ceiling(($dt - $now).TotalDays)
            }
        }
    } | Sort-Object Dagarkvar
)

$staleComputers = @(
    $d.computers | ForEach-Object {
        $dt = Parse-Date $_.lastLogon
        if ($dt -and $dt -lt $now.AddDays(-30)) {
            [pscustomobject]@{name = $_.name; site = $_.site;
                SenastSedd = $dt.ToString('yyyy-MM-dd HH:mm', $culture); DagarSedan = [int]($now - $dt).TotalDays
            }
        }
    } | Sort-Object DagarSedan -Descending
)
    
$oldpasswords = @(
    $d.users | ForEach-Object {
        $dt = Parse-Date $_.passwordLastSet
        if ($dt -and -not $_.passwordNeverExpires -and ($now - $dt).TotalDays -gt 90) {
            [pscustomobject]@{displayName = $_.displayName; samAccountName = $_.samAccountName;
                SenastBytt = $dt.ToString('yyyy-MM-dd', $culture); DagarSedan = [int]($now - $dt).TotalDays
            }
        }
    } | Sort-Object DagarSedan -Descending
)

# Rapport here-string
$report = @"
ACTIVE DIRECTORY AUDIT
======================

Domän:         $($d.domain)
Forest:        $($d.forest)
Exportdatum:   $(if($expDate){$expDate.ToString('yyyy-MM-dd HH:mm',$culture)}else{$d.export_date})
Genererat:     $($now.ToString('yyyy-MM-dd HH:mm',$culture))

Översikt
--------
Totala användare:  $($d.users.Count)
Totala datorer:    $($d.computers.Count)

Executive Summary (varningar)
-----------------------------
- konton som löper ut inom 30 dagar: $($expiringUsers.Count)
- Datorer som inte setts på 30+ dagar: $($staleComputers.Count)
- Användare med lösenord äldre än 90 dagar: $($oldpasswords.Count)

Detaljer: konton som löper ut ≤30 dagar
---------------------------------------
$(
 if($expiringUsers.Count){ ($expiringUsers | Format-Table -AutoSize | Out-String).Trim() } else { "Inga kommande utgångsdatum inom 30 dagar." }
)

Detaljer: Datorer ej sedda på 30+ dagar
---------------------------------------
$(
 if($staleComputers.Count){ ($staleComputers | Format-Table -AutoSize | Out-String).Trim() } else { "Inga datorer äldre än 30 dagar sedan senaste inloggning." }
)

Detaljer: Lösenord äldre än 90 dagar
------------------------------------
$(
 if($oldpasswords.Count){ ($oldpasswords | Format-Table -AutoSize | Out-String).Trim() } else { "Inga lösenord äldre än 90 dagar." }
)

Användare per avdelning
-----------------------------
$(
 ($usersByDept | Select-Object @{n='Avdelning';e={$_.Name}}, @{n='Antal';e={$_.Count}} |
  Format-Table -AutoSize | Out-String).Trim()
)

(Rapport skapad av George i filen ad_audit_report.ps1)
"@

$path = ".\ad_audit_report.txt"
$report | Out-File $path -Encoding UTF8
Write-Host "Rapport sparad: $path"

