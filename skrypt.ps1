#region ===========================================
#region KONFIGURACJA CSV
# ===========================================

$outputHTML   = Join-Path $PSScriptRoot "raport.html"
$csvPAG       = Join-Path $PSScriptRoot "AnalizaPAG.csv"
$csvZdarzenia = Join-Path $PSScriptRoot "Zdarzenia.csv"
$csvLista     = Join-Path $PSScriptRoot "ListaZwierzat.csv"

#endregion

#region ===========================================
#region WCZYTANIE LISTY ZWIERZĄT
# ===========================================

$lista = Import-Csv $csvLista -Delimiter ";" | Select-Object -SkipLast 2
$listaKolczyki = $lista.Kolczyk

#endregion

#region ===========================================
#region WCZYTANIE DANYCH PAG (TYLKO Z LISTY)
# ===========================================

$data = Import-Csv $csvPAG -Delimiter ";" |
        Where-Object { $listaKolczyki -contains $_.Kolczyk }

#endregion

#region ===========================================
#region MIESIĄCE (12 wstecz + 1 w przód)
# ===========================================

$today = Get-Date
$startMonth = $today.AddMonths(-12)
$endMonth   = $today.AddMonths(1)

$months = @()
$cursor = Get-Date -Year $startMonth.Year -Month $startMonth.Month -Day 1
$end    = Get-Date -Year $endMonth.Year   -Month $endMonth.Month   -Day 1

while ($cursor -le $end) {
    $months += $cursor.ToString("yyyy-MM")
    $cursor = $cursor.AddMonths(1)
}

#endregion

#region ===========================================
#region BUDOWA PIVOTU
# ===========================================

$pivot = @{}

foreach ($k in $listaKolczyki) {

    $rows  = $data | Where-Object { $_.Kolczyk -eq $k }
    $nazwa = ($lista | Where-Object { $_.Kolczyk -eq $k } | Select-Object -First 1).Nazwa

    $pivot[$k] = [ordered]@{
        Kolczyk = $k
        Nazwa   = $nazwa
    }

    foreach ($m in $months) {
        $pivot[$k][$m] = ""
    }

    foreach ($r in $rows) {

        $dt = $r."Pobranie próbki" -as [datetime]
        if (-not $dt) { continue }

        $ym = $dt.ToString("yyyy-MM")
        if ($months -contains $ym) {
            $pivot[$k][$ym] = $r.Wynik
        }
    }
}

#endregion

#region ===========================================
#region WCZYTANIE ZDARZEŃ (TYLKO Z LISTY)
# ===========================================

$zdarzenia = Import-Csv $csvZdarzenia -Delimiter ";" |
             Where-Object { $listaKolczyki -contains $_.Zwierzę }

#endregion

#region ===========================================
#region WYCELENIA (WIELOKROTNE)
# ===========================================

$wycMap = @{}

foreach ($r in $zdarzenia) {

    if ($r."Rodzaj zdarzenia" -ne "Wycielenie") { continue }

    $d = $r."Data zdarzenia" -as [datetime]
    if (-not $d) { continue }

    $k = $r.Zwierzę
    $m = $d.ToString("yyyy-MM")

    if (-not $wycMap.ContainsKey($k)) {
        $wycMap[$k] = @()
    }

    $wycMap[$k] += $m
}

#endregion

#region ===========================================
#region ZASUSZENIA (WIELOKROTNE)
# ===========================================

$zasMap = @{}

foreach ($r in $zdarzenia) {

    if ($r."Rodzaj zdarzenia" -ne "Zasuszenie") { continue }

    $d = $r."Data zdarzenia" -as [datetime]
    if (-not $d) { continue }

    $k = $r.Zwierzę
    $m = $d.ToString("yyyy-MM")

    if (-not $zasMap.ContainsKey($k)) {
        $zasMap[$k] = @()
    }

    $zasMap[$k] += $m
}

#endregion

#region ===========================================
#region NAKŁADANIE ZASUSZENIA (ANULOWANE PRZEZ WYCELENIE)
# ===========================================

foreach ($k in $pivot.Keys) {

    if (-not $zasMap.ContainsKey($k)) { continue }

    foreach ($start in ($zasMap[$k] | Sort-Object)) {

        $end = $null
        if ($wycMap.ContainsKey($k)) {
            $end = $wycMap[$k] |
                   Where-Object { $_ -ge $start } |
                   Sort-Object |
                   Select-Object -First 1
        }

        foreach ($m in $months) {

            if ($m -lt $start) { continue }
            if ($end -and $m -gt $end) { break }

            # NIE NADPISUJ wycielenia
            if ($pivot[$k][$m] -ne "WYCIELENIE") {
                $pivot[$k][$m]


#region ===========================================
#region NAKŁADANIE WYCELEN (WSZYSTKICH)
# ===========================================

foreach ($k in $listaKolczyki) {

    if (-not $wycMap.ContainsKey($k)) { continue }

    foreach ($m in $wycMap[$k]) {
        if ($months -contains $m) {
            $pivot[$k][$m] = "WYCIELENIE"
        }
    }
}

#endregion

#region ===========================================
#region GENEROWANIE HTML
# ===========================================

$css = @"
<style>
body {
    font-family: Segoe UI, Arial, sans-serif;
    background: #0f172a;
    color: #e5e7eb;
}
table {
    border-collapse: collapse;
    width: 100%;
    font-size: 13px;
    table-layout: fixed;
}
th, td {
    border: 1px solid #334155;
    padding: 4px 6px;
    text-align: center;
    white-space: nowrap;
}
th {
    background: #1e293b;
    position: sticky;
    top: 0;
    z-index: 2;
}
td.left {
    text-align: left;
    font-weight: 600;
}
.cielna     { background: #14532d; }
.niecielna  { background: #7f1d1d; }
.zasuszona  { background: #1e40af; }
.wycielenie { background: #c2410c; font-weight: 700; }
</style>
"@

$html = "<html><head><meta charset='UTF-8'>$css</head><body><table>"
$html += "<thead><tr><th>Lp</th><th>Kolczyk</th><th>Nazwa</th>"

foreach ($m in $months) { $html += "<th>$m</th>" }

$html += "</tr></thead><tbody>"

$lp = 1
foreach ($row in $pivot.Values) {

    $html += "<tr>"
    $html += "<td>$lp</td>"
    $html += "<td class='left'>$($row.Kolczyk)</td>"
    $html += "<td class='left'>$($row.Nazwa)</td>"

    foreach ($m in $months) {

        $v = $row[$m]
        $class = ""

        switch ($v) {
            "Cielna"     { $class = "cielna" }
            "Niecielna"  { $class = "niecielna" }
            "ZASUSZONA"  { $class = "zasuszona" }
            "WYCIELENIE" { $class = "wycielenie" }
        }

        $html += "<td class='$class'>$v</td>"
    }

    $html += "</tr>"
    $lp++
}

$html += "</tbody></table></body></html>"
$html | Out-File $outputHTML -Encoding UTF8

#endregion


<#
# =========================
# GIT PUSH (AUTO-DEPLOY)
# =========================

Push-Location $PSScriptRoot

git add .
git commit -m "Auto update report $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
git push

Pop-Location
#>