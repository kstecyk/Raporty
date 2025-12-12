#region # ===========================================
#region KONFIGURACJA CSV
# ===========================================
<#
$outputHTML    = "C:\Users\karst\OneDrive\Documents\skrypt rozrod\raport.html"
$csvPAG        = "C:\Users\karst\OneDrive\Documents\skrypt rozrod\AnalizaPAG.csv"
$csvZdarzenia = "C:\Users\karst\OneDrive\Documents\skrypt rozrod\Zdarzenia.csv"
$csvLista      = "C:\Users\karst\OneDrive\Documents\skrypt rozrod\ListaZwierzat.csv"
#>
$outputHTML   = Join-Path $PSScriptRoot "raport.html"
$csvPAG       = Join-Path $PSScriptRoot "AnalizaPAG.csv"
$csvZdarzenia = Join-Path $PSScriptRoot "Zdarzenia.csv"
$csvLista     = Join-Path $PSScriptRoot "ListaZwierzat.csv"

# ===========================================
#region WCZYTANIE LISTY ZWIERZĄT
# ===========================================
$lista = Import-Csv $csvLista -Delimiter ";" |  Select-Object -SkipLast 2

# ===========================================
#region WCZYTANIE DANYCH PAG
# ===========================================
$data = Import-Csv $csvPAG -Delimiter ";"

# ===========================================
#region MIESIĄCE: PUSTY PRZED + DANE + PUSTY PO
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

# ===========================================
#region BUDOWA PIVOTU
# ===========================================
$pivot = @{}

foreach ($k in $lista.Kolczyk) {

    # wszystkie wpisy PAG dla danego zwierzęcia
    $rows = $data | Where-Object { $_.Kolczyk -eq $k }

    # nazwa (jeśli brak danych, będzie $null)
    $nazwa = ($lista | Where-Object { $_.Kolczyk -eq $k } | Select-Object -First 1).Nazwa

    # inicjalizacja wiersza
    $pivot[$k] = [ordered]@{
        Kolczyk = $k
        Nazwa   = $nazwa
    }

    # puste miesiące
    foreach ($m in $months) {
        $pivot[$k][$m] = ""
    }

    # wypełnianie wynikami PAG
    foreach ($r in $rows) {

        $dt = $r."Pobranie próbki" -as [datetime]
        if (-not $dt) { continue }

        $ym = $dt.ToString("yyyy-MM")

        if ($months -contains $ym) {
            $pivot[$k][$ym] = $r.Wynik
        }
    }
}
# ===========================================
#region WCZYTANIE WYCIELEŃ
# ===========================================
#region WCZYTANIE WYCIELEŃ
$wyc = Import-Csv $csvZdarzenia -Delimiter ";"
$wycMap = @{}

foreach ($r in $wyc) {
    if ($r."Rodzaj zdarzenia" -eq "Wycielenie") {
        $d = $r."Data zdarzenia" -as [datetime]
        if ($d) {
            $wycMap[$r.Zwierzę] = $d.ToString("yyyy-MM")
        }
    }
}

#endregion


# ===========================================
#region WCZYTANIE ZASUSZEŃ
# ===========================================
$zas = Import-Csv $csvZdarzenia -Delimiter ";"
$zasMap = @{}

foreach ($r in $zas) {
    if ($r."Rodzaj zdarzenia" -eq "Zasuszenie") {
        $d = $r."Data zdarzenia" -as [datetime]
        $zasMap[$r.Zwierzę] = $d.ToString("yyyy-MM")

    }
}

#region NAKŁADANIE ZASUSZENIA

foreach ($k in $pivot.Keys) {

    if (-not $zasMap.ContainsKey($k)) { continue }

    $start = $zasMap[$k]
    $end   = $wycMap[$k]  # może być $null

    foreach ($m in $months) {

        if ($m -lt $start) { continue }

        # jeśli jest wycielenie koniec zasuszenia
        if ($end -and $m -ge $end) { break }

        # Wypelnij statusem Zasuszona
        if ($pivot[$k][$m] -eq "") {
            $pivot[$k][$m] = "ZASUSZONA"
        }
    }
}
#endregion

#region NAKŁADANIE WYCIELEŃ

foreach ($k in $pivot.Keys) {

    if (-not $wycMap.ContainsKey($k)) { continue }

    $m = $wycMap[$k]

    if ($months -contains $m) {
        $pivot[$k][$m] = "WYCIELENIE"
    }
}

#endregion

#region GENEROWANIE HTML

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

$html = "<html><head><meta charset='UTF-8'>$css</head><body>"
$html += "<table>"

# ===== NAGŁÓWEK =====
$html += "<thead>"
$html += "<tr><th>Lp</th><th>Kolczyk</th><th>Nazwa</th>"

foreach ($m in $months) {
    $html += "<th>$m</th>"
}

$html += "</tr>"
$html += "</thead>"

# ===== BODY =====
$html += "<tbody>"

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

$html += "</tbody>"
$html += "</table></body></html>"

$html | Out-File $outputHTML -Encoding UTF8

#endregion

# =========================
# GIT PUSH (AUTO-DEPLOY)
# =========================

Push-Location $PSScriptRoot

git add .
git commit -m "Auto update report $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
git push

Pop-Location
