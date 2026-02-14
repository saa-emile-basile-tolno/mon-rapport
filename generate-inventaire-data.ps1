$ErrorActionPreference = 'Stop'

$excelPath = "C:\CONSERVATION\projet de stage 2\inventaire Departement Sciences Economiques & Gestion.xlsx"
$outPath = "C:\CONSERVATION\projet de stage 2\inventaire-data.js"

# NOTE: This script is executed with Windows PowerShell 5.1 in this project.
# PS 5.1 can misread UTF-8 scripts without BOM, so we avoid relying on accented
# literals by normalizing text (remove diacritics + punctuation) before matching.

function Remove-Diacritics([string]$text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    $normalized = $text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalized.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Normalize([object]$v) {
    if ($null -eq $v) { return '' }
    return ($v.ToString()).Trim()
}

function NormalizeKey([object]$v) {
    $s = Normalize $v
    if (-not $s) { return '' }

    $s = Remove-Diacritics $s
    $s = $s.ToLowerInvariant()
    $s = ($s -replace '[^a-z0-9]+', ' ').Trim()
    $s = ($s -replace '\s+', ' ')
    return $s
}

function Get-RowCells($row) {
    $cells = [ordered]@{}
    $props = $row.PSObject.Properties |
        Where-Object { $_.Name -match '^P\d+$' } |
        Sort-Object { [int]($_.Name.Substring(1)) }

    foreach ($p in $props) {
        $cells[$p.Name] = Normalize $p.Value
    }
    return $cells
}

function Get-HeaderMap($cells) {
    $map = @{}

    foreach ($col in $cells.Keys) {
        $k = NormalizeKey $cells[$col]
        if (-not $k) { continue }

        if (-not $map.num -and ($k -eq 'n' -or $k -eq 'no' -or $k -like 'n *')) { $map.num = $col; continue }
        if (-not $map.designation -and $k -match 'designation') { $map.designation = $col; continue }
        if (-not $map.marque -and $k -match 'marque') { $map.marque = $col; continue }
        if (-not $map.inventaire -and $k -match 'inventaire') { $map.inventaire = $col; continue }
    }

    if ($map.designation -and $map.inventaire) {
        if (-not $map.num) { $map.num = 'P1' }
        if (-not $map.marque) { $map.marque = 'P3' }
        return $map
    }

    return $null
}

function Get-SheetId([string]$sheetNameTrim) {
    $name = NormalizeKey $sheetNameTrim

    if ($name -match '^salle\s*(\d+)$') { return "salle$($Matches[1])" }
    if ($name -match '^labo\s*(\d+)$') { return "labo$($Matches[1])" }
    if ($name -eq 'eep') { return 'eep' }
    if ($name -eq 'salle excellence') { return 'excellence' }
    if ($name -eq 'salle enseignant') { return 'enseignant' }
    if ($name -eq 'qualite') { return 'qualite' }
    if ($name -eq 'l21') { return 'l21' }
    if ($name -eq 'b11') { return 'b11' }
    if ($name -eq 'mouna daaafousse') { return 'mouna' }
    if ($name -eq 'awatef chamaghi') { return 'awatef' }
    if ($name -like 'chef de departement gestion*') { return 'chef' }

    # fallback: keep letters/digits only
    $fallback = ($name -replace '[^a-z0-9]+','')
    if ([string]::IsNullOrWhiteSpace($fallback)) { $fallback = 'sheet' }
    return $fallback
}

function Get-SheetLabel([string]$sheetNameTrim) {
    # Keep labels ASCII to avoid encoding surprises.
    $name = $sheetNameTrim.Trim()
    $key = NormalizeKey $sheetNameTrim

    if ($key -match '^salle\s*(\d+)$') { return "Salle $($Matches[1])" }
    if ($key -match '^labo\s*(\d+)$') { return "Laboratoire $($Matches[1])" }
    if ($key -eq 'eep') { return 'EEP' }
    if ($key -eq 'salle excellence') { return 'Salle Excellence' }
    if ($key -eq 'salle enseignant') { return 'Salle Enseignant' }
    if ($key -eq 'qualite') { return 'Bureau Qualite' }
    if ($key -eq 'l21') { return 'L21' }
    if ($key -eq 'b11') { return 'B11' }
    if ($key -eq 'mouna daaafousse') { return 'Bureau Mouna Daaafousse' }
    if ($key -eq 'awatef chamaghi') { return 'Bureau Awatef Chamaghi' }
    if ($key -like 'chef de departement gestion*') { return 'Chef de Departement Gestion' }

    return $name
}

function IsItemRow($cells, $colMap) {
    $numCol = if ($colMap -and $colMap.num) { $colMap.num } else { 'P1' }
    $desCol = if ($colMap -and $colMap.designation) { $colMap.designation } else { 'P2' }

    $p1 = Normalize $cells[$numCol]
    $p2 = Normalize $cells[$desCol]

    if ([string]::IsNullOrWhiteSpace($p1) -or [string]::IsNullOrWhiteSpace($p2)) { return $false }
    if ((NormalizeKey $p1) -eq 'n' -and (NormalizeKey $p2) -match 'designation') { return $false }

    # Accept numeric and also letters used in some sheets (e.g. 's')
    if ($p1 -match '^[0-9]+$') { return $true }
    if ($p1 -match '^[a-zA-Z]+$') { return $true }

    return $false
}

# Ensure ImportExcel is available
try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    throw "Le module PowerShell 'ImportExcel' n'est pas disponible. Relancez PowerShell puis exécutez: Set-ExecutionPolicy -Scope Process Bypass; Install-Module ImportExcel -Scope CurrentUser"
}

$sheets = Get-ExcelSheetInfo -Path $excelPath

$data = [ordered]@{}
$order = New-Object System.Collections.Generic.List[string]

foreach ($sheet in $sheets) {
    $sheetNameActual = $sheet.Name
    $sheetNameTrim = $sheetNameActual.Trim()
    $id = Get-SheetId $sheetNameTrim
    $label = Get-SheetLabel $sheetNameTrim

    $rows = Import-Excel -Path $excelPath -WorksheetName $sheetNameActual -NoHeader

    $currentCategory = $null
    $currentColMap = $null
    $meubles = New-Object System.Collections.Generic.List[object]
    $info = New-Object System.Collections.Generic.List[object]

    foreach ($row in $rows) {
        $cells = Get-RowCells $row
        $rowKey = NormalizeKey (($cells.Values) -join ' ')

        if ($rowKey -match '\bmeuble(s)?\b') {
            $currentCategory = 'meubles'
            $currentColMap = $null
            continue
        }
        if ($rowKey -match '\bmateriel(s)?\s+informatique(s)?\b') {
            $currentCategory = 'informatique'
            $currentColMap = $null
            continue
        }

        if (-not $currentCategory) { continue }

        $maybeMap = Get-HeaderMap $cells
        if ($maybeMap) {
            $currentColMap = $maybeMap
            continue
        }

        if (-not (IsItemRow $cells $currentColMap)) { continue }

        $numCol = if ($currentColMap -and $currentColMap.num) { $currentColMap.num } else { 'P1' }
        $desCol = if ($currentColMap -and $currentColMap.designation) { $currentColMap.designation } else { 'P2' }
        $marCol = if ($currentColMap -and $currentColMap.marque) { $currentColMap.marque } else { 'P3' }
        $invCol = if ($currentColMap -and $currentColMap.inventaire) { $currentColMap.inventaire } else { 'P4' }

        $num = Normalize $cells[$numCol]
        $designation = Normalize $cells[$desCol]
        $marqueRaw = Normalize $cells[$marCol]
        $inventaireRaw = Normalize $cells[$invCol]

        $marque = if ($marqueRaw) { $marqueRaw } else { '-' }
        $inventaire = if ($inventaireRaw) { $inventaireRaw } else { '-' }

        $item = [ordered]@{
            num = $num
            designation = $designation
            marque = $marque
            inventaire = $inventaire
        }

        if ($currentCategory -eq 'meubles') { $meubles.Add($item) } else { $info.Add($item) }
    }

    $data[$id] = [ordered]@{
        label = $label
        categories = [ordered]@{
            meubles = $meubles
            informatique = $info
        }
    }

    $order.Add($id)
}

# Build a JS file: window.INVENTAIRE_DATA / window.INVENTAIRE_ORDER
$json = $data | ConvertTo-Json -Depth 6
$orderJson = $order | ConvertTo-Json

$js = @()
$js += "// Auto-genere depuis: inventaire Departement Sciences Economiques & Gestion.xlsx"
$js += "// Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$js += "window.INVENTAIRE_DATA = $json;"
$js += "window.INVENTAIRE_ORDER = $orderJson;"

$js -join "`n" | Out-File -FilePath $outPath -Encoding utf8

Write-Host "OK -> $outPath"