<#
.SYNOPSIS
    Vergleicht zwei CSV-Dateien anhand einer eindeutigen ID-Spalte,
    optional mit einer Spaltenmapping-Datei fuer unterschiedlich benannte Spalten.

.DESCRIPTION
    Ermittelt:
      - Datensaetze, die nur in Datei 1 vorkommen
      - Datensaetze, die nur in Datei 2 vorkommen
      - Datensaetze mit gleicher ID, aber abweichenden Werten in den (gemappten) Spalten
      - Datensaetze, die in beiden Dateien identisch sind

    Ohne -Mapping werden alle Spalten mit identischem Namen in beiden Dateien
    verglichen (Fallback-Verhalten). Mit -Mapping werden NUR die im Mapping
    aufgefuehrten Spaltenpaare verglichen; in der Ausgabe wird jeweils der
    Spaltenname aus Datei1 angezeigt.

    Kompatibel mit Windows PowerShell 5.1 und PowerShell 7 (pwsh).

.PARAMETER Datei1
    Pfad zur ersten CSV-Datei.

.PARAMETER Datei2
    Pfad zur zweiten CSV-Datei.

.PARAMETER IdSpalte
    Name der ID-Spalte in Datei1 (Standard: ID). Der Name der ID-Spalte in
    Datei2 wird - falls -Mapping angegeben ist - aus dem Mapping ermittelt.

.PARAMETER Mapping
    Optional: Pfad zu einer CSV-Datei mit den Spalten "Datei1" und "Datei2",
    die festlegt, welche Spalten aus Datei1 mit welchen Spalten aus Datei2
    verglichen werden sollen. Nur im Mapping enthaltene Spalten werden
    verglichen. Beispiel-Inhalt (Semikolon-getrennt):

        Datei1;Datei2
        ID;ID
        Name;Bezeichnung
        Wert;Betrag
        Status;Zustand

.PARAMETER Ausgabe
    Optional: Pfad, unter dem das Diff-Ergebnis als CSV gespeichert werden soll.

.PARAMETER Delimiter
    Trennzeichen der CSV- und Mapping-Dateien (Standard: Semikolon).

.EXAMPLE
    .\csv_vergleich.ps1 -Datei1 .\datei1.csv -Datei2 .\datei2.csv

.EXAMPLE
    .\csv_vergleich.ps1 -Datei1 .\datei1.csv -Datei2 .\datei2.csv -Mapping .\mapping.csv -Ausgabe .\diff.csv
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Datei1,

    [Parameter(Mandatory = $true)]
    [string]$Datei2,

    [string]$IdSpalte = "ID",

    [string]$Mapping,

    [string]$Ausgabe,

    [string]$Delimiter = ";"
)

function Get-CsvAlsHashtable {
    param(
        [string]$Pfad,
        [string]$IdSpalte,
        [string]$Delimiter
    )

    if (-not (Test-Path -LiteralPath $Pfad)) {
        Write-Error "Datei nicht gefunden: $Pfad"
        exit 1
    }

    $zeilen = Import-Csv -LiteralPath $Pfad -Delimiter $Delimiter

    if ($zeilen.Count -eq 0) {
        Write-Error "Datei ist leer oder konnte nicht gelesen werden: $Pfad"
        exit 1
    }

    $spalten = $zeilen[0].PSObject.Properties.Name
    if ($spalten -notcontains $IdSpalte) {
        Write-Error "Spalte '$IdSpalte' nicht in $Pfad gefunden. Vorhandene Spalten: $($spalten -join ', ')"
        exit 1
    }

    $tabelle = @{}
    foreach ($zeile in $zeilen) {
        $tabelle[$zeile.$IdSpalte] = $zeile
    }
    return $tabelle
}

function Get-Mapping {
    param(
        [string]$Pfad,
        [string]$Delimiter
    )

    if (-not (Test-Path -LiteralPath $Pfad)) {
        Write-Error "Mapping-Datei nicht gefunden: $Pfad"
        exit 1
    }

    $zeilen = Import-Csv -LiteralPath $Pfad -Delimiter $Delimiter

    if ($zeilen.Count -eq 0) {
        Write-Error "Mapping-Datei ist leer oder konnte nicht gelesen werden: $Pfad"
        exit 1
    }

    $spalten = $zeilen[0].PSObject.Properties.Name
    if (($spalten -notcontains "Datei1") -or ($spalten -notcontains "Datei2")) {
        Write-Error "Mapping-Datei muss die Spalten 'Datei1' und 'Datei2' enthalten. Vorhandene Spalten: $($spalten -join ', ')"
        exit 1
    }

    return $zeilen
}

function Compare-CsvDaten {
    param(
        [hashtable]$Daten1,
        [hashtable]$Daten2,
        [string]$IdSpalte,
        $VergleichsSpalten
    )

    $ids2 = [System.Collections.Generic.HashSet[string]]::new([string[]]$Daten2.Keys)
    $ids1 = [System.Collections.Generic.HashSet[string]]::new([string[]]$Daten1.Keys)

    $nurIn1 = $Daten1.Keys | Where-Object { -not $ids2.Contains($_) } | Sort-Object
    $nurIn2 = $Daten2.Keys | Where-Object { -not $ids1.Contains($_) } | Sort-Object
    $gemeinsam = $Daten1.Keys | Where-Object { $ids2.Contains($_) } | Sort-Object

    $abweichend = [ordered]@{}
    $identisch = @()

    foreach ($id in $gemeinsam) {
        $zeile1 = $Daten1[$id]
        $zeile2 = $Daten2[$id]
        $unterschiede = [ordered]@{}

        if ($VergleichsSpalten) {
            # Nur die im Mapping definierten Spaltenpaare vergleichen.
            # In der Ausgabe wird der Spaltenname aus Datei1 verwendet.
            foreach ($paar in $VergleichsSpalten) {
                $wert1 = $zeile1.($paar.Datei1)
                $wert2 = $zeile2.($paar.Datei2)
                if ($wert1 -ne $wert2) {
                    $unterschiede[$paar.Datei1] = @{ Alt = $wert1; Neu = $wert2 }
                }
            }
        }
        else {
            # Kein Mapping: alle gleichnamigen Spalten vergleichen (ausser der ID-Spalte).
            foreach ($spalte in $zeile1.PSObject.Properties.Name) {
                if ($spalte -eq $IdSpalte) { continue }
                $wert1 = $zeile1.$spalte
                $wert2 = $zeile2.$spalte
                if ($wert1 -ne $wert2) {
                    $unterschiede[$spalte] = @{ Alt = $wert1; Neu = $wert2 }
                }
            }
        }

        if ($unterschiede.Count -gt 0) {
            $abweichend[$id] = $unterschiede
        }
        else {
            $identisch += $id
        }
    }

    return [PSCustomObject]@{
        NurIn1     = $nurIn1
        NurIn2     = $nurIn2
        Abweichend = $abweichend
        Identisch  = $identisch
    }
}

function Write-Bericht {
    param(
        [PSCustomObject]$Ergebnis,
        [string]$Name1,
        [string]$Name2
    )

    Write-Host "=== Vergleichsbericht: $Name1 vs. $Name2 ===" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Nur in $Name1 vorhanden ($($Ergebnis.NurIn1.Count)):" -ForegroundColor Yellow
    foreach ($id in $Ergebnis.NurIn1) {
        Write-Host "  ID $id"
    }

    Write-Host ""
    Write-Host "Nur in $Name2 vorhanden ($($Ergebnis.NurIn2.Count)):" -ForegroundColor Yellow
    foreach ($id in $Ergebnis.NurIn2) {
        Write-Host "  ID $id"
    }

    Write-Host ""
    Write-Host "Abweichende Datensaetze ($($Ergebnis.Abweichend.Count)):" -ForegroundColor Yellow
    foreach ($id in $Ergebnis.Abweichend.Keys) {
        Write-Host "  ID $id`:"
        foreach ($spalte in $Ergebnis.Abweichend[$id].Keys) {
            $diff = $Ergebnis.Abweichend[$id][$spalte]
            Write-Host "      $spalte`: '$($diff.Alt)' -> '$($diff.Neu)'"
        }
    }

    Write-Host ""
    Write-Host "Identische Datensaetze: $($Ergebnis.Identisch.Count)" -ForegroundColor Green
}

function Export-DiffCsv {
    param(
        [PSCustomObject]$Ergebnis,
        [string]$Pfad,
        [string]$IdSpalte,
        [string]$Delimiter
    )

    $ausgabeZeilen = @()

    foreach ($id in $Ergebnis.NurIn1) {
        $ausgabeZeilen += [PSCustomObject]@{
            $IdSpalte = $id
            Status    = "nur_in_datei1"
            Details   = ""
        }
    }

    foreach ($id in $Ergebnis.NurIn2) {
        $ausgabeZeilen += [PSCustomObject]@{
            $IdSpalte = $id
            Status    = "nur_in_datei2"
            Details   = ""
        }
    }

    foreach ($id in $Ergebnis.Abweichend.Keys) {
        $teile = @()
        foreach ($spalte in $Ergebnis.Abweichend[$id].Keys) {
            $diff = $Ergebnis.Abweichend[$id][$spalte]
            $teile += "$spalte`: '$($diff.Alt)' -> '$($diff.Neu)'"
        }
        $ausgabeZeilen += [PSCustomObject]@{
            $IdSpalte = $id
            Status    = "abweichend"
            Details   = ($teile -join "; ")
        }
    }

    $ausgabeZeilen | Export-Csv -LiteralPath $Pfad -Delimiter $Delimiter -NoTypeInformation -Encoding UTF8
}

# --- Hauptprogramm ---

$idSpalteDatei2 = $IdSpalte
$vergleichsSpalten = $null

if ($Mapping) {
    $mappingZeilen = Get-Mapping -Pfad $Mapping -Delimiter $Delimiter

    $idZeile = $mappingZeilen | Where-Object { $_.Datei1 -eq $IdSpalte }
    if ($idZeile) {
        $idSpalteDatei2 = $idZeile.Datei2
    }
    else {
        Write-Warning "ID-Spalte '$IdSpalte' nicht im Mapping gefunden - verwende denselben Namen '$IdSpalte' fuer Datei2."
    }

    $vergleichsSpalten = $mappingZeilen | Where-Object { $_.Datei1 -ne $IdSpalte }
}

$daten1 = Get-CsvAlsHashtable -Pfad $Datei1 -IdSpalte $IdSpalte -Delimiter $Delimiter
$daten2 = Get-CsvAlsHashtable -Pfad $Datei2 -IdSpalte $idSpalteDatei2 -Delimiter $Delimiter

# Bei aktivem Mapping: Spaltenpaare herausfiltern, deren Spalten nicht existieren,
# statt das Skript mit einem Fehler abzubrechen.
if ($vergleichsSpalten) {
    $spalten1 = ($daten1.Values | Select-Object -First 1).PSObject.Properties.Name
    $spalten2 = ($daten2.Values | Select-Object -First 1).PSObject.Properties.Name

    $vergleichsSpaltenGueltig = @()
    foreach ($paar in $vergleichsSpalten) {
        if ($spalten1 -notcontains $paar.Datei1) {
            Write-Warning "Spalte '$($paar.Datei1)' aus dem Mapping nicht in $Datei1 gefunden - wird uebersprungen."
            continue
        }
        if ($spalten2 -notcontains $paar.Datei2) {
            Write-Warning "Spalte '$($paar.Datei2)' aus dem Mapping nicht in $Datei2 gefunden - wird uebersprungen."
            continue
        }
        $vergleichsSpaltenGueltig += $paar
    }
    $vergleichsSpalten = $vergleichsSpaltenGueltig
}

$ergebnis = Compare-CsvDaten -Daten1 $daten1 -Daten2 $daten2 -IdSpalte $IdSpalte -VergleichsSpalten $vergleichsSpalten

Write-Bericht -Ergebnis $ergebnis -Name1 (Split-Path -Leaf $Datei1) -Name2 (Split-Path -Leaf $Datei2)

if ($Ausgabe) {
    Export-DiffCsv -Ergebnis $ergebnis -Pfad $Ausgabe -IdSpalte $IdSpalte -Delimiter $Delimiter
    Write-Host ""
    Write-Host "Diff-Ergebnis gespeichert unter: $Ausgabe" -ForegroundColor Cyan
}