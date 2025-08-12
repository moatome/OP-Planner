# Schichtplan Import Format

## Übersicht
Die Schichtplan-Import-Funktion verarbeitet Excel-Dateien mit täglichen Schichtplänen. Diese Dateien enthalten Spalten für verschiedene Schichttypen, in denen die zugewiesenen Mitarbeiter aufgelistet sind.

## Erwartete Spaltenstruktur
Die Excel-Datei sollte Spalten für die folgenden Schichttypen enthalten:

### Unterstützte Schichttypen
- **Bereitschaften (BD)**: Bereitschaftsdienste
- **Rufdienste (RD)**: Rufdienste
- **Frühdienste (Früh)**: Frühdienste
- **Zwischendienste/Mitteldienste (Mittel)**: Zwischen- oder Mitteldienste
- **Spätdienste (Spät)**: Spätdienste

### Spaltenname-Variationen
Das System erkennt verschiedene Schreibweisen für die Spalten:
- **Bereitschaften**: `bereitschaften`, `bd`, `bereitschaft`
- **Rufdienste**: `rufdienste`, `rd`, `rufdienst`, `ruf`
- **Frühdienste**: `frühdienste`, `früh`, `fruh`, `frühdienst`, `early`
- **Zwischendienste**: `zwischendienste`, `mitteldienste`, `mittel`, `zwischen`, `middle`
- **Spätdienste**: `spätdienste`, `spät`, `spaet`, `spätdienst`, `late`

## Format der Personalzuweisungen
In jeder Spalte können mehrere Mitarbeiter aufgelistet werden, getrennt durch Zeilenwechsel:

### Erwartetes Format
```
Nachname, Vorname (Abteilung) (Code)
Nachname, Vorname (Abteilung) (Code)
...
```

### Beispiel: Spalte "Bereitschaften (BD)"
```
Findeisen, Sarah (OP KLD) (VB)
Klatt, Olga (OP KLD) (BD)
Möbus, Maximilian (OP KLD) (VB)
Pochalla, Albine (OP KLD) (BD)
Voigt, Svenja (OP KLD) (BD)
```

## Beispiel-Excel-Struktur

| Bereitschaften (BD) | Rufdienste (RD) | Frühdienste (Früh) | Zwischendienste/Mitteldienste (Mittel) | Spätdienste (Spät) |
|---------------------|-----------------|-------------------|---------------------------------------|-------------------|
| Findeisen, Sarah (OP KLD) (VB)<br>Klatt, Olga (OP KLD) (BD)<br>Möbus, Maximilian (OP KLD) (VB) | Block, Aylien (OP KLD) (KR)<br>Herden, Thomas (OP KLD) (KR) | Biere, Annegret (OP KLD) (OK)<br>Lackmann, Tatjana (OP KLD) (V) | Beyer, Johanna (OP KLD) (I,B2)<br>Block, Aylien (OP KLD) (I) | Bongardt-Kleiner, Susanne (OP KLD) (SK)<br>Gaudl, Meinolf (OP KLD) (S) |

## Verarbeitungslogik

### 1. Excel-Verarbeitung
- Das System liest alle Arbeitsblätter der Excel-Datei
- Erkennt Spalten anhand der Schichttyp-Namen
- Extrahiert Mitarbeiternamen aus mehrzeiligen Zellen

### 2. Name-Parsing
- Trennt "Nachname, Vorname" Format
- Erstellt vollständigen Namen: "Vorname Nachname"
- Ignoriert Zusatzinformationen in Klammern

### 3. Verfügbarkeits-Update
- **Zugewiesene Mitarbeiter**: Werden auf "verfügbar" gesetzt mit entsprechendem Schichttyp
- **Nicht zugewiesene Mitarbeiter**: Werden auf "nicht verfügbar" gesetzt
- **Planungsansicht**: Zeigt nur verfügbare Mitarbeiter in der Seitenleiste

## Automatische Funktionen
- **Name-Normalisierung**: Automatische Konvertierung von "Nachname, Vorname" zu "Vorname Nachname"
- **Verfügbarkeits-Aktualisierung**: Automatische Anpassung aller Personalverfügbarkeiten
- **Seitenleisten-Filter**: Nur verfügbare Mitarbeiter werden im Planner angezeigt
- **Schichttyp-Zuordnung**: Verfügbarkeitskategorie wird auf Schichttyp gesetzt

## Dateianforderungen
- **Format**: Excel-Dateien (.xlsx oder .xls)
- **Struktur**: Mindestens eine Spalte mit erkanntem Schichttyp
- **Inhalt**: Mehrzeilige Zellen mit Mitarbeiternamen im erwarteten Format

## Nach dem Import
Nach erfolgreichem Import:
1. Alle Mitarbeiter-Verfügbarkeiten werden aktualisiert
2. Nur verfügbare Mitarbeiter erscheinen in der Planner-Seitenleiste
3. Schichttyp wird als Verfügbarkeitskategorie gesetzt
4. Nicht zugewiesene Mitarbeiter werden als "nicht verfügbar" markiert

## Validierung
Das System überprüft automatisch:
- Eindeutige Namen (keine Duplikate)
- Mindestlänge für Namen (2 Zeichen)
- Gültige Gruppenzuordnungen
- Datenformat-Konsistenz

## Unterstützte Dateiformate
- Excel (.xlsx)
- Excel Legacy (.xls)

## Fehlerbehandlung
- Ungültige Einträge werden in der Vorschau angezeigt
- Detaillierte Fehlermeldungen für jede Zeile
- Möglichkeit, nur gültige Daten zu importieren
