# Deutsche Bahn Dienstplan zu Kalender Konverter


![Alt-Text](https://github.com/TetsuGuy/Deutsche-Bahn-Dienstplan-Konverter/blob/master/DB%20Kalender%20Toolk.png)

Dieses Projekt ist eine Windows-Anwendung, die einen Dienstplan der Deutschen Bahn aus einer Excel-Datei (.xlsx) in eine Kalenderdatei (.ics) umwandelt.

## Funktionen
- Auswahl einer Excel-Datei mit Dienstplandaten
- Automatische Extraktion von Diensten und Umwandlung in Kalendereinträge
- Speicherung des generierten Kalenders als .ics-Datei
- Benutzerfreundliche GUI für einfache Bedienung

## Voraussetzungen
- Windows-Betriebssystem
- .NET Framework (WPF-Unterstützung erforderlich)
- [EPPlus](https://www.epplussoftware.com/) für die Verarbeitung von Excel-Dateien

## Installation
1. Repository klonen:
   ```sh
   git clone https://github.com/deinbenutzername/deutsche-bahn-dienstplan-konverter.git
   ```
2. Abhängigkeiten installieren (falls nicht vorhanden):
   ```sh
   dotnet add package EPPlus
   ```
3. Projekt in Visual Studio oder einer anderen .NET-kompatiblen IDE öffnen und kompilieren.

## Nutzung
1. Anwendung starten
2. Eine Excel-Datei mit einem Dienstplan auswählen
3. Speicherort für die .ics-Kalenderdatei wählen
4. Datei speichern und in den Kalender importieren

## Format der Excel-Datei
Die Excel-Datei sollte folgendes Format haben:
- Erste Spalte: Monat und Jahr (z.B. `01,2025`)
- Danach Spalten für jeden Tag des Monats mit Dienstinformationen
- Schichtarten in einer separaten Zeile unter den Datumseinträgen

## Beispiel für generierte .ics-Datei
```plaintext
BEGIN:VCALENDAR
VERSION:2.0
CALSCALE:GREGORIAN
BEGIN:VEVENT
SUMMARY:Frühschicht
DTSTART;VALUE=DATE:20250203
DTEND;VALUE=DATE:20250204
END:VEVENT
END:VCALENDAR
```

## Lizenz
Dieses Projekt steht unter der MIT-Lizenz.

## Autor
R.H.

## Kontakt
Falls du Fragen oder Vorschläge hast, erstelle ein Issue oder kontaktiere mich über GitHub.

