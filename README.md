# Excel Editor

Python-Paket zum Lesen und Bearbeiten von Excel-Dateien (RISE Planungsexcel).  
Formatierung (Farben, Schrift, Rahmen) bleibt beim Bearbeiten vollständig erhalten.

---

## Installation

### Option A: Als Python-Paket (Linux / WSL / Windows mit Python)

```bash
# Repository clonen, dann im Projektordner:
cd excel-editor

# Venv aktivieren (präferiert – isoliert vom System-Python)
source .excel-editor/bin/activate      # Linux/macOS/WSL
# .excel-editor\Scripts\activate       # Windows CMD
# .excel-editor\Scripts\Activate.ps1   # Windows PowerShell

pip install -e .

# → Der Befehl excel-editor ist jetzt verfügbar, solange das venv aktiv ist.
```

```bash
# Ohne venv: Direkt ins User-Python installieren
pip install --user -e .

# Linux/WSL: Danach einmalig PATH erweitern:
echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc && source ~/.bashrc
```

### Option B: Als standalone EXE (Windows, kein Python nötig)

Einmalig unter **Windows** (nicht WSL) ausführen:

```powershell
# Windows PowerShell (als normaler Benutzer)
cd excel-editor
.\build_exe.ps1
```

Das Skript installiert automatisch PyInstaller, baut die EXE und legt sie unter `dist\excel-editor.exe` ab.  
Die fertige EXE funktioniert **ohne Python-Installation** auf jedem Windows-Rechner.

```powershell
# Verwendung nach dem Build:
.\dist\excel-editor.exe --file "C:\Pfad\zur\Datei.xlsx" --sheet Sheet1 --rows 10
```

> **Hinweis:** EXE und Python-Paket sind funktional identisch. Für das Python-Paket gilt:  
> Das Projekt nutzt `pyproject.toml` (modernes Python-Packaging) – `pip install -e .` funktioniert unverändert.

---

## Verwendung

### 1. Zugriff prüfen / Sheet-Inhalt anzeigen

Testet ob die Datei erreichbar ist und gibt Spalten sowie die ersten N Zeilen aus.  
Hier passiert nichts an der Datei – nur lesen.

```bash
excel-editor \
  --file "/mnt/c/users/KuhneFa/BASF/BASF & SAP Collaboration - Dokumente/BASF_intern/RISE/100_Migration/999_TEST - System Linie/COP_TEST_Migration.xlsx" \
  --sheet Sheet1 \
  --rows 10
```

### 2. Zeile verschieben – Trockentest (kein Speichern)

Verschiebt Zeile `No=1010` direkt **nach** die Zeile mit `No=1026`.  
Das neue No wird automatisch als Mittelwert berechnet: `int((1026 + 1030) / 2) = 1028`.  
Am Ende wird interaktiv gefragt ob gespeichert werden soll → `n` eingeben zum Verwerfen.



### ⚠️ WICHTIG (Windows EXE)

Die standalone EXE ist ein **Windows‑Programm**.
Sie akzeptiert **keine Linux-/WSL‑Pfade** wie `/mnt/c/...`.

✅ Richtig:
C:\Users\Name\Pfad\Datei.xlsx

```bash
excel-editor \
  --file "/mnt/c/users/KuhneFa/BASF/BASF & SAP Collaboration - Dokumente/BASF_intern/RISE/100_Migration/999_TEST - System Linie/COP_TEST_Migration.xlsx" \
  --sheet Sheet1 \
  --move-from 1010 \
  --move-after 1026
# → Ausgabe: "Verschoben. Neues No=1028"
# → Frage: "Datei speichern? [J/n]:" → n eingeben
```

> **Wie das neue No berechnet wird:**  
> `--move-after 1026` → findet die nächste Zeile unter 1026 (z.B. No=1030)  
> → `new_no = int((1026 + 1030) / 2) = 1028`  
> Bricht ab wenn kein ganzzahliger Mittelwert möglich ist (z.B. 1026 und 1027 haben keinen Integer-Abstand) oder das berechnete No bereits existiert.

### 3. Zeile verschieben – direkt in Originaldatei speichern

Das Original wird **überschrieben**. Vorher eine Sicherungskopie anlegen!

```bash
excel-editor \
  --file "PFAD/ZUR/DATEI.xlsx" \
  --sheet Sheet1 \
  --move-from 1010 \
  --move-after 1025 \
  --save
```

### 4. Zeile verschieben – in neue Zieldatei speichern (empfohlen)

Das Original bleibt **unverändert**. Das Ergebnis wird in eine neue Datei geschrieben.  
Ideal zum Testen bevor man das Original anfasst.

```bash
excel-editor \
  --file "/mnt/c/users/KuhneFa/BASF/BASF & SAP Collaboration - Dokumente/BASF_intern/RISE/100_Migration/999_TEST - System Linie/COP_TEST_Migration.xlsx" \
  --sheet Sheet1 \
  --move-from 1010 \
  --move-after 1025 \
  --output "./test.xlsx"
```

---

## Alle Optionen

| Option | Kurz | Beschreibung |
|---|---|---|
| `--file` | `-f` | Pfad zur Excel-Datei (optional – wird sonst abgefragt) |
| `--sheet` | `-s` | Sheet-Name (optional – wird sonst abgefragt) |
| `--list-sheets` | | Zeigt alle Sheet-Namen und beendet |
| `--info` | | Zeigt Spalten und Metadaten des Sheets |
| `--rows N` | | Zeigt N Datenzeilen an (Standard: 5) |
| `--header-row N` | | Header-Zeile manuell angeben (Standard: automatisch erkannt) |
| `--move-from NO` | | `No`-Wert der zu verschiebenden Zeile |
| `--move-after NO` | | `No`-Wert der Zeile, **nach** der eingefügt wird. Neues No = Mittelwert zur darauffolgenden Zeile |
| `--save` | | Speichert Änderungen direkt ins Original |
| `--output` | `-o` | Speichert Ergebnis in angegebene Zieldatei |

---

## Interaktiver Modus

Ohne Argumente fragt das Tool Schritt für Schritt nach Pfad und Sheet:

```bash
excel-editor
```

---

## Als Python-Modul nutzen

Neben der CLI kann das Paket auch direkt in Python-Code importiert und aufgerufen werden.  
Das ist nützlich für:
- **Eigene Skripte**: z.B. ein Automatisierungsskript das mehrere Zeilen verschiebt
- **Agentic AI** (z.B. Copilot Agent, LangChain-Agent): Der Agent führt Python-Code aus und ruft die Funktionen direkt auf – **voraussetzung ist, dass `pip install -e .` in der Python-Umgebung des Agenten ausgeführt wurde**

```python
from pathlib import Path
from excel_editor import ExcelEditor, ExcelReadConfig

config = ExcelReadConfig(
    file_path=Path("PFAD/ZUR/DATEI.xlsx"),
    sheet_name="Sheet1",
    header_row=10,          # optional, wird sonst automatisch erkannt
)

with ExcelEditor(config) as editor:
    # Zeilen lesen
    rows = editor.get_rows()
    for row in rows[:5]:
        print(row.row_index, row.get_value(3))  # Spalte 3 = "No"

    # Zeile verschieben: No=1010 direkt nach No=1026
    # new_no = int((1026 + 1030) / 2) = 1028
    new_no = editor.move_row_after("1010", "1026")
    print(f"Neues No: {new_no}")  # → 1028

    # Speichern (Original überschreiben)
    editor.save()

    # Oder: in neue Datei speichern
    editor.save(output_path=Path("DATEI_KOPIE.xlsx"))
```

---

## Hinweise

- `No`-Werte und Excel-Zeilennummern sind **unabhängig voneinander**.  
  `--move-from 1010` bedeutet: die Zeile deren `No`-Spalte den Wert `1010` enthält.
- Die Formatierung (Hintergrundfarbe, Schrift, Rahmen, Zeilenhöhe) der verschobenen Zeile bleibt erhalten.
- Formeln in Zellen werden beim Verschieben erhalten (openpyxl liest und schreibt Formelstrings). Excel berechnet die Werte beim nächsten Öffnen neu.

### Speichern bei SharePoint / Netzwerkdateien

Dateien auf SharePoint oder Netzlaufwerken (`/mnt/c/...`) können **nicht direkt überschrieben** werden wenn:
- die Datei in Excel geöffnet ist (Windows sperrt sie)
- der Pfad schreibgeschützt ist

**Empfohlener Workflow:**

```bash
# 1. Ergebnis lokal speichern
excel-editor \
  --file "/mnt/c/users/KuhneFa/BASF/.../COP_TEST_Migration.xlsx" \
  --sheet Sheet1 \
  --move-from 1010 \
  --move-after 1026 \
  --output "/mnt/c/users/KuhneFa/Downloads/COP_TEST_ergebnis.xlsx"

# 2. Originaldatei in Excel schließen
# 3. Ergebnisdatei prüfen, dann manuell zurückkopieren
#    oder: Original überschreiben mit --save (nur wenn Excel geschlossen!)
```
