# SD-OfficeDeploy

Modernes Installationswerkzeug für Microsoft Office unter Windows.

<img width="1002" height="572" alt="SD-OfficeDeploy" src="https://github.com/user-attachments/assets/21c72f41-55e6-4970-80b4-c0ee75fae331" />

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

**Version:** 3.5  
**Maintainer:** SD-ITLab

---

## Beschreibung

**SD-OfficeDeploy** ist ein modernes, tastaturfreundliches Installationswerkzeug für Windows. Es ermöglicht die unbeaufsichtigte Installation verschiedener Microsoft-Office-Versionen über das Microsoft Office Deployment Tool (ODT).

Das Tool bietet:

- eine moderne Benutzeroberfläche mit CustomTkinter
- vordefinierte Profile für Retail- und LTSC-/Volumenlizenzprodukte
- vollständige Tastatursteuerung
- Sprach- und Architekturauswahl
- automatische Erstellung der ODT-Konfiguration
- automatische Bereinigung temporärer Installationsdateien

Geeignet für IT-Administratoren, Werkstätten, Rollouts und wiederkehrende Neuinstallationen.

---

## Unterstützte Office-Versionen

### Retail

- Microsoft 365 Single und Family
- Office 2024 Home, Home & Business und Professional Plus
- Office 2021 Home & Student, Home & Business und Professional Plus
- Office 2019 Home & Student, Home & Business und Professional Plus
- Office 2016 Home & Student, Home & Business und Professional

### LTSC / Volumenlizenz

- Office LTSC 2024 Standard und Professional Plus
- Office LTSC 2021 Standard und Professional Plus
- Office 2019 Standard und Professional Plus (Volume)

> **Hinweis:** LTSC- und Volume-Versionen benötigen eine passende Volumenlizenz. Die Aktivierung erfolgt abhängig von der vorhandenen Lizenzumgebung über KMS oder MAK.

---

## Funktionen

- Unbeaufsichtigte Office-Installation
- Vollständige Bedienung per Tastatur und Maus
- Getrennte Auswahl von Retail- und LTSC-Versionen
- Vordefinierte, versionsgenaue Produkt-IDs und Updatekanäle
- Sprachwahl sowie 32-Bit-, 64-Bit- und ARM64-Auswahl
- Automatische XML-Erstellung für das Office Deployment Tool
- Prüfung auf bereits vorhandene Office-Installationen
- OneDrive wird bei der Office-Erkennung ignoriert
- Verwaiste Registry-Einträge ohne installierte Office-Anwendung werden ignoriert
- Automatische Bereinigung temporärer Dateien

---

## Tastatursteuerung

| Taste | Funktion |
| --- | --- |
| Tab | Fokus wechseln |
| Pfeil hoch / runter | Office-Edition wechseln |
| Pfeil links / rechts | Office-Version oder Produktgruppe wechseln |
| Enter | Installation starten |
| Leertaste | Installation starten |

---

## Verwendung

1. Python-Anwendung beziehungsweise bereitgestellte Anwendung als Administrator starten.
2. Retail- oder LTSC-Version auswählen.
3. Gewünschte Edition, Sprache und Architektur festlegen.
4. Auf **Installieren** klicken oder die Installation per Tastatur starten.
5. Office wird automatisch über das Office Deployment Tool heruntergeladen und installiert.

---

## Technischer Hintergrund

- GUI: Python und CustomTkinter
- Backend: PowerShell
- Installation: Microsoft Office Deployment Tool
- Lizenz: MIT

### Zentrale Dateien

```text
SD-OfficeDeploy/
├── installergui.py
├── office_installer.ps1
└── readme.md
```

`installergui.py` und `office_installer.ps1` müssen sich für die direkte Ausführung im selben Verzeichnis befinden.

---

## Changelog

### Version 3.5

- Unterstützung für Office LTSC 2024 Standard und Professional Plus ergänzt
- Unterstützung für Office LTSC 2021 Standard und Professional Plus ergänzt
- Unterstützung für Office 2019 Standard und Professional Plus (Volume) ergänzt
- Navigation in Retail- und LTSC-Versionen unterteilt
- Produkt-IDs für Office 2019 Retail korrigiert
- Produkt-ID für Office 2021 Professional Plus Retail korrigiert
- Office 2016 Professional Plus korrekt in Office 2016 Professional umbenannt
- OneDrive aus der Prüfung auf vorhandene Office-Installationen ausgeschlossen
- Verwaiste Office-Registry-Einträge blockieren die Installation nicht mehr
- Öffentliches Branding vollständig auf SD-ITLab vereinheitlicht

### Version 3.0.0

- Erste öffentliche Veröffentlichung mit grafischer Benutzeroberfläche
- Auswahl verschiedener Office-Versionen und Editionen
- Vollständige Tastatursteuerung
- Automatisierte ODT-Konfiguration

### Version 2.x

- Interne PowerShell-Version
- Unbeaufsichtigte, XML-basierte Office-Installation

---

## Lizenz

Dieses Projekt steht unter der [MIT License](https://opensource.org/licenses/MIT).

---

## Entwickler

**SD-OfficeDeploy** wird von **SD-ITLab** entwickelt und gepflegt.

> Built with ❤️ for clean, fast and repeatable Office deployments.

---

## English

**SD-OfficeDeploy 3.5** is a keyboard-friendly Windows tool for unattended Microsoft Office deployments using the Office Deployment Tool.

It supports Microsoft 365, Office 2016–2024 Retail, Office LTSC 2021/2024, and Office 2019 Volume editions. LTSC and Volume installations require an appropriate KMS or MAK volume license.

### Key features

- Retail and LTSC/Volume installation profiles
- Unattended Office installation
- Full keyboard navigation
- Language and architecture selection
- Improved detection of existing Office installations
- OneDrive and stale registry entries are ignored during Office detection
- Automatic cleanup of temporary files

### License

MIT License

---

© 2026 SD-ITLab
