# SD-OfficeDeploy
Modern Office installer for Windows

<img width="1002" height="572" alt="image" src="https://github.com/user-attachments/assets/52d10561-21ee-4ffe-a8f4-e71a6fb2f710" />

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

# SD-OfficeDeploy

**Version:** 3.0.0
**Maintainer:** SD-ITLab

---

## ✨ Beschreibung

**SD-OfficeDeploy** ist ein modernes, **keyboard-freundliches Office-Installations-Tool für Windows**, das die **unbeaufsichtigte Installation** verschiedener Microsoft-Office-Versionen ermöglicht.

Das Tool kombiniert:

* eine **moderne GUI (CustomTkinter)**
* **vordefinierte Installationsprofile**
* volle **Tastatursteuerung (Tab / Pfeiltasten / Enter)**
* sowie eine saubere **Automatisierung über Office Deployment Tool (ODT)**

Ideal für **IT-Administratoren**, **Werkstätten**, **Rollouts** und **wiederkehrende Neuinstallationen**.

---

## 🔧 Unterstützte Office-Versionen

* 🟦 **Office 2016**
* 🟦 **Office 2019**
* 🟦 **Office 2021**
* 🟦 **Office 2024**
* 🟩 **Microsoft 365** (Single / Family)

### Editionen (abhängig vom Jahr)

* Home & Student
* Home & Business
* Professional / Professional Plus

---

## ⚙ Funktionen

* ✅ Unbeaufsichtigte Office-Installation
* 🎹 Vollständige **Tastaturbedienung**
* 🧩 Vordefinierte Editionen & Konfigurationen
* 🌍 Sprach- & Architektur-Auswahl (32 / 64 / ARM64)
* 🧠 Automatische XML-Erstellung für ODT
* 🛡️ Keine Freitexteingaben – nur valide Auswahlmöglichkeiten
* 🧼 Temporäre Dateien werden automatisch bereinigt

---

## ⌨ Tastatursteuerung

| Taste     | Funktion                             |
| --------- | ------------------------------------ |
| Tab       | Fokus wechseln                       |
| ↑ / ↓     | Office-Edition wechseln              |
| ← / →     | Office-Jahr / Produktgruppe wechseln |
| Enter     | Installation starten                 |
| Leertaste | Installation starten                 |

---

## 🚀 Verwendung

1. **SD-OfficeDeploy.exe** als Administrator starten
2. Gewünschte **Office-Version & Edition** auswählen
3. Sprache, Architektur und Kanal festlegen
4. **Enter drücken** oder auf *Installieren* klicken
5. Office wird **vollautomatisch im Hintergrund** installiert

---

## 🛠 Technischer Hintergrund

* Backend: **PowerShell + Office Deployment Tool (ODT)**
* GUI: **Python (CustomTkinter)**
* Paketierung: **PyInstaller**
* Lizenz: **MIT**

---

## 📁 Verzeichnisstruktur (vereinfacht)

```text
SD-OfficeDeploy/
├─ assets/
│  └─ logo.png
├─ scripts/
│  └─ install_office.ps1
├─ config/
│  └─ office_config.xml
└─ SD-OfficeDeploy.exe
```

---

## 📝 Changelog

* **3.0.0** – Erste öffentliche Veröffentlichung mit GUI

  * Neue **grafische Oberfläche (GUI)**
  * Erweiterte Auswahlmöglichkeiten für Office-Versionen & Editionen
  * Vollständige Tastatursteuerung (Tab / Pfeiltasten / Enter)
  * Basierend auf vorheriger interner PowerShell-Version

* **2.x** – Interne PowerShell-Version (unveröffentlicht)

  * Unbeaufsichtigte Office-Installation
  * XML-basierte ODT-Konfiguration

---

## 📜 Lizenz

Dieses Projekt steht unter der **MIT License**.

---

## 👨‍💻 Entwickler

**SD-OfficeDeploy** wird entwickelt und gepflegt von **SD-ITLab**.

> Built with ❤️ for clean, fast and repeatable Office deployments.

---

# English

## SD-OfficeDeploy

**SD-OfficeDeploy** is a modern, keyboard-driven Office installation tool for Windows, designed for unattended and repeatable deployments using the Microsoft Office Deployment Tool.

### Key Features

* Unattended Office installation
* Full keyboard navigation
* Predefined Office editions and configurations
* Language & architecture selection
* Clean temporary file handling

### License

MIT License

---

© 2026 SD-ITLab
