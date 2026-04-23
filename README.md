# Milestone Toolkit v4.2 — Fork by Vincent

Outil d'administration pour Milestone XProtect VMS, base sur le module PowerShell **MilestonePSTools**.

## Lancement

```powershell
.\Launch.ps1
```

Ou clic-droit sur `Launch.ps1` > **Executer avec PowerShell**.

---

## Fonctionnalites

### Snapshots

| Action | Description |
|--------|-------------|
| **Snapshot - Selection** | Capture un snapshot de la camera selectionnee via le dialogue Milestone |
| **Snapshot - Toutes les cameras** | Capture un snapshot de toutes les cameras en parallele (jusqu'a 12 simultanees) |
| **Snapshot - Presets PTZ** | Parcourt les presets PTZ et capture un snapshot a chaque position |

Toutes les actions snapshot supportent deux modes :
- **Live** : derniere image disponible
- **Historique** : image la plus proche d'une date/heure choisie

### Gestion

| Action | Description |
|--------|-------------|
| **Export Hardware** | Rapport Excel de tous les equipements (IP, MAC, firmware, identifiants). Option : snapshot integre par camera, recuperes en parallele |
| **Grouper par Modele** | Cree des groupes de cameras dans Milestone organises par modele |

### Monitoring

| Action | Description |
|--------|-------------|
| **Etat des cameras** | Etat temps reel de chaque camera (OK / hors ligne / erreur) via l'Event Server. Les cameras desactivees sont ignorees. Export CSV. |
| **Dates d'enregistrement** | Premier et dernier enregistrement disponible par camera, avec duree totale de retention. Export CSV. |

### Diagnostic

| Action | Description |
|--------|-------------|
| **Stats Enregistrement (7j)** | Statistiques d'enregistrement et de mouvement par camera sur 7 jours (CSV) |
| **Informations Licence** | Affiche les produits licencies, dates d'expiration et canaux utilises |

---

## Prerequis

- Windows 10/11 ou Windows Server 2016+
- PowerShell 5.1 (inclus dans Windows)
- Excel installe sur le poste (pour l'export Hardware avec snapshots)
- Acces reseau au serveur Milestone XProtect Management Server

Le module **MilestonePSTools** est installe automatiquement au premier lancement si Internet est disponible.

---

## Modes d'installation

### Online (par defaut)
Le module est telecharge automatiquement depuis PowerShell Gallery. Aucune action requise.

### Offline (machine sans Internet)
1. Sur une machine **avec** Internet, cliquer **Preparer offline** dans l'ecran de demarrage
2. Copier le projet entier (avec `Dependencies/`) sur la machine cible
3. Lancer normalement — le mode Offline est detecte automatiquement

---

## Configuration

Modifier `config.json` :

```json
{
    "outputDirectory": "./Output",
    "snapshotQuality": 95,
    "csvDelimiter": ";",
    "csvEncoding": "UTF8"
}
```

| Parametre | Description | Defaut |
|-----------|-------------|--------|
| `outputDirectory` | Dossier de sortie par defaut (snapshots, Excel, CSV) | `./Output` |
| `snapshotQuality` | Qualite JPEG des snapshots (1-100) | `95` |
| `csvDelimiter` | Separateur des fichiers CSV | `;` |
| `csvEncoding` | Encodage des fichiers CSV | `UTF8` |

Le dossier de sortie peut aussi etre change en cours d'utilisation via le bouton **Changer** dans la sidebar.

Le mode Online/Offline est detecte automatiquement selon la presence du dossier `Dependencies/`.

---

## Structure du projet

```
MilestonePSTools/
├── Launch.ps1                  # Point d'entree
├── config.json                 # Configuration
├── Dependencies/               # Modules offline (optionnel)
├── Logs/                       # Logs journaliers (auto)
├── Output/                     # Fichiers generes (auto)
└── src/
    ├── App.ps1                 # Chargement UI et evenements
    ├── UI/
    │   └── MainWindow.xaml     # Interface WPF (theme sombre Catppuccin)
    ├── Actions/
    │   ├── Get-SnapshotSelected.ps1
    │   ├── Get-SnapshotAll.ps1
    │   ├── Get-PtzPresetSnapshot.ps1
    │   ├── Export-HardwareReport.ps1
    │   ├── Set-CameraGroupByModel.ps1
    │   ├── Get-RecordingStats.ps1
    │   ├── Get-LicenseInfo.ps1
    │   ├── Get-CameraStatus.ps1
    │   └── Get-PlaybackReport.ps1
    └── Core/
        ├── Initialize-Modules.ps1
        ├── Show-StartupCheck.ps1
        ├── Write-ActivityLog.ps1
        └── Invoke-PtzPreset.ps1
```
