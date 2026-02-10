# Milestone Toolkit

Outil d'administration pour Milestone XProtect VMS, base sur le module PowerShell **MilestonePSTools**.

## Fonctionnalites

| Action | Description |
|--------|-------------|
| **Snapshot - Selection** | Capture un snapshot d'une camera selectionnee via le dialogue Milestone |
| **Snapshot - Toutes** | Capture un snapshot de toutes les cameras du VMS |
| **Snapshot - Presets PTZ** | Parcourt les presets PTZ d'une camera et capture un snapshot a chaque position |
| **Export Hardware (CSV)** | Genere un rapport CSV de tous les equipements (IP, MAC, firmware, credentials, etc.) |
| **Grouper par Modele** | Cree des Device Groups dans Milestone organises par modele de camera |

## Prerequis

- Windows 10/11 ou Windows Server 2016+
- PowerShell 5.1+ (inclus dans Windows) ou PowerShell 7+
- Acces reseau au serveur Milestone XProtect Management Server

Les modules suivants sont installes automatiquement au premier lancement :
- [MilestonePSTools](https://www.powershellgallery.com/packages/MilestonePSTools)
- [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel)

## Lancement

```powershell
.\Launch.ps1
```

Ou clic-droit sur `Launch.ps1` > **Executer avec PowerShell**.

## Configuration

Modifier `config.json` pour personnaliser le comportement :

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
| `outputDirectory` | Repertoire de sortie (snapshots, CSV, etc.) | `./Output` |
| `snapshotQuality` | Qualite JPEG des snapshots (1-100) | `95` |
| `csvDelimiter` | Separateur pour l'export CSV | `;` |
| `csvEncoding` | Encodage du fichier CSV | `UTF8` |

## Structure du projet

```
MilestonePSTools/
├── Launch.ps1                  # Point d'entree
├── config.json                 # Configuration utilisateur
├── src/
│   ├── App.ps1                 # Bootstrap, chargement UI, wiring evenements
│   ├── UI/
│   │   └── MainWindow.xaml     # Interface WPF (theme sombre)
│   ├── Actions/                # 1 fichier = 1 action
│   │   ├── Get-SnapshotSelected.ps1
│   │   ├── Get-SnapshotAll.ps1
│   │   ├── Export-HardwareReport.ps1
│   │   ├── Set-CameraGroupByModel.ps1
│   │   └── Get-PtzPresetSnapshot.ps1
│   └── Core/                   # Utilitaires partages
│       ├── Initialize-Modules.ps1
│       ├── Write-ActivityLog.ps1
│       └── Invoke-PtzPreset.ps1
├── Logs/                       # Logs d'activite (generes automatiquement)
├── Output/                     # Fichiers de sortie (generes automatiquement)
└── Old/                        # Script original archive
```

## Ajouter une nouvelle action

1. Creer un fichier dans `src/Actions/` (ex: `New-MyAction.ps1`)
2. Implementer la fonction avec l'interface standard :

```powershell
function New-MyAction {
    param(
        [hashtable]$Config,
        [scriptblock]$Log
    )

    & $Log "Debut de l'action..."
    # Votre code ici
    & $Log "Terminee."
}
```

3. Dans `src/App.ps1`, ajouter le dot-source et le bouton correspondant

## Logs

Les logs sont ecrits dans le dossier `Logs/` avec un fichier par jour :
`MilestoneToolkit_2026-02-10.log`
