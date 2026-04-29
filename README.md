# Outil d’export XProtect vers Excel (GUI MilestonePSTools)

Outil avec interface graphique permettant d’exporter les caméras, le hardware, la rétention et les données d’enregistrement depuis Milestone XProtect vers Excel ou CSV, en s’appuyant sur le module PowerShell MilestonePSTools.

## Pourquoi cet outil

Milestone XProtect ne propose pas simplement :
- l’export complet des caméras vers Excel
- l’audit de la rétention d’enregistrement
- la génération de rapports exploitables

Cet outil permet de :
- gagner du temps sur les grosses installations
- générer des rapports clairs pour les clients
- auditer rapidement un système de vidéosurveillance

---

# Milestone Toolkit v4.6 — Fork by Vincent

Outil d'administration pour Milestone XProtect VMS, base sur le module PowerShell **MilestonePSTools**.
<img width="1244" height="864" alt="image" src="https://github.com/user-attachments/assets/909442eb-a8ec-4d50-b90f-e99df0070a9c" />

## Lancement

Cliquer sur Demarrer Milestone Toolkit.bat

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
| **Export Hardware** | Rapport Excel configurable : fenetre de selection des colonnes a inclure (hardware, flux video, retention, snapshot). Les mots de passe sont exclus par defaut et ne s'affichent que si la colonne est explicitement cochee. |
<img width="513" height="556" alt="image" src="https://github.com/user-attachments/assets/d8bb88fe-1b12-4ec0-b948-c56892808e2f" />

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

## Comment exporter toutes les caméras XProtect vers Excel

1. Lancer l’outil
2. Cliquer sur **Export Hardware**
3. Sélectionner les colonnes souhaitées
4. Exporter vers Excel

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
<img width="559" height="515" alt="image" src="https://github.com/user-attachments/assets/43cd7c30-020b-4fba-a45a-e524bd20feec" />

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

---

## Cas d’usage

- Audit de parc caméras
- Vérification de la rétention
- Export client
- Analyse de fonctionnement
- Maintenance système

---

## Mots-clés

Milestone XProtect export Excel  
export caméras XProtect  
outil audit vidéosurveillance  
MilestonePSTools GUI  
rapport caméras sécurité  
export hardware XProtect  
