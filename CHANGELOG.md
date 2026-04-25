# Release Notes

## v4.3
> Enrichissement de l'export Hardware

### Ameliorations
- **Export Hardware — Flux video** : ajout de 7 nouvelles colonnes groupees par couleur dans le fichier Excel : `Codec (Enreg.)`, `Resolution (Enreg.)`, `FPS (Enreg.)`, `Codec (Live)`, `Resolution (Live)`, `FPS (Live)`, `Flux supplementaires`. Si le flux live et le flux enregistre sont identiques, les colonnes Live restent vides pour eviter la redondance.
- **Export Hardware — Retention disponible** : nouvelle colonne `Retention disponible` indiquant la duree totale d'archive accessible par camera, calculee via `Get-PlaybackInfo`.
- **Export Hardware — Adresse IP** : suppression automatique du protocole (`http://`) et du port (`:8000`) — seule l'adresse IP pure (4 blocs) est affichee.
- **En-tetes Excel colores par groupe** : hardware (sable), flux video (bleu marine), retention (vert), snapshot (violet) pour une lecture rapide.

---

## v4.2
> Nouvelle categorie Monitoring

### Nouveautes
- **Etat des cameras** (`Get-ItemState`) : etat temps reel de chaque camera via l'Event Server. Les cameras Responding sont affichees en vert, les erreurs en rouge, les cameras desactivees sont automatiquement ignorees. Export CSV `Etat_Cameras.csv`.
- **Dates d'enregistrement** (`Get-PlaybackInfo`) : premier et dernier enregistrement disponible par camera avec la duree de retention calculee. Utilise le mode parallele natif du SDK pour les grands systemes. Export CSV `Dates_Enregistrement.csv`.
- **Nouvelle categorie MONITORING** dans la sidebar (couleur cyan) separee de DIAGNOSTIC pour distinguer les donnees temps reel des analyses historiques.

---

## v4.1
> Optimisation des performances

### Ameliorations
- **Snapshot - Toutes les cameras** : les snapshots sont desormais recuperes en parallele via un pool de threads (jusqu'a 12 simultanees). Le temps de capture est divise par le nombre de cameras actives simultanement au lieu d'etre sequentiel.
- **Export Hardware - Snapshots** : meme amelioration — les snapshots sont pre-telecharges en parallele avant la construction du fichier Excel. Le log affiche chaque snapshot des qu'il est recu, dans l'ordre d'arrivee.
- **Configuration** : suppression du parametre `installMode` dans `config.json` — le mode Online/Offline est toujours auto-detecte (presence du dossier `Dependencies/`).

---

## v4.0
> Refonte de l'export hardware + nouvelles fonctions de diagnostic + mode capture historique

### Nouveautes

#### Export Hardware
- Format de sortie change de **CSV vers Excel (.xlsx)** avec mise en forme (en-tete colore, bordures, colonne figee)
- Option d'inclure un **snapshot integre dans la cellule** de chaque camera (ancre, redimensionne automatiquement)
- Suppression des colonnes GPS et Activation (inutiles)
- Demande de confirmation avant l'inclusion des snapshots (operation longue)

#### Mode capture historique
- Nouveau selecteur **Live / Image historique** dans la sidebar
- En mode historique : choix de la date et de l'heure (hh:mm)
- Applicable aux trois actions snapshot : Selection, Toutes les cameras, Presets PTZ
- Utilise `Get-Snapshot -Behavior GetNearest` pour trouver l'image la plus proche

#### Diagnostic (nouvelle categorie)
- **Stats Enregistrement (7j)** : sequences, pourcentage de temps enregistre, duree totale, detection de mouvement et statistiques live (FPS, bitrate, resolution) par camera. Export CSV.
- **Informations Licence** : produits licencies, date d'expiration avec alerte si < 30 jours, canaux utilises/total avec alerte si >= 90 %. Utilise `Get-VmsLicensedProducts`.

### Corrections
- Rapport Stockage supprime (les donnees d'espace disque ne sont pas accessibles sans droits WMI sur le serveur distant)
- `Get-VmsLicensedProducts` : gestion correcte des valeurs d'expiration non-date (`"Unrestricted"`)
- Proprietes internes du SDK Milestone filtrees de l'affichage licence (Path, ParentPath, ServerId, etc.)

---

## v3.0
> Fork initial — refonte UX complete

### Nouveautes
- **Journal d'activite colore** : remplacement du TextBox monochrome par un RichTextBox avec codes couleur par niveau (erreur, avertissement, succes, action)
- **Bouton Annuler** : toutes les actions longues peuvent etre interrompues entre chaque iteration
- **Barre de progression** : deterministe (avancement reel) pour les actions iteratives
- **Selecteur de dossier de sortie** : changement du repertoire sans modifier `config.json`
- **Snapshot - Presets PTZ** : progression basee sur le nombre total de presets toutes cameras confondues
- **Grouper par Modele** : verification des doublons avant ajout au groupe

### Corrections
- Encodage du titre de fenetre corrige (`—` affiché correctement via entite XML `&#x2014;`)
- Banniere de demarrage : distinction correcte entre mode hors ligne (cache local) et connexion Internet disponible
- ComboBox mode capture : template sombre complet pour etre lisible sur fond sombre
- `Get-CameraRecordingStats` : parametres corriges (`-StartTime`/`-EndTime`, proprietes `PercentRecorded`/`TimeRecorded`)
- `Get-VmsStorageRetention` : retourne un `[TimeSpan]` directement, `.TotalDays` utilise
- Confirmation avant inclusion des mots de passe en clair dans l'export hardware
