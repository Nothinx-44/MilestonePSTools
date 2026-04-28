$em = [char]0x2014   # em dash, safe pour PowerShell 5.1

$script:T = @{
    # LANGUAGE PICKER
    LP_Title    = 'Milestone Toolkit'
    LP_Subtitle = 'Choisir la langue / Select language'
    LP_FR       = 'Francais'
    LP_EN       = 'English'

    # STARTUP CHECK
    SC_WindowTitle    = "Milestone Toolkit v4.5 $em Demarrage"
    SC_Subtitle       = 'Verification des dependances au demarrage'
    SC_ModuleDesc     = "SDK Milestone VMS $em connexion au Management Server"
    SC_StatusWaiting  = 'En attente...'
    SC_BtnQuit        = 'Quitter'
    SC_BtnInstall     = 'Installer les dependances'
    SC_BtnLaunch      = "Lancer l'application"
    SC_BtnSaveDeps    = 'Preparer offline'
    SC_BtnUpdateCache = 'Mettre a jour le cache'
    SC_StatusInit     = 'Verification en cours...'
    SC_Checking       = 'Verification...'
    SC_LocalCache     = 'Disponible (cache local)'
    SC_Installed      = 'Installe  v{0}'
    SC_Missing        = 'Non installe'
    SC_Installing     = 'Installation...'
    SC_Saving         = 'Sauvegarde en cours...'
    SC_CacheOk        = 'Cache local cree'
    SC_OfflineCacheTitle   = 'Cache local detecte'
    SC_OfflineCacheMissing = "Module(s) absent(s) du cache local : {0}. Connectez-vous a Internet et cliquez 'Preparer offline'."
    SC_OfflineCacheOk      = 'Tous les modules sont en cache local. Le projet peut etre utilise sans Internet.'
    SC_OnlineTitle    = 'Connexion Internet disponible'
    SC_OnlineText     = "Cliquez 'Preparer offline' pour sauvegarder les modules et utiliser le projet sans connexion."
    SC_AllOk          = 'Toutes les dependances sont disponibles. Pret a lancer.'
    SC_OfflineMissing = "Impossible d'installer sans Internet. Voir le message ci-dessus."
    SC_NeedInstall    = 'Des dependances sont manquantes. Cliquez sur "Installer les dependances".'
    SC_NuGet          = 'Preparation du fournisseur NuGet...'
    SC_InstallingMod  = 'Installation de {0} depuis PowerShell Gallery...'
    SC_DownloadingMod = 'Telechargement de {0} pour usage offline...'
    SC_InstallDone    = 'Installation terminee. Pret a lancer.'
    SC_InstallError   = 'Certaines installations ont echoue. Verifiez votre connexion Internet.'
    SC_SaveError      = 'La sauvegarde a echoue. Verifiez votre connexion Internet.'
    SC_SaveConfirm    = "Cette operation va telecharger les modules dans le dossier Dependencies/.`n`nVous pourrez ensuite copier tout le projet sur une machine sans Internet.`n`nContinuer ?"
    SC_SaveTitle      = 'Preparer pour usage offline'
    SC_ErrTitle       = "Milestone Toolkit $em Erreur"
    SC_ErrCheck       = "Erreur lors de la verification :`n`n{0}"
    SC_ErrInstall     = "Erreur lors de l'installation :`n`n{0}"
    SC_ErrGeneric     = 'Erreur : {0}'

    # MAIN WINDOW
    MW_StatusConnected   = 'Connecte'
    MW_LblOutputDir      = 'DOSSIER DE SORTIE'
    MW_BtnOutputDir      = 'Changer'
    MW_AppTitle          = 'Milestone Toolkit v4.5'
    MW_Version           = "v4.5 $em Fork by Vincent"
    MW_LblModeCapture    = 'MODE CAPTURE'
    MW_CbiLive           = 'Live (derniere image)'
    MW_CbiHistorique     = 'Image historique'
    MW_LblDate           = 'Date'
    MW_LblHeure          = 'Heure (HH : mm)'
    MW_LblSnapshots      = 'SNAPSHOTS'
    MW_BtnSnapshotSel    = 'Snapshot - Selection'
    MW_BtnSnapshotAll    = 'Snapshot - Toutes les cameras'
    MW_BtnPtz            = 'Snapshot - Presets PTZ'
    MW_LblGestion        = 'GESTION'
    MW_BtnExportHardware = 'Export Hardware (Excel)'
    MW_BtnGroupByModel   = 'Grouper par Modele'
    MW_LblMonitoring     = 'MONITORING'
    MW_BtnCameraStatus   = 'Etat des cameras'
    MW_BtnPlaybackReport = "Dates d'enregistrement"
    MW_LblDiagnostic     = 'DIAGNOSTIC'
    MW_BtnRecordingStats = 'Stats Enregistrement (7j)'
    MW_BtnLicenseInfo    = 'Informations Licence'
    MW_BtnClearLog       = 'Effacer'
    MW_BtnCancel         = 'Annuler'
    MW_LblJournal        = "Journal d'activite"
    MW_StatusReady       = 'Pret'
    MW_StatusCancelling  = 'Annulation en cours...'

    # APP MESSAGES
    App_DateMissing     = 'Veuillez selectionner une date.'
    App_DateTitle       = 'Date manquante'
    App_HourInvalid     = 'Heure invalide. Entrez une valeur entre 0 et 23.'
    App_HourTitle       = 'Heure invalide'
    App_MinInvalid      = 'Minutes invalides. Entrez une valeur entre 0 et 59.'
    App_MinTitle        = 'Minutes invalides'
    App_ChooseDir       = 'Choisir le dossier de sortie'
    App_OutputChanged   = 'Dossier de sortie change : {0}'
    App_Started         = 'Application demarree. Connecte au serveur Milestone.'
    App_OutputDir       = 'Repertoire de sortie : {0}'
    App_ActionDone      = 'Action terminee avec succes.'
    App_ActionCancelled = 'Operation annulee.'
    App_Closing         = 'Fermeture de l application'

    # ACTION NAMES
    Act_SnapshotSel  = 'Snapshot - Selection'
    Act_SnapshotAll  = 'Snapshot - Toutes les cameras'
    Act_SnapshotPtz  = 'Snapshot - Presets PTZ'
    Act_ExportHW     = 'Export Hardware'
    Act_GroupModel   = 'Grouper par Modele'
    Act_CamStatus    = 'Etat des cameras'
    Act_Playback     = "Dates d enregistrement"
    Act_RecStats     = 'Stats Enregistrement (7 jours)'
    Act_License      = 'Informations Licence'

    # EXPORT HARDWARE DIALOG
    EH_DialogTitle   = "Options d'export"
    EH_SelectCols    = "Selectionnez les colonnes a inclure dans l'export :"
    EH_BtnSelectAll  = 'Tout cocher'
    EH_BtnDeselAll   = 'Tout decocher'
    EH_GrpHardware   = 'Informations hardware'
    EH_GrpFlux       = 'Flux video'
    EH_GrpRetention  = 'Retention'
    EH_GrpOptions    = 'Options'
    EH_ChkNom        = 'Nom'
    EH_ChkFabricant  = 'Fabricant'
    EH_ChkModele     = 'Modele'
    EH_ChkIP         = 'IP'
    EH_ChkMAC        = 'MAC'
    EH_ChkFirmware   = 'Firmware'
    EH_ChkServeurRec = 'Serveur Enreg.'
    EH_ChkUser       = 'Utilisateur'
    EH_ChkPassword   = 'Mot de passe (!)'
    EH_ChkCodecRec   = 'Codec (Enreg.)'
    EH_ChkResRec     = 'Resolution (Enreg.)'
    EH_ChkFpsRec     = 'FPS (Enreg.)'
    EH_ChkCodecLive  = 'Codec (Live)'
    EH_ChkResLive    = 'Resolution (Live)'
    EH_ChkFpsLive    = 'FPS (Live)'
    EH_ChkFluxSupp   = 'Flux supplementaires'
    EH_ChkRetention  = 'Retention disponible'
    EH_ChkSnapshot   = 'Snapshot (lent - 1 image par camera)'
    EH_BtnCancel     = 'Annuler'
    EH_BtnExport     = 'Exporter'
    EH_NoColumn      = 'Veuillez selectionner au moins une colonne.'
    EH_NoColumnTitle = 'Aucune colonne selectionnee'
    EH_Cancelled     = "Export annule par l'utilisateur."

    # EXCEL
    XL_FileName     = 'Liste_des_Cameras.xlsx'
    XL_SheetName    = 'Cameras'
    XL_Nom          = 'Nom'
    XL_Fabricant    = 'Fabricant'
    XL_Modele       = 'Modele'
    XL_IP           = 'IP'
    XL_MAC          = 'MAC'
    XL_Firmware     = 'Firmware'
    XL_ServeurRec   = 'ServeurRec'
    XL_Utilisateur  = 'Utilisateur'
    XL_MotDePasse   = 'MotDePasse'
    XL_CodecEnreg   = 'Codec (Enreg.)'
    XL_ResEnreg     = 'Resolution (Enreg.)'
    XL_FpsEnreg     = 'FPS (Enreg.)'
    XL_CodecLive    = 'Codec (Live)'
    XL_ResLive      = 'Resolution (Live)'
    XL_FpsLive      = 'FPS (Live)'
    XL_FluxSupp     = 'Flux supplementaires'
    XL_Retention    = 'Retention disponible'
    XL_Snapshot     = 'Snapshot'
    XL_Aucun        = 'Aucun'
    XL_ExtraFlux    = '{0} flux supp.'

    # EXPORT HARDWARE LOGS
    EH_LogGenerating  = 'Generation du rapport hardware...'
    EH_LogFound       = '{0} equipements trouves.'
    EH_LogLoadCams    = 'Chargement des objets camera...'
    EH_LogStreams      = 'Recuperation des configurations de flux video...'
    EH_LogStreamsOk   = '{0} flux trouves pour {1} cameras.'
    EH_LogStreamsErr   = 'AVERTISSEMENT: Impossible de recuperer les flux video : {0}'
    EH_LogPlayback    = "Recuperation des dates d'enregistrement..."
    EH_LogPlaybackOk  = 'Dates recuperees pour {0} cameras.'
    EH_LogPlaybackErr = "AVERTISSEMENT: Impossible de recuperer les dates d'enregistrement : {0}"
    EH_LogSnaps       = 'Recuperation des snapshots en parallele...'
    EH_LogSnapOk      = '  [OK {0}/{1}] {2}'
    EH_LogSnapEmpty   = "  AVERTISSEMENT: Snapshot vide '{0}'"
    EH_LogSnapErr     = "  AVERTISSEMENT: '{0}' : {1}"
    EH_LogSnapsDone   = '{0} / {1} snapshots recuperes.'
    EH_LogBuilding    = 'Construction du fichier Excel...'
    EH_LogCancelled   = 'AVERTISSEMENT: Operation annulee apres {0} / {1} cameras.'
    EH_LogCamRow      = '[{0}/{1}] {2}'
    EH_LogImgErr      = "  AVERTISSEMENT: Image '{0}' : {1}"
    EH_LogSaved       = 'Rapport exporte : {0}'
    EH_LogNoExcel     = "ERREUR: Excel n'est pas installe sur ce poste."

    # CAMERA STATUS
    CS_LogQuerying  = "Interrogation de l'Event Server..."
    CS_LogError     = 'ERREUR: Get-ItemState : {0}'
    CS_LogFound     = '{0} cameras interrogees.'
    CS_LogCancelled = 'AVERTISSEMENT: Operation annulee apres {0} / {1}.'
    CS_LogOk        = '  [OK] {0}'
    CS_LogErr       = '  ERREUR: {0} : {1}'
    CS_LogDisabled  = '  ({0} camera(s) desactivee(s) ignoree(s))'
    CS_LogKo        = 'AVERTISSEMENT: {0} camera(s) hors ligne ou en erreur sur {1} actives.'
    CS_LogAllOk     = 'Toutes les cameras actives ({0}) sont operationnelles.'
    CS_LogExported  = 'Rapport exporte : {0}'
    CS_CsvNom       = 'Nom'
    CS_CsvEtat      = 'Etat'
    CS_CsvType      = 'Type'
    CS_CsvOk        = 'OK'
    CS_CsvFileName  = 'Etat_Cameras.csv'

    # PLAYBACK REPORT
    PR_LogCams      = 'Recuperation des cameras...'
    PR_LogFound     = '{0} cameras trouvees. Recuperation des plages d enregistrement...'
    PR_LogError     = 'ERREUR: Get-PlaybackInfo : {0}'
    PR_LogCancelled = 'AVERTISSEMENT: Operation annulee apres {0} / {1}.'
    PR_LogRow       = '  {0} : {1} -> {2} ({3})'
    PR_LogNoRec     = '  AVERTISSEMENT: {0} : aucun enregistrement trouve'
    PR_LogExported  = 'Rapport exporte : {0}'
    PR_CsvNom       = 'Nom'
    PR_CsvPremier   = 'PremierEnreg'
    PR_CsvDernier   = 'DernierEnreg'
    PR_CsvDuree     = 'DureeDisponible'
    PR_CsvFileName  = 'Dates_Enregistrement.csv'
    PR_DurDays      = '{0}j {1}h'
    PR_DurHours     = '{0}h {1}m'

    # RECORDING STATS
    RS_LogPeriod    = 'Periode analysee : {0} -> {1}'
    RS_LogCams      = 'Recuperation des cameras...'
    RS_LogFound     = '{0} cameras trouvees.'
    RS_LogCancelled = 'AVERTISSEMENT: Operation annulee apres {0} / {1} cameras.'
    RS_LogProgress  = '[{0}/{1}] {2}'
    RS_LogRec       = '  Enreg. : {0} sequences | {1} | {2}'
    RS_LogRecWarn   = '  AVERTISSEMENT: Stats enregistrement : {0}'
    RS_LogMot       = '  Motion : {0} sequences | {1}'
    RS_LogMotWarn   = '  AVERTISSEMENT: Stats mouvement : {0}'
    RS_LogLive      = '  Live   : {0} FPS | {1} kbps | {2}'
    RS_LogExported  = 'Rapport exporte : {0}'
    RS_CsvNom       = 'Nom'
    RS_CsvActive    = 'Active'
    RS_CsvSeqRec    = 'Sequences_Rec'
    RS_CsvPctRec    = 'PctTemps_Rec'
    RS_CsvTimeRec   = 'TempsEnregistre'
    RS_CsvSeqMot    = 'Sequences_Motion'
    RS_CsvPctMot    = 'PctTemps_Motion'
    RS_CsvFps       = 'FPS_Live'
    RS_CsvBitrate   = 'Bitrate_Live_kbs'
    RS_CsvRes       = 'Resolution'
    RS_CsvFileName  = 'Stats_Enregistrement.csv'
    RS_DurDays      = '{0}j {1}h {2}m'
    RS_DurHours     = '{0}h {1}m'

    # SNAPSHOT ALL
    SA_LogCams      = 'Recuperation de la liste des cameras...'
    SA_LogFound     = '{0} cameras trouvees. Capture en parallele...'
    SA_LogHistorique = 'Mode historique : {0}'
    SA_LogCancelled = 'AVERTISSEMENT: Operation annulee avant lancement.'
    SA_LogOk        = '  [OK {0}/{1}] {2}'
    SA_LogFailed    = "  AVERTISSEMENT: Echec snapshot '{0}'"
    SA_LogError     = "  ERREUR '{0}': {1}"
    SA_LogDone      = '{0} snapshots enregistres dans : {1}'
    SA_LogDoneErr   = '{0} snapshots enregistres, {1} echecs dans : {2}'

    # SNAPSHOT SELECTED
    SS_LogOpening   = 'Ouverture du selecteur de camera...'
    SS_LogNone      = 'Aucune camera selectionnee. Operation annulee.'
    SS_LogHistorique = 'Mode historique : {0}'
    SS_LogCapturing = "Capture du snapshot de '{0}'..."
    SS_LogSaved     = 'Snapshot enregistre dans : {0}'
    SS_LogError     = "ERREUR: Echec du snapshot de '{0}' : {1}"

    # PTZ SNAPSHOT
    PTZ_LogSelecting = 'Selection des cameras PTZ...'
    PTZ_LogNone      = 'Aucune camera PTZ avec presets selectionnee.'
    PTZ_LogFound     = '{0} camera(s) PTZ avec presets trouvee(s).'
    PTZ_LogHistorique = 'Mode historique : {0}'
    PTZ_LogCamera    = "Camera '{0}' : {1} preset(s)."
    PTZ_LogMoving    = "  Deplacement vers preset '{0}'..."
    PTZ_LogPosErr    = '  AVERTISSEMENT: Verification de position echouee: {0}'
    PTZ_LogCapturing = '  Capture du snapshot...'
    PTZ_LogSaved     = "  Snapshot '{0}' enregistre."
    PTZ_LogError     = "ERREUR: Snapshot '{0}' echoue : {1}"
    PTZ_LogDone      = 'Capture PTZ terminee. Fichiers dans : {0}'
    PTZ_LogCancelled = 'AVERTISSEMENT: Operation annulee.'

    # GROUP BY MODEL
    GM_LogRetrieving   = 'Recuperation des informations cameras...'
    GM_LogFound        = '{0} cameras trouvees.'
    GM_LogParentCreated = "Dossier parent '{0}' cree."
    GM_LogModels       = '{0} modeles differents detectes.'
    GM_LogCancelled    = 'AVERTISSEMENT: Operation annulee apres {0} / {1} modeles.'
    GM_LogModel        = "Modele '{0}' : {1} camera(s) ajoutee(s)."
    GM_LogDone         = 'Organisation par modele terminee.'
    GM_Unknown         = 'Inconnu'
    GM_ParentFolder    = 'Modele'

    # LICENSE INFO
    LI_LogHeader     = '--- Produits licencies ---'
    LI_LogNone       = '  Aucun produit licence trouve.'
    LI_LogProduct    = 'Produit : {0}'
    LI_LogExpired    = '  ERREUR: Licence expiree le {0} ({1} jours depasses)'
    LI_LogExpSoon    = '  AVERTISSEMENT: Expiration le {0} (dans {1} jours)'
    LI_LogExpiry     = '  Expiration      : {0} (dans {1} jours)'
    LI_LogExpiryRaw  = '  Expiration      : {0}'
    LI_LogPerpetual  = '  Expiration      : Aucune (licence perpetuelle)'
    LI_LogSlc        = '  SLC             : {0}'
    LI_LogChanWarn   = '  AVERTISSEMENT: Canaux : {0} / {1} ({2} % utilises)'
    LI_LogChan       = '  Canaux          : {0} / {1} ({2} % utilises)'
    LI_LogChanSingle = '  Canaux          : {0} utilises / {1}'
    LI_LogChanRaw    = '  Canaux          : {0}'
    LI_LogError      = 'ERREUR: Get-VmsLicensedProducts : {0}'
}
