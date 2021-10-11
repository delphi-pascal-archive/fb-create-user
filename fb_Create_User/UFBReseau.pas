unit UFBReseau;

// |===========================================================================|
// | unit UFBReseau                                                            |
// | Copyright © 2006 F.BASSO                                                  |
// |___________________________________________________________________________|
// | Unité d'interface du Lan Manager                                          |
// |                                                                           |
// | Portions created by Microsoft are                                         |
// | Copyright (C) 1995-1999 Microsoft Corporation                             | 
// | All Rights Reserved.                                                      |
// |___________________________________________________________________________|
// | Ce programme est libre, vous pouvez le redistribuer et ou le modifier     |
// | selon les termes de la Licence Publique Générale GNU publiée par la       |
// | Free Software Foundation .                                                |
// | Ce programme est distribué car potentiellement utile,                     |
// | mais SANS AUCUNE GARANTIE, ni explicite ni implicite,                     |
// | y compris les garanties de commercialisation ou d'adaptation              |
// | dans un but spécifique.                                                   |
// | Reportez-vous à la Licence Publique Générale GNU pour plus de détails.    |
// |                                                                           |
// | anbasso@wanadoo.fr                                                        |
// |___________________________________________________________________________|
// | Versions                                                                  |
// | ______                                                                    |
// || 4.0  \__________________________________________________________________ |
// || Ajout function CreateUser                                               ||
// ||______                                                                   ||
// || 3.0  \__________________________________________________________________||
// || Ajout function ListePartage                                             ||
// ||______                                                                   ||
// || 2.0  \__________________________________________________________________||
// || Ajout function ArretPartage                                             ||
// ||______                                                                   ||
// || 1.0  \__________________________________________________________________||
// || Création de l'unité avce les fonctions suivante                         ||
// || Function  GetinfoUtilisateur                                            ||
// || Function  GetinfoUtilisateur                                            ||
// || Function  GetListeComptes                                               ||
// || Function  GetListeGroupeGlobaux                                         ||
// || Function  GetListeGroupeLocaux                                          ||
// || Function  GetListeUtilisateursGroupeGlobal                              ||
// || Function  GetListeUtilisateursGroupeGlobal                              ||
// || Function  GetNomDomaine                                                 ||
// || function  GetNomTousPdc                                                 ||
// || function  GetNomPdc                                                     ||
// || function  GetListeServeur                                               ||
// || function  GetListeDisque                                                ||
// || function  GetListeGroupeUtilisateur                                     || 
// ||                                                                         ||
// || function  AjouteUtilisateurGroupe                                       ||
// || function  SupprimeUtilisateurGroupe                                     ||
// ||_________________________________________________________________________||
// |===========================================================================|

interface
uses comctrls,accctrl;
Type

  TDisque = array[0..2] of char;
  TTabDisque = array of Tdisque;

  Tacl = record
    accesmode : ACCESS_MODE;
    accesrigth : cardinal;
    strDomain : string;
    strNom : string;
  end;
  TTabAcl = array of Tacl;

  Tpartage = record
    strNom : string;
    strChemin : string;
    intType : cardinal;
    tabacl : ttabacl;
  end;
  TTabPartage = array of Tpartage;

// =============================================================================
//  Type Tutilisateur0
// =============================================================================

  Tutilisateur0 = record
    login                   : string;           // Nom de l'utilisateur
  end;

  TTabUtilisateur0 = array of Tutilisateur0 ;

// =============================================================================
//  Type Tutilisateur
// =============================================================================

  Tutilisateur = record
    login                   : string;           // Nom de l'utilisateur
    mot_de_passe            : string;           // Mot de passe vide lors de la lecture
    age_mot_de_passe        : longword;         // Age du mot de passe (en seconde)
    privilege               : longword ;        // indicateur des priviléges de l'utilisateur  (Invité,utilisateur,administrateur)
    dossier_de_base         : string;           // Chemin du répertoire de l'utilsateur
    description             : string;           // Description de l'utilisateur
    flags                   : longword;         // indicateur d'état du compte
    dossier_script          : string;           // nom du script d'ouverture de session
    flags_privilege         : longword;         // indicateur des priviléges d'execution
    nom_complet             : string;           // nom complet de l'utilisateur
    commentaire_utilisateur : string;           // Commentaire
    parametre               : string;           // reservé
    ordinateurs             : string;           // Liste des ordinateurs où l'utilisateur à le droit de se loguer
    date_connexion          : TDateTime ;       // Date de la derniere connexion
    date_deconnexion        : TDateTime ;       // date de la derniere déconnexion
    date_expiration         : TDateTime ;       // date d'expiration du compte
    stockage_maxi           : longword;         // Taille maxi sur le repertoire de l'utilisateur
    units_per_week          : longword;         //
    logon_hours             : pbyte ;           //
    nb_tentative_login      : longword;         // Nombre de tentative de connexion
    nb_login                : longword;         // Nombre de connexion
    serveur_de_login        : string;           // Serveur ayant reçue la demande de login
    code_region             : longword;         // code régional
    code_page               : longword;         // code page
    id_utilisateur          : longword;         // Relative IDentifier de l'utilisateur
    id_groupe_principal     : longword;         // Relative IDentifier du groupe global principale de l'utilisateur
    profile                 : String;           // Chemin du profile utilisateur
    lettre_dossier_de_base  : string;           // Lecteur où est mappé le répertoire utilisateur
    mot_passe_expire        : boolean;          // Le mot de passe à expiré
  end;

  TTabUtilisateur = array of Tutilisateur ;

// =============================================================================
//  Type TGroupeGlobaux
// =============================================================================

  TGroupeGlobaux = record
    nom         : string;                       // Nom du groupe
    commentaire : string;                       // Commentaire
    id_group    : longword;                     // Relative IDentifier du groupe
    attributs   : longword;                     // Attrribut du groupe
  end;

  TTabGroupeGlobaux = array of TGroupeGlobaux ;

// =============================================================================
//  Type TGroupeLocaux
// =============================================================================

  TGroupeLocaux = record
    nom : string ;                             // Nom du groupe
    commentaire : string ;                     // Commentaire
  end;

  TTabgroupeLocaux = array of TGroupeLocaux ;

// =============================================================================
//  Type TServeur
// =============================================================================

  TServeur = record
     platformId   : LongWord ;
     nom          : string ;
     versionMajor : longword;
     versionMinor : longword;
     typeServeur  : longword;
     commentaire  : string;
  end;
  TTabServeur = array of TServeur ;

// =============================================================================
//  Type Tgroupe0
// =============================================================================

  Tgroupe0 = record
    nom : string;
  end;
  TTabGroupe0 = array of tgroupe0 ;

const

  {$EXTERNALSYM LongueurMaxi}
  LongueurMaxi = longword(-1);

// =============================================================================
//  Constantes de filtre pour la fonction NetuserEnum
// =============================================================================

  fltCompteDuplique         = $0001;
  {$EXTERNALSYM fltCompteDuplique}
  fltCompteNormal           = $0002;             // Liste que les utilisateurs
  {$EXTERNALSYM fltCompteNormal}
  fltCompteProxy            = $0004;             // Plus utilisé semble tous afficher
  {$EXTERNALSYM fltCompteProxy}
  fltCompteDomaineConfiance = $0008;             // Liste les domaines "trusted"
  {$EXTERNALSYM fltCompteDomaineConfiance}
  fltCompteMachine          = $0010;             // Liste tous les ordinateurs hors DC
  {$EXTERNALSYM fltCompteMachine}
  fltCompteServeurDomaine   = $0020;             // Liste les PDC et BDC
  {$EXTERNALSYM fltCompteServeurDomaine}

// =============================================================================
//  Constantes des proprietes d'un utilisateur
// =============================================================================

  UF_SCRIPT                 = $0001;
  UF_ACCOUNTDISABLE         = $0002;
  UF_HOMEDIR_REQUIRED       = $0008;
  UF_LOCKOUT                = $0010;
  UF_PASSWD_NOTREQD         = $0020;
  UF_PASSWD_CANT_CHANGE     = $0040;
  UF_TEMP_DUPLICATE_ACCOUNT = $0100;
  UF_NORMAL_ACCOUNT         = $0200;
  UF_DONT_EXPIRE_PASSWD     = $10000;
  USER_PRIV_USER            = 1;
// =============================================================================
//  Constantes de type de serveur permet de faire des filtres de recherche ou
//  de determiner le type de serveur
// =============================================================================

  srvTypeOrdinateur           = $00000001;           // Ordinateur
  {$EXTERNALSYM srvTypeOrdinateur}
  srvTypeServeur              = $00000002;           // Serveur
  {$EXTERNALSYM srvTypeServeur}
  srvTypeServeurSQL           = $00000004;           // Serveur SQL
  {$EXTERNALSYM srvTypeServeurSQL}
  srvTypePDC                  = $00000008;           // Controleur principale de domaine
  {$EXTERNALSYM srvTypePDC}
  srvTypeBDC                  = $00000010;           // Controleur secondaire de domaine
  {$EXTERNALSYM srvTypeBDC}
  srvTypeSourceTemps          = $00000020;           // Serveur de syncronisation du temps
  {$EXTERNALSYM srvTypeSourceTemps}
  srvTypeAFP                  = $00000040;           // Apple File Protocol server
  {$EXTERNALSYM srvTypeAFP}
  srvTypeNovell               = $00000080;           // serveur novell
  {$EXTERNALSYM srvTypeNovell}
  srvTypeMembreDomaine        = $00000100;           // Membre du domaine
  {$EXTERNALSYM srvTypeMembreDomaine}
  srvTypeServeurImpression    = $00000200;           // serveur d'impression
  {$EXTERNALSYM srvTypeServeurImpression}
  srvTypeServeurDialin        = $00000400;           // serveur executant le service Dial-in
  {$EXTERNALSYM srvTypeServeurDialin}
  srvTypeServeurXenix         = $00000800;           // serveur Xenix
  {$EXTERNALSYM srvTypeServeurXenix}
  srvTypeServeurUNIX          = srvTypeServeurXenix; // serveur Unix
  {$EXTERNALSYM srvTypeServeurUNIX}
  srvTypeNT                   = $00001000;           // Serveur 2003 , windows XP, Windows 2000 et NT
  {$EXTERNALSYM srvTypeNT}
  srvTypeWFW                  = $00002000;           // Serveur executant windows pour workgroup
  {$EXTERNALSYM srvTypeWFW}
  srvTypeMFPN                 = $00004000;           // Microsoft File and Print for NetWare
  {$EXTERNALSYM srvTypeMFPN}
  srvTypeServeurNT            = $00008000;           // Serveur n'etant pas controleur de domaine
  {$EXTERNALSYM srvTypeServeurNT}
  srvTypeExplorateurPossible  = $00010000;           // Serveur qui peut executer le service explorateur
  {$EXTERNALSYM srvTypeExplorateurPossible}
  srvTypeExplorateurBackup    = $00020000;           // Serveur qui execute le service explorateur en Backup
  {$EXTERNALSYM srvTypeExplorateurBackup}
  srvTypeExplorateurMaitre    = $00040000;           // Serveur qui execute le service explorateur en tant que maitre
  {$EXTERNALSYM srvTypeExplorateurMaitre}
  srvTypeMaitreDomaine        = $00080000;           // Serveur qui execute le service explorateur en tant que domaine maitre
  {$EXTERNALSYM srvTypeMaitreDomaine}
  srvTypeServeurOSF           = $00100000;           //  ??
  {$EXTERNALSYM srvTypeServeurOSF}
  srvTypeServeurVMS           = $00200000;           //  ??
  {$EXTERNALSYM srvTypeServeurVMS}
  srvTypeWindows              = $00400000;           // Windows95 et superieur
  {$EXTERNALSYM srvTypeWindows}
  srvTypeDFS                  = $00800000;           // Root of a DFS tree
  {$EXTERNALSYM srvTypeDFS}
  srvTypeClusterNT            = $01000000;           // Serveur en Cluster
  {$EXTERNALSYM srvTypeClusterNT}
  srvTypeTerminalserveur      = $02000000;           // Terminal Serveur
  {$EXTERNALSYM srvTypeTerminalserveur}
  srvTypeClusterVirtuel       = $04000000;           // Cluster de serveur virtuel
  {$EXTERNALSYM srvTypeClusterVirtuel}
  srvTypeDCE                  = $10000000;           // IBM DSS (Directory and Security Services) ou equivalent
  {$EXTERNALSYM srvTypeDCE}
  srvTypeAutreTransport       = $20000000;           // renvoie la liste des autres type de serveur
  {$EXTERNALSYM srvTypeAutreTransport}
  srvTypeListeLocale          = $40000000;           // Renvoie la list local
  {$EXTERNALSYM srvTypeListeLocale}
  srvTypeDomainePrincipale    = $80000000;           // Domaine principale
  {$EXTERNALSYM srvTypeDomainePrincipale}
  srvTypeTous                 = $FFFFFFFF;           //
  {$EXTERNALSYM srvTypeTous}

// =============================================================================
//  Constantes d'erreur
// =============================================================================

  {$EXTERNALSYM errSuccess}
  errSuccess = 0;
  {$EXTERNALSYM errBASE}
  errBASE = 2100;
  {$EXTERNALSYM errNetNotStarted}
  errNetNotStarted       = (2102);	// The workstation driver is not installed.
  {$EXTERNALSYM errUnknownServer}
  errUnknownServer       = (2103);	// The server could not be located.
  {$EXTERNALSYM errShareMem}
  errShareMem            = (2104);	// An internal error occurred.  The network cannot access a shared memory segment.
  {$EXTERNALSYM errNoNetworkResource}
  errNoNetworkResource   = (2105);	// A network resource shortage occurred .
  {$EXTERNALSYM errRemoteOnly}
  errRemoteOnly          = (2106);	// This operation is not supported on workstations.
  {$EXTERNALSYM errDevNotRedirected}
  errDevNotRedirected    = (2107);	  // The device is not connected.
  // 2108 is used for ERROR_CONNECTED_OTHER_PASSWORD
  {$EXTERNALSYM errServerNotStarted}
  errServerNotStarted    = (2114);	// The Server service is not started.
  {$EXTERNALSYM errItemNotFound}
  errItemNotFound        = (2115);	// The queue is empty.
  {$EXTERNALSYM errUnknownDevDir}
  errUnknownDevDir       = (2116);	// The device or directory does not exist.
  {$EXTERNALSYM errRedirectedPath}
  errRedirectedPath      = (2117);	// The operation is invalid on a redirected resource.
  {$EXTERNALSYM errDuplicateShare}
  errDuplicateShare      = (2118);	// The name has already been shared.
  {$EXTERNALSYM errNoRoom}
  errNoRoom              = (2119);	// The server is currently out of the requested resource.
  {$EXTERNALSYM errTooManyItems}
  errTooManyItems        = (2121);	// Requested addition of items exceeds the maximum allowed.
  {$EXTERNALSYM errInvalidMaxUsers}
  errInvalidMaxUsers     = (2122);	// The Peer service supports only two simultaneous users.
  {$EXTERNALSYM errBufTooSmall}
  errBufTooSmall         = (2123);	// The API return buffer is too small.
  {$EXTERNALSYM errRemoteErr}
  errRemoteErr           = (2127);	// A remote API error occurred.
  {$EXTERNALSYM errLanmanIniError}
  errLanmanIniError      = (2131);	// An error occurred when opening or reading the configuration file.
  {$EXTERNALSYM errNetworkError}
  errNetworkError        = (2136);	// A general network error occurred.
  {$EXTERNALSYM errWkstaInconsistentState}
  errWkstaInconsistentState  = (2137);	// The Workstation service is in an inconsistent state. Restart the computer before restarting the Workstation service.
  {$EXTERNALSYM errWkstaNotStarted}
  errWkstaNotStarted     = (2138);	// The Workstation service has not been started.
  {$EXTERNALSYM errBrowserNotStarted}
  errBrowserNotStarted   = (2139);	// The requested information is not available.
  {$EXTERNALSYM errInternalError}
  errInternalError       = (2140);	// An internal Windows NT error occurred.
  {$EXTERNALSYM errBadTransactConfig}
  errBadTransactConfig   = (2141);	// The server is not configured for transactions.
  {$EXTERNALSYM errInvalidAPI}
  errInvalidAPI          = (2142);	// The requested API is not supported on the remote server.
  {$EXTERNALSYM errBadEventName}
  errBadEventName        = (2143);	// The event name is invalid.
  {$EXTERNALSYM errDupNameReboot}
  errDupNameReboot       = (2144);	// The computer name already exists on the network. Change it and restart the computer.
// Config API related
// Error codes from BASE+45 to BASE+49
  {$EXTERNALSYM errCfgCompNotFound}
  errCfgCompNotFound     = (2146);	// The specified component could not be found in the configuration information.
  {$EXTERNALSYM errCfgParamNotFound}
  errCfgParamNotFound    = (2147);	// The specified parameter could not be found in the configuration information.
  {$EXTERNALSYM errLineTooLong}
  errLineTooLong         = (2149);	// A line in the configuration file is too long.
// Spooler API related
// Error codes from BASE+50 to BASE+79
  {$EXTERNALSYM errQNotFound}
  errQNotFound           = (2150);	// The printer does not exist.
  {$EXTERNALSYM errJobNotFound}
  errJobNotFound         = (2151);	// The print job does not exist.
  {$EXTERNALSYM errDestNotFound}
  errDestNotFound        = (2152);	// The printer destination cannot be found.
  {$EXTERNALSYM errDestExists}
  errDestExists          = (2153);	// The printer destination already exists.
  {$EXTERNALSYM errQExists}
  errQExists             = (2154);	// The printer queue already exists.
  {$EXTERNALSYM errQNoRoom}
  errQNoRoom             = (2155);	// No more printers can be added.
  {$EXTERNALSYM errJobNoRoom}
  errJobNoRoom           = (2156);	// No more print jobs can be added.
  {$EXTERNALSYM errDestNoRoom}
  errDestNoRoom          = (2157);	// No more printer destinations can be added.
  {$EXTERNALSYM errDestIdle}
  errDestIdle            = (2158);	// This printer destination is idle and cannot accept control operations.
  {$EXTERNALSYM errDestInvalidOp}
  errDestInvalidOp       = (2159);	// This printer destination request contains an invalid control function.
  {$EXTERNALSYM errProcNoRespond}
  errProcNoRespond       = (2160);	// The print processor is not responding.
  {$EXTERNALSYM errSpoolerNotLoaded}
  errSpoolerNotLoaded    = (2161);	// The spooler is not running.
  {$EXTERNALSYM errDestInvalidState}
  errDestInvalidState    = (2162);	// This operation cannot be performed on the print destination in its current state.
  {$EXTERNALSYM errQInvalidState}
  errQInvalidState       = (2163);	// This operation cannot be performed on the printer queue in its current state.
  {$EXTERNALSYM errJobInvalidState}
  errJobInvalidState     = (2164);	// This operation cannot be performed on the print job in its current state.
  {$EXTERNALSYM errSpoolNoMemory}
  errSpoolNoMemory       = (2165);	// A spooler memory allocation failure occurred.
  {$EXTERNALSYM errDriverNotFound}
  errDriverNotFound      = (2166);	// The device driver does not exist.
  {$EXTERNALSYM errDataTypeInvalid}
  errDataTypeInvalid     = (2167);	// The data type is not supported by the print processor.
  {$EXTERNALSYM errProcNotFound}
  errProcNotFound        = (2168);	// The print processor is not installed.
// Service API related
// Error codes from BASE+80 to BASE+99
  {$EXTERNALSYM errServiceTableLocked}
  errServiceTableLocked  = (2180);	// The service database is locked.
  {$EXTERNALSYM errServiceTableFull}
  errServiceTableFull    = (2181);	// The service table is full.
  {$EXTERNALSYM errServiceInstalled}
  errServiceInstalled    = (2182);	// The requested service has already been started.
  {$EXTERNALSYM errServiceEntryLocked}
  errServiceEntryLocked  = (2183);	// The service does not respond to control actions.
  {$EXTERNALSYM errServiceNotInstalled}
  errServiceNotInstalled  = (2184);	// The service has not been started.
  {$EXTERNALSYM errBadServiceName}
  errBadServiceName      = (2185);	// The service name is invalid.
  {$EXTERNALSYM errServiceCtlTimeout}
  errServiceCtlTimeout   = (2186);	// The service is not responding to the control function.
  {$EXTERNALSYM errServiceCtlBusy}
  errServiceCtlBusy      = (2187);	// The service control is busy.
  {$EXTERNALSYM errBadServiceProgName}
  errBadServiceProgName  = (2188);	// The configuration file contains an invalid service program name.
  {$EXTERNALSYM errServiceNotCtrl}
  errServiceNotCtrl      = (2189);	// The service could not be controlled in its present state.
  {$EXTERNALSYM errServiceKillProc}
  errServiceKillProc     = (2190);	// The service ended abnormally.
  {$EXTERNALSYM errServiceCtlNotValid}
  errServiceCtlNotValid  = (2191);	// The requested pause or stop is not valid for this service.
  {$EXTERNALSYM errNotInDispatchTbl}
  errNotInDispatchTbl    = (2192);	// The service control dispatcher could not find the service name in the dispatch table.
  {$EXTERNALSYM errBadControlRecv}
  errBadControlRecv      = (2193);	// The service control dispatcher pipe read failed.
  {$EXTERNALSYM errServiceNotStarting}
  errServiceNotStarting  = (2194);	// A thread for the new service could not be created.
// Wksta and Logon API related
// Error codes from BASE+100 to BASE+118
  {$EXTERNALSYM errAlreadyLoggedOn}
  errAlreadyLoggedOn     = (2200);	// This workstation is already logged on to the local-area network.
  {$EXTERNALSYM errNotLoggedOn}
  errNotLoggedOn         = (2201);	// The workstation is not logged on to the local-area network.
  {$EXTERNALSYM errBadUsername}
  errBadUsername         = (2202);	// The user name or group name parameter is invalid.
  {$EXTERNALSYM errBadPassword}
  errBadPassword         = (2203);	// The password parameter is invalid.
  {$EXTERNALSYM errUnableToAddName_W}
  errUnableToAddName_W   = (2204);	// @W The logon processor did not add the message alias.
  {$EXTERNALSYM errUnableToAddName_F}
  errUnableToAddName_F   = (2205);	// The logon processor did not add the message alias.
  {$EXTERNALSYM errUnableToDelName_W}
  errUnableToDelName_W   = (2206);	// @W The logoff processor did not delete the message alias.
  {$EXTERNALSYM errUnableToDelName_F}
  errUnableToDelName_F   = (2207);	// The logoff processor did not delete the message alias.
  {$EXTERNALSYM errLogonsPaused}
  errLogonsPaused        = (2209);	// Network logons are paused.
  {$EXTERNALSYM errLogonServerConflict}
  errLogonServerConflict  = (2210);	// A centralized logon-server conflict occurred.
  {$EXTERNALSYM errLogonNoUserPath}
  errLogonNoUserPath     = (2211);	// The server is configured without a valid user path.
  {$EXTERNALSYM errLogonScriptError}
  errLogonScriptError    = (2212);	// An error occurred while loading or running the logon script.
  {$EXTERNALSYM errStandaloneLogon}
  errStandaloneLogon     = (2214);	// The logon server was not specified.  Your computer will be logged on as STANDALONE.
  {$EXTERNALSYM errLogonServerNotFound}
  errLogonServerNotFound  = (2215);	// The logon server could not be found.
  {$EXTERNALSYM errLogonDomainExists}
  errLogonDomainExists   = (2216);	// There is already a logon domain for this computer.
  {$EXTERNALSYM errNonValidatedLogon}
  errNonValidatedLogon   = (2217);	  // The logon server could not validate the logon.
// ACF API related (access, user, group);
// Error codes from BASE+119 to BASE+149
  {$EXTERNALSYM errACFNotFound}
  errACFNotFound         = (2219);	// The security database could not be found.
  {$EXTERNALSYM errGroupNotFound}
  errGroupNotFound       = (2220);	// The group name could not be found.
  {$EXTERNALSYM errUserNotFound}
  errUserNotFound        = (2221);	// The user name could not be found.
  {$EXTERNALSYM errResourceNotFound}
  errResourceNotFound    = (2222);	// The resource name could not be found.
  {$EXTERNALSYM errGroupExists}
  errGroupExists         = (2223);	// The group already exists.
  {$EXTERNALSYM errUserExists}
  errUserExists          = (2224);	// The user account already exists.
  {$EXTERNALSYM errResourceExists}
  errResourceExists      = (2225);	// The resource permission list already exists.
  {$EXTERNALSYM errNotPrimary}
  errNotPrimary          = (2226);	// This operation is only allowed on the primary domain controller of the domain.
  {$EXTERNALSYM errACFNotLoaded}
  errACFNotLoaded        = (2227);	// The security database has not been started.
  {$EXTERNALSYM errACFNoRoom}
  errACFNoRoom           = (2228);	// There are too many names in the user accounts database.
  {$EXTERNALSYM errACFFileIOFail}
  errACFFileIOFail       = (2229);	// A disk I/O failure occurred.
  {$EXTERNALSYM errACFTooManyLists}
  errACFTooManyLists     = (2230);	// The limit of 64 entries per resource was exceeded.
  {$EXTERNALSYM errUserLogon}
  errUserLogon           = (2231);	// Deleting a user with a session is not allowed.
  {$EXTERNALSYM errACFNoParent}
  errACFNoParent         = (2232);	// The parent directory could not be located.
  {$EXTERNALSYM errCanNotGrowSegment}
  errCanNotGrowSegment   = (2233);	// Unable to add to the security database session cache segment.
  {$EXTERNALSYM errSpeGroupOp}
  errSpeGroupOp          = (2234);	// This operation is not allowed on this special group.
  {$EXTERNALSYM errNotInCache}
  errNotInCache          = (2235);	// This user is not cached in user accounts database session cache.
  {$EXTERNALSYM errUserInGroup}
  errUserInGroup         = (2236);	// The user already belongs to this group.
  {$EXTERNALSYM errUserNotInGroup}
  errUserNotInGroup      = (2237);	// The user does not belong to this group.
  {$EXTERNALSYM errAccountUndefined}
  errAccountUndefined    = (2238);	// This user account is undefined.
  {$EXTERNALSYM errAccountExpired}
  errAccountExpired      = (2239);	// This user account has expired.
  {$EXTERNALSYM errInvalidWorkstation}
  errInvalidWorkstation  = (2240);	// The user is not allowed to log on from this workstation.
  {$EXTERNALSYM errInvalidLogonHours}
  errInvalidLogonHours   = (2241);	// The user is not allowed to log on at this time.
  {$EXTERNALSYM errPasswordExpired}
  errPasswordExpired     = (2242);	// The password of this user has expired.
  {$EXTERNALSYM errPasswordCantChange}
  errPasswordCantChange  = (2243);	// The password of this user cannot change.
  {$EXTERNALSYM errPasswordHistConflict}
  errPasswordHistConflict  = (2244);	// This password cannot be used now.
  {$EXTERNALSYM errPasswordTooShort}
  errPasswordTooShort    = (2245);	// The password is shorter than required.
  {$EXTERNALSYM errPasswordTooRecent}
  errPasswordTooRecent   = (2246);	// The password of this user is too recent to change.
  {$EXTERNALSYM errInvalidDatabase}
  errInvalidDatabase     = (2247);	// The security database is corrupted.
  {$EXTERNALSYM errDatabaseUpToDate}
  errDatabaseUpToDate    = (2248);	// No updates are necessary to this replicant network/local security database.
  {$EXTERNALSYM errSyncRequired}
  errSyncRequired        = (2249);	  // This replicant database is outdated; synchronization is required.
// Use API related
// Error codes from BASE+150 to BASE+169
  {$EXTERNALSYM errUseNotFound}
  errUseNotFound         = (2250);	// The network connection could not be found.
  {$EXTERNALSYM errBadAsgType}
  errBadAsgType          = (2251);	// This asg_type is invalid.
  {$EXTERNALSYM errDeviceIsShared}
  errDeviceIsShared      = (2252);	// This device is currently being shared.
// Message Server related
// Error codes BASE+170 to BASE+209
  {$EXTERNALSYM errNoComputerName}
  errNoComputerName      = (2270);	// The computer name could not be added as a message alias.  The name may already exist on the network.
  {$EXTERNALSYM errMsgAlreadyStarted}
  errMsgAlreadyStarted   = (2271);	// The Messenger service is already started.
  {$EXTERNALSYM errMsgInitFailed}
  errMsgInitFailed       = (2272);	// The Messenger service failed to start.
  {$EXTERNALSYM errNameNotFound}
  errNameNotFound        = (2273);	// The message alias could not be found on the network.
  {$EXTERNALSYM errAlreadyForwarded}
  errAlreadyForwarded    = (2274);	// This message alias has already been forwarded.
  {$EXTERNALSYM errAddForwarded}
  errAddForwarded        = (2275);	// This message alias has been added but is still forwarded.
  {$EXTERNALSYM errAlreadyExists}
  errAlreadyExists       = (2276);	// This message alias already exists locally.
  {$EXTERNALSYM errTooManyNames}
  errTooManyNames        = (2277);	// The maximum number of added message aliases has been exceeded.
  {$EXTERNALSYM errDelComputerName}
  errDelComputerName     = (2278);	// The computer name could not be deleted.
  {$EXTERNALSYM errLocalForward}
  errLocalForward        = (2279);	// Messages cannot be forwarded back to the same workstation.
  {$EXTERNALSYM errGrpMsgProcessor}
  errGrpMsgProcessor     = (2280);	// An error occurred in the domain message processor.
  {$EXTERNALSYM errPausedRemote}
  errPausedRemote        = (2281);	// The message was sent, but the recipient has paused the Messenger service.
  {$EXTERNALSYM errBadReceive}
  errBadReceive          = (2282);	// The message was sent but not received.
  {$EXTERNALSYM errNameInUse}
  errNameInUse           = (2283);	// The message alias is currently in use. Try again later.
  {$EXTERNALSYM errMsgNotStarted}
  errMsgNotStarted       = (2284);	// The Messenger service has not been started.
  {$EXTERNALSYM errNotLocalName}
  errNotLocalName        = (2285);	// The name is not on the local computer.
  {$EXTERNALSYM errNoForwardName}
  errNoForwardName       = (2286);	// The forwarded message alias could not be found on the network.
  {$EXTERNALSYM errRemoteFull}
  errRemoteFull          = (2287);	// The message alias table on the remote station is full.
  {$EXTERNALSYM errNameNotForwarded}
  errNameNotForwarded    = (2288);	// Messages for this alias are not currently being forwarded.
  {$EXTERNALSYM errTruncatedBroadcast}
  errTruncatedBroadcast  = (2289);	// The broadcast message was truncated.
  {$EXTERNALSYM errInvalidDevice}
  errInvalidDevice       = (2294);	// This is an invalid device name.
  {$EXTERNALSYM errWriteFault}
  errWriteFault          = (2295);	// A write fault occurred.
  {$EXTERNALSYM errDuplicateName}
  errDuplicateName       = (2297);	// A duplicate message alias exists on the network.
  {$EXTERNALSYM errDeleteLater}
  errDeleteLater         = (2298);	// @W This message alias will be deleted later.
  {$EXTERNALSYM errIncompleteDel}
  errIncompleteDel       = (2299);	// The message alias was not successfully deleted from all networks.
  {$EXTERNALSYM errMultipleNets}
  errMultipleNets        = (2300);	// This operation is not supported on computers with multiple networks.
// Server API related
// Error codes BASE+210 to BASE+229
  {$EXTERNALSYM errNetNameNotFound}
  errNetNameNotFound     = (2310);	// This shared resource does not exist.
  {$EXTERNALSYM errDeviceNotShared}
  errDeviceNotShared     = (2311);	// This device is not shared.
  {$EXTERNALSYM errClientNameNotFound}
  errClientNameNotFound  = (2312);	// A session does not exist with that computer name.
  {$EXTERNALSYM errFileIdNotFound}
  errFileIdNotFound      = (2314);	// There is not an open file with that identification number.
  {$EXTERNALSYM errExecFailure}
  errExecFailure         = (2315);	// A failure occurred when executing a remote administration command.
  {$EXTERNALSYM errTmpFile}
  errTmpFile             = (2316);	// A failure occurred when opening a remote temporary file.
  {$EXTERNALSYM errTooMuchData}
  errTooMuchData         = (2317);	// The data returned from a remote administration command has been truncated to 64K.
  {$EXTERNALSYM errDeviceShareConflict}
  errDeviceShareConflict  = (2318);	// This device cannot be shared as both a spooled and a non-spooled resource.
  {$EXTERNALSYM errBrowserTableIncomplete}
  errBrowserTableIncomplete  = (2319);	 // The information in the list of servers may be incorrect.
  {$EXTERNALSYM errNotLocalDomain}
  errNotLocalDomain      = (2320);	// The computer is not active in this domain.
  {$EXTERNALSYM errIsDfsShare}
  errIsDfsShare          = (2321);	// The share must be removed from the Distributed File System before it can be deleted.
// CharDev API related
// Error codes BASE+230 to BASE+249
  {$EXTERNALSYM errDevInvalidOpCode}
  errDevInvalidOpCode    = (2331);	// The operation is invalid for this device.
  {$EXTERNALSYM errDevNotFound}
  errDevNotFound         = (2332);	// This device cannot be shared.
  {$EXTERNALSYM errDevNotOpen}
  errDevNotOpen          = (2333);	// This device was not open.
  {$EXTERNALSYM errBadQueueDevString}
  errBadQueueDevString   = (2334);	// This device name list is invalid.
  {$EXTERNALSYM errBadQueuePriority}
  errBadQueuePriority    = (2335);	// The queue priority is invalid.
  {$EXTERNALSYM errNoCommDevs}
  errNoCommDevs          = (2337);	// There are no shared communication devices.
  {$EXTERNALSYM errQueueNotFound}
  errQueueNotFound       = (2338);	// The queue you specified does not exist.
  {$EXTERNALSYM errBadDevString}
  errBadDevString        = (2340);	// This list of devices is invalid.
  {$EXTERNALSYM errBadDev}
  errBadDev              = (2341);	// The requested device is invalid.
  {$EXTERNALSYM errInUseBySpooler}
  errInUseBySpooler      = (2342);	// This device is already in use by the spooler.
  {$EXTERNALSYM errCommDevInUse}
  errCommDevInUse        = (2343);	// This device is already in use as a communication device.
// NetICanonicalize and NetIType and NetIMakeLMFileName
// NetIListCanon and NetINameCheck
// Error codes BASE+250 to BASE+269
  {$EXTERNALSYM errInvalidComputer}
  errInvalidComputer    = (2351);	// This computer name is invalid.
  {$EXTERNALSYM errMaxLenExceeded}
  errMaxLenExceeded     = (2354);	// The string and prefix specified are too long.
  {$EXTERNALSYM errBadComponent}
  errBadComponent       = (2356);	// This path component is invalid.
  {$EXTERNALSYM errCantType}
  errCantType           = (2357);	// Could not determine the type of input.
  {$EXTERNALSYM errTooManyEntries}
  errTooManyEntries     = (2362);	// The buffer for types is not big enough.
// NetProfile
// Error codes BASE+270 to BASE+276
  {$EXTERNALSYM errProfileFileTooBig}
  errProfileFileTooBig   = (2370);	// Profile files cannot exceed 64K.
  {$EXTERNALSYM errProfileOffset}
  errProfileOffset       = (2371);	// The start offset is out of range.
  {$EXTERNALSYM errProfileCleanup}
  errProfileCleanup      = (2372);	// The system cannot delete current connections to network resources.
  {$EXTERNALSYM errProfileUnknownCmd}
  errProfileUnknownCmd   = (2373);	// The system was unable to parse the command line in this file.
  {$EXTERNALSYM errProfileLoadErr}
  errProfileLoadErr      = (2374);	// An error occurred while loading the profile file.
  {$EXTERNALSYM errProfileSaveErr}
  errProfileSaveErr      = (2375);	// @W Errors occurred while saving the profile file.  The profile was partially saved.
// NetAudit and NetErrorLog
// Error codes BASE+277 to BASE+279
  {$EXTERNALSYM errLogOverflow}
  errLogOverflow            = (2377);	// Log file %1 is full.
  {$EXTERNALSYM errLogFileChanged}
  errLogFileChanged         = (2378);	// This log file has changed between reads.
  {$EXTERNALSYM errLogFileCorrupt}
  errLogFileCorrupt         = (2379);	// Log file %1 is corrupt.
// NetRemote
// Error codes BASE+280 to BASE+299
  {$EXTERNALSYM errSourceIsDir}
  errSourceIsDir    = (2380);	// The source path cannot be a directory.
  {$EXTERNALSYM errBadSource}
  errBadSource      = (2381);	// The source path is illegal.
  {$EXTERNALSYM errBadDest}
  errBadDest        = (2382);	// The destination path is illegal.
  {$EXTERNALSYM errDifferentServers}
  errDifferentServers    = (2383);	// The source and destination paths are on different servers.
  {$EXTERNALSYM errRunSrvPaused}
  errRunSrvPaused        = (2385);	// The Run server you requested is paused.
  {$EXTERNALSYM errErrCommRunSrv}
  errErrCommRunSrv       = (2389);	// An error occurred when communicating with a Run server.
  {$EXTERNALSYM errErrorExecingGhost}
  errErrorExecingGhost   = (2391);	// An error occurred when starting a background process.
  {$EXTERNALSYM errShareNotFound}
  errShareNotFound       = (2392);	// The shared resource you are connected to could not be found.
// NetWksta.sys (redir); returned error codes.
// errBASE + (300-329);
  {$EXTERNALSYM errInvalidLana}
  errInvalidLana         = (2400);	// The LAN adapter number is invalid.
  {$EXTERNALSYM errOpenFiles}
  errOpenFiles           = (2401);	// There are open files on the connection.
  {$EXTERNALSYM errActiveConns}
  errActiveConns         = (2402);	// Active connections still exist.
  {$EXTERNALSYM errBadPasswordCore}
  errBadPasswordCore     = (2403);	// This share name or password is invalid.
  {$EXTERNALSYM errDevInUse}
  errDevInUse            = (2404);	// The device is being accessed by an active process.
  {$EXTERNALSYM errLocalDrive}
  errLocalDrive          = (2405);	// The drive letter is in use locally.
//  Alert error codes.
//  errBASE + (330-339);
  {$EXTERNALSYM errAlertExists}
  errAlertExists         = (2430);	// The specified client is already registered for the specified event.
  {$EXTERNALSYM errTooManyAlerts}
  errTooManyAlerts       = (2431);	// The alert table is full.
  {$EXTERNALSYM errNoSuchAlert}
  errNoSuchAlert         = (2432);	// An invalid or nonexistent alert name was raised.
  {$EXTERNALSYM errBadRecipient}
  errBadRecipient        = (2433);	// The alert recipient is invalid.
  {$EXTERNALSYM errAcctLimitExceeded}
  errAcctLimitExceeded   = (2434);	// A user's session with this server has been deleted because the user's logon hours are no longer valid.
// Additional Error and Audit log codes.
// errBASE +(340-343)
  {$EXTERNALSYM errInvalidLogSeek}
  errInvalidLogSeek      = (2440);	// The log file does not contain the requested record number.
// Additional UAS and NETLOGON codes
// errBASE +(350-359)
  {$EXTERNALSYM errBadUasConfig}
  errBadUasConfig        = (2450);	// The user accounts database is not configured correctly.
  {$EXTERNALSYM errInvalidUASOp}
  errInvalidUASOp        = (2451);	// This operation is not permitted when the Netlogon service is running.
  {$EXTERNALSYM errLastAdmin}
  errLastAdmin           = (2452);	// This operation is not allowed on the last administrative account.
  {$EXTERNALSYM errDCNotFound}
  errDCNotFound          = (2453);	// Could not find domain controller for this domain.
  {$EXTERNALSYM errLogonTrackingError}
  errLogonTrackingError  = (2454);	// Could not set logon information for this user.
  {$EXTERNALSYM errNetlogonNotStarted}
  errNetlogonNotStarted  = (2455);	// The Netlogon service has not been started.
  {$EXTERNALSYM errCanNotGrowUASFile}
  errCanNotGrowUASFile   = (2456);	// Unable to add to the user accounts database.
  {$EXTERNALSYM errTimeDiffAtDC}
  errTimeDiffAtDC        = (2457);	// This server's clock is not synchronized with the primary domain controller's clock.
  {$EXTERNALSYM errPasswordMismatch}
  errPasswordMismatch    = (2458);	// A password mismatch has been detected.
// Server Integration error codes.
// errBASE +(360-369)
  {$EXTERNALSYM errNoSuchServer}
  errNoSuchServer        = (2460);	// The server identification does not specify a valid server.
  {$EXTERNALSYM errNoSuchSession}
  errNoSuchSession       = (2461);	// The session identification does not specify a valid session.
  {$EXTERNALSYM errNoSuchConnection}
  errNoSuchConnection    = (2462);	// The connection identification does not specify a valid connection.
  {$EXTERNALSYM errTooManyServers}
  errTooManyServers      = (2463);	// There is no space for another entry in the table of available servers.
  {$EXTERNALSYM errTooManySessions}
  errTooManySessions     = (2464);	// The server has reached the maximum number of sessions it supports.
  {$EXTERNALSYM errTooManyConnections}
  errTooManyConnections  = (2465);	// The server has reached the maximum number of connections it supports.
  {$EXTERNALSYM errTooManyFiles}
  errTooManyFiles        = (2466);	// The server cannot open more files because it has reached its maximum number.
  {$EXTERNALSYM errNoAlternateServers}
  errNoAlternateServers  = (2467);	// There are no alternate servers registered on this server.
  {$EXTERNALSYM errTryDownLevel}
  errTryDownLevel        = (2470);	// Try down-level (remote admin protocol); version of API instead.
// UPS error codes.
// errBASE + (380-384);
  {$EXTERNALSYM errUPSDriverNotStarted}
  errUPSDriverNotStarted     = (2480);	// The UPS driver could not be accessed by the UPS service.
  {$EXTERNALSYM errUPSInvalidConfig}
  errUPSInvalidConfig        = (2481);	// The UPS service is not configured correctly.
  {$EXTERNALSYM errUPSInvalidCommPort}
  errUPSInvalidCommPort      = (2482);	// The UPS service could not access the specified Comm Port.
  {$EXTERNALSYM errUPSSignalAsserted}
  errUPSSignalAsserted       = (2483);	// The UPS indicated a line fail or low battery situation. Service not started.
  {$EXTERNALSYM errUPSShutdownFailed}
  errUPSShutdownFailed       = (2484);	// The UPS service failed to perform a system shut down.
// Remoteboot error codes.
// errBASE + (400-419);
// Error codes 400 - 405 are used by RPLBOOT.SYS.
// Error codes 403, 407 - 416 are used by RPLLOADR.COM,
// Error code 417 is the alerter message of REMOTEBOOT (RPLSERVR.EXE);.
// Error code 418 is for when REMOTEBOOT can't start
// Error code 419 is for a disallowed 2nd rpl connection
  {$EXTERNALSYM errBadDosRetCode}
  errBadDosRetCode       = (2500);	// The program below returned an MS-DOS error code:
  {$EXTERNALSYM errProgNeedsExtraMem}
  errProgNeedsExtraMem   = (2501);	// The program below needs more memory:
  {$EXTERNALSYM errBadDosFunction}
  errBadDosFunction      = (2502);	// The program below called an unsupported MS-DOS function:
  {$EXTERNALSYM errRemoteBootFailed}
  errRemoteBootFailed    = (2503);	// The workstation failed to boot.
  {$EXTERNALSYM errBadFileCheckSum}
  errBadFileCheckSum     = (2504);	// The file below is corrupt.
  {$EXTERNALSYM errNoRplBootSystem}
  errNoRplBootSystem     = (2505);	// No loader is specified in the boot-block definition file.
  {$EXTERNALSYM errRplLoadrNetBiosErr}
  errRplLoadrNetBiosErr  = (2506);	// NetBIOS returned an error: The NCB and SMB are dumped above.
  {$EXTERNALSYM errRplLoadrDiskErr}
  errRplLoadrDiskErr     = (2507);	// A disk I/O error occurred.
  {$EXTERNALSYM errImageParamErr}
  errImageParamErr       = (2508);	// Image parameter substitution failed.
  {$EXTERNALSYM errTooManyImageParams}
  errTooManyImageParams  = (2509);	// Too many image parameters cross disk sector boundaries.
  {$EXTERNALSYM errNonDosFloppyUsed}
  errNonDosFloppyUsed    = (2510);	// The image was not generated from an MS-DOS diskette formatted with /S.
  {$EXTERNALSYM errRplBootRestart}
  errRplBootRestart      = (2511);	// Remote boot will be restarted later.
  {$EXTERNALSYM errRplSrvrCallFailed}
  errRplSrvrCallFailed   = (2512);	// The call to the Remoteboot server failed.
  {$EXTERNALSYM errCantConnectRplSrvr}
  errCantConnectRplSrvr  = (2513);	// Cannot connect to the Remoteboot server.
  {$EXTERNALSYM errCantOpenImageFile}
  errCantOpenImageFile   = (2514);	// Cannot open image file on the Remoteboot server.
  {$EXTERNALSYM errCallingRplSrvr}
  errCallingRplSrvr      = (2515);	// Connecting to the Remoteboot server...
  {$EXTERNALSYM errStartingRplBoot}
  errStartingRplBoot     = (2516);	// Connecting to the Remoteboot server...
  {$EXTERNALSYM errRplBootServiceTerm}
  errRplBootServiceTerm  = (2517);	// Remote boot service was stopped; check the error log for the cause of the problem.
  {$EXTERNALSYM errRplBootStartFailed}
  errRplBootStartFailed  = (2518);	// Remote boot startup failed; check the error log for the cause of the problem.
  {$EXTERNALSYM errRPL_CONNECTED}
  errRPL_CONNECTED       = (2519);	// A second connection to a Remoteboot resource is not allowed.
// FTADMIN API error codes
// errBASE + (425-434)
// (Currently not used in NT);
// Browser service API error codes
// errBASE + (450-475)
  {$EXTERNALSYM errBrowserConfiguredToNotRun}
  errBrowserConfiguredToNotRun      = (2550);	// The browser service was configured with MaintainServerList=No.
// Additional Remoteboot error codes.
// errBASE + (510-550);
  {$EXTERNALSYM errRplNoAdaptersStarted}
  errRplNoAdaptersStarted           = (2610);	//Service failed to start since none of the network adapters started with this service.
  {$EXTERNALSYM errRplBadRegistry}
  errRplBadRegistry                 = (2611);	//Service failed to start due to bad startup information in the registry.
  {$EXTERNALSYM errRplBadDatabase}
  errRplBadDatabase                 = (2612);	//Service failed to start because its database is absent or corrupt.
  {$EXTERNALSYM errRplRplfilesShare}
  errRplRplfilesShare               = (2613);	//Service failed to start because RPLFILES share is absent.
  {$EXTERNALSYM errRplNotRplServer}
  errRplNotRplServer                = (2614);	//Service failed to start because RPLUSER group is absent.
  {$EXTERNALSYM errRplCannotEnum}
  errRplCannotEnum                  = (2615);	//Cannot enumerate service records.
  {$EXTERNALSYM errRplWkstaInfoCorrupted}
  errRplWkstaInfoCorrupted          = (2616);	//Workstation record information has been corrupted.
  {$EXTERNALSYM errRplWkstaNotFound}
  errRplWkstaNotFound               = (2617);	//Workstation record was not found.
  {$EXTERNALSYM errRplWkstaNameUnavailable}
  errRplWkstaNameUnavailable        = (2618);	//Workstation name is in use by some other workstation.
  {$EXTERNALSYM errRplProfileInfoCorrupted}
  errRplProfileInfoCorrupted        = (2619);	//Profile record information has been corrupted.
  {$EXTERNALSYM errRplProfileNotFound}
  errRplProfileNotFound             = (2620);	//Profile record was not found.
  {$EXTERNALSYM errRplProfileNameUnavailable}
  errRplProfileNameUnavailable      = (2621);	//Profile name is in use by some other profile.
  {$EXTERNALSYM errRplProfileNotEmpty}
  errRplProfileNotEmpty             = (2622);	//There are workstations using this profile.
  {$EXTERNALSYM errRplConfigInfoCorrupted}
  errRplConfigInfoCorrupted         = (2623);	//Configuration record information has been corrupted.
  {$EXTERNALSYM errRplConfigNotFound}
  errRplConfigNotFound              = (2624);	//Configuration record was not found.
  {$EXTERNALSYM errRplAdapterInfoCorrupted}
  errRplAdapterInfoCorrupted        = (2625);	//Adapter id record information has been corrupted.
  {$EXTERNALSYM errRplInternal}
  errRplInternal                    = (2626);	//An internal service error has occurred.
  {$EXTERNALSYM errRplVendorInfoCorrupted}
  errRplVendorInfoCorrupted         = (2627);	//Vendor id record information has been corrupted.
  {$EXTERNALSYM errRplBootInfoCorrupted}
  errRplBootInfoCorrupted           = (2628);	//Boot block record information has been corrupted.
  {$EXTERNALSYM errRplWkstaNeedsUserAcct}
  errRplWkstaNeedsUserAcct          = (2629);	//The user account for this workstation record is missing.
  {$EXTERNALSYM errRplNeedsRPLUSERAcct}
  errRplNeedsRPLUSERAcct            = (2630);	//The RPLUSER local group could not be found.
  {$EXTERNALSYM errRplBootNotFound}
  errRplBootNotFound                = (2631);	//Boot block record was not found.
  {$EXTERNALSYM errRplIncompatibleProfile}
  errRplIncompatibleProfile         = (2632);	//Chosen profile is incompatible with this workstation.
  {$EXTERNALSYM errRplAdapterNameUnavailable}
  errRplAdapterNameUnavailable      = (2633);	//Chosen network adapter id is in use by some other workstation.
  {$EXTERNALSYM errRplConfigNotEmpty}
  errRplConfigNotEmpty              = (2634);	//There are profiles using this configuration.
  {$EXTERNALSYM errRplBootInUse}
  errRplBootInUse                   = (2635);	//There are workstations, profiles or configurations using this boot block.
  {$EXTERNALSYM errRplBackupDatabase}
  errRplBackupDatabase              = (2636);	//Service failed to backup Remoteboot database.
  {$EXTERNALSYM errRplAdapterNotFound}
  errRplAdapterNotFound             = (2637);	//Adapter record was not found.
  {$EXTERNALSYM errRplVendorNotFound}
  errRplVendorNotFound              = (2638);	//Vendor record was not found.
  {$EXTERNALSYM errRplVendorNameUnavailable}
  errRplVendorNameUnavailable       = (2639);	//Vendor name is in use by some other vendor record.
  {$EXTERNALSYM errRplBootNameUnavailable}
  errRplBootNameUnavailable         = (2640);	//(boot name, vendor id); is in use by some other boot block record.
  {$EXTERNALSYM errRplConfigNameUnavailable}
  errRplConfigNameUnavailable       = (2641);	//Configuration name is in use by some other configuration.
// Dfs API error codes.
// errBASE + (560-590);
  {$EXTERNALSYM errDfsInternalCorruption}
  errDfsInternalCorruption          = (2660);	//The internal database maintained by the Dfs service is corrupt
  {$EXTERNALSYM errDfsVolumeDataCorrupt}
  errDfsVolumeDataCorrupt           = (2661);	//One of the records in the internal Dfs database is corrupt
  {$EXTERNALSYM errDfsNoSuchVolume}
  errDfsNoSuchVolume                = (2662);	//There is no volume whose entry path matches the input Entry Path
  {$EXTERNALSYM errDfsVolumeAlreadyExists}
  errDfsVolumeAlreadyExists         = (2663);	//A volume with the given name already exists
  {$EXTERNALSYM errDfsAlreadyShared}
  errDfsAlreadyShared               = (2664);	//The server share specified is already shared in the Dfs
  {$EXTERNALSYM errDfsNoSuchShare}
  errDfsNoSuchShare                 = (2665);	//The indicated server share does not support the indicated Dfs volume
  {$EXTERNALSYM errDfsNotALeafVolume}
  errDfsNotALeafVolume              = (2666);	//The operation is not valid on a non-leaf volume
  {$EXTERNALSYM errDfsLeafVolume}
  errDfsLeafVolume                  = (2667);	//The operation is not valid on a leaf volume
  {$EXTERNALSYM errDfsVolumeHasMultipleServers}
  errDfsVolumeHasMultipleServers    = (2668);	//The operation is ambiguous because the volume has multiple servers
  {$EXTERNALSYM errDfsCantCreateJunctionPoint}
  errDfsCantCreateJunctionPoint     = (2669);	//Unable to create a junction point
  {$EXTERNALSYM errDfsServerNotDfsAware}
  errDfsServerNotDfsAware           = (2670);	//The server is not Dfs Aware
  {$EXTERNALSYM errDfsBadRenamePath}
  errDfsBadRenamePath               = (2671);	//The specified rename target path is invalid
  {$EXTERNALSYM errDfsVolumeIsOffline}
  errDfsVolumeIsOffline             = (2672);	//The specified Dfs volume is offline
  {$EXTERNALSYM errDfsNoSuchServer}
  errDfsNoSuchServer                = (2673);	//The specified server is not a server for this volume
  {$EXTERNALSYM errDfsCyclicalName}
  errDfsCyclicalName                = (2674);	//A cycle in the Dfs name was detected
  {$EXTERNALSYM errDfsNotSupportedInServerDfs}
  errDfsNotSupportedInServerDfs     = (2675);	//The operation is not supported on a server-based Dfs
  {$EXTERNALSYM errDfsDuplicateService}
  errDfsDuplicateService            = (2676);	//This volume is already supported by the specified server-share
  {$EXTERNALSYM errDfsCantRemoveLastServerShare}
  errDfsCantRemoveLastServerShare   = (2677);	//Can't remove the last server-share supporting this volume
  {$EXTERNALSYM errDfsVolumeIsInterDfs}
  errDfsVolumeIsInterDfs            = (2678);	//The operation is not supported for an Inter-Dfs volume
  {$EXTERNALSYM errDfsInconsistent}
  errDfsInconsistent                = (2679);	//The internal state of the Dfs Service has become inconsistent
  {$EXTERNALSYM errDfsServerUpgraded}
  errDfsServerUpgraded              = (2680);	//The Dfs Service has been installed on the specified server
  {$EXTERNALSYM errDfsDataIsIdentical}
  errDfsDataIsIdentical             = (2681);	//The Dfs data being reconciled is identical
  {$EXTERNALSYM errDfsCantRemoveDfsRoot}
  errDfsCantRemoveDfsRoot           = (2682);	//The Dfs root volume cannot be deleted - Uninstall Dfs if required
  {$EXTERNALSYM errDfsChildOrParentInDfs}
  errDfsChildOrParentInDfs          = (2683);	//A child or parent directory of the share is already in a Dfs
  {$EXTERNALSYM errDfsInternalError}
  errDfsInternalError               = (2690);	//Dfs internal error
// Net setup error codes.
// errBASE + (591-600);
  {$EXTERNALSYM errSetupAlreadyJoined}
  errSetupAlreadyJoined             = (2691);	//This machine is already joined to a domain.
  {$EXTERNALSYM errSetupNotJoined}
  errSetupNotJoined                 = (2692);	//This machine is not currently joined to a domain.
  {$EXTERNALSYM errSetupDomainController}
  errSetupDomainController          = (2693);	//This machine is a domain controller and cannot be unjoined from a domain.
  {$EXTERNALSYM errDefaultJoinRequired}
  errDefaultJoinRequired            = (2694);	//*The destination domain controller does not support creating machine accounts in OUs.
  {$EXTERNALSYM errInvalidWorkgroupName}
  errInvalidWorkgroupName           = (2695);	//*The specified workgroup name is invalid
  {$EXTERNALSYM errNameUsesIncompatibleCodePage}
  errNameUsesIncompatibleCodePage   = (2696);	//*The specified computer name is incompatible with the default language used on the domain controller.
  {$EXTERNALSYM errComputerAccountNotFound}
  errComputerAccountNotFound        = (2697);	//*The specified computer account could not be found.
  {$EXTERNALSYM errMax}
  errMax                            = (2999);	// This is the last error in NERR range.

  STYPE_DISKTREE = 0;
  STYPE_PRINTQ   = 1;
  STYPE_DEVICE   = 2;
  STYPE_IPC      = 3;
  STYPE_DFS      = 100;
  STYPE_SPECIAL  = $80000000;

  ACCESS_NONE   = 0;
  ACCESS_READ   = 1;
  ACCESS_WRITE  = 2;
  ACCESS_CREATE = 4;
  ACCESS_EXEC   = 8;
  ACCESS_DELETE = 16;
  ACCESS_ATRIB  = 32;
  ACCESS_PERM   = 64;
  ACCESS_ALL    = (ACCESS_READ+ACCESS_WRITE+ACCESS_CREATE+ACCESS_EXEC+ACCESS_DELETE+ACCESS_ATRIB+ACCESS_PERM);

// =============================================================================
//  Déclaration des fonctions
// =============================================================================

Function  CreateUser (const strNomServeur : string ; Information : Tutilisateur ) : integer;
Function  GetinfoUtilisateur    (const strNomServeur : string ; const strNomUtilisateur : string ; var Information : Tutilisateur ) : integer;
Function  SetUserHomeFolder (const strNomServeur : string ; const strNomUtilisateur : string ; const strLecteur : string;const strPartage : string ):integer;
Function  GetListeComptes  (const strNomServeur : string; fltType : longword ; var tabUtilisateur : TTabUtilisateur;posi : tprogressbar =nil ) : integer;
Function  GetListeGroupeGlobaux (const strNomServeur : string ; var tabGroupeGlobaux : TTabGroupeGlobaux ) : integer;
Function  GetListeGroupeLocaux  (const strNomServeur : string ; var tabGroupeLocaux : TTabGroupeLoCaux ) : integer;
Function  GetListeUtilisateursGroupeGlobal (const strNomServeur : string ;const strNomGroupe : string; var tabUtilisateurs : TTabUtilisateur0 ) : integer;overload;
Function  GetListeUtilisateursGroupeGlobal (const strNomServeur : string ;const strNomGroupe : string; var tabUtilisateurs : TTabUtilisateur;posi : tprogressbar =nil) : integer;overload;
Function  GetNomDomaine ( var nomDomaine : string ;var EtatIntegration : longword): integer;
function  GetNomTousPdc ( strServeur : string;strDomaine : string ; var listepdc : string ):integer;
function  GetNomPdc ( strServeur : string;strDomaine : string ; var listepdc : string ):integer;
function  GetListeServeur (const strNomdomaine :String ;srvType : longword ;var tabServeur : TTabServeur ) : integer;
function  GetListeDisque (const strNomServeur : string ; Var ListeDisque : TTabDisque ) : integer;
function  GetListeGroupeUtilisateur ( const strNomServeur : string; const strNomUtilisateur : string ; var tabGroupe : TTabGroupe0) :integer;

function  AjouteUtilisateurGroupeLocal (const strNomServeur : string; const strNomUtilisateur : string ;const strNomGoupe : string  ) : integer;
function  AjouteUtilisateurGroupe (const strNomServeur : string; const strNomUtilisateur : string ;const strNomGoupe : string  ) : integer;
function  SupprimeUtilisateurGroupe (const strNomServeur : string; const strNomUtilisateur : string ;const strNomGoupe : string  ) : integer;

function  booPing(InetAddress : string) : boolean;

function  PartageToChemin(strServeur : string ; strPartage : string ; var strChemin : string) : boolean;
function  CreerPartage(strServeur :string; strChemin : string; strPartage : string) : boolean;
function  ArretPartage(strserveur : string; strpartage : string):boolean;
function  ListerPartage(strserveur : string;var tabListePartage : TTabPartage ):boolean;

function accessmodetostr(accessmode : ACCESS_MODE ):string;
function accesspermissiontoStr(accesspermission : cardinal ) : string;

function  GetUfbReseauErreur : string;

// -----------------------------------------------------------------------------

implementation

uses sysutils,windows, Classes,forms,WinSock,aclapi;

type
   tnomGroupe = record
    strnom : string;
    strdomaine : string;
  end;
  TConvertSidToStringSidA = function (Sid: PSID;
    var StringSid: LPSTR): BOOL; stdcall;

// =============================================================================
//  Type utilisés par la fonction de ping                                   
// =============================================================================

  TPaquetsB = packed record
    s_b1, s_b2, s_b3, s_b4: byte;
  end;

  TPaquetsW = packed record
    s_w1, s_w2: word;
  end;

  PIPAddr = ^TIPAddr;
  TIPAddr = record
    case integer of
      0: (S_un_b: TPaquetsB);
      1: (S_un_w: TPaquetsW);
      2: (S_addr: longword);
  end;

  IPAddr = TIPAddr;

  Pdisque = ^Tdisque;
// =============================================================================
//  Type PUserInfo1
// =============================================================================
  PUserInfo1 = ^TUserInfo1;
  {$EXTERNALSYM _USER_INFO_1}
  _USER_INFO_1 = record
    usri1_name: LPWSTR;
    usri1_password: LPWSTR;
    usri1_password_age: DWORD;
    usri1_priv: DWORD;
    usri1_home_dir: LPWSTR;
    usri1_comment: LPWSTR;
    usri1_flags: DWORD;
    usri1_script_path: LPWSTR;
  end;
  TUserInfo1 = _USER_INFO_1;
  {$EXTERNALSYM USER_INFO_1}
  USER_INFO_1 = _USER_INFO_1;

// =============================================================================
//  Type PUserInfo3
// =============================================================================

  PUserInfo3 = ^TUserInfo3;
  {$EXTERNALSYM _USER_INFO_3}
  _USER_INFO_3 = record
    usri3_name: LPWSTR;
    usri3_password: LPWSTR;
    usri3_password_age: DWORD;
    usri3_priv: DWORD;
    usri3_home_dir: LPWSTR;
    usri3_comment: LPWSTR;
    usri3_flags: DWORD;
    usri3_script_path: LPWSTR;
    usri3_auth_flags: DWORD;
    usri3_full_name: LPWSTR;
    usri3_usr_comment: LPWSTR;
    usri3_parms: LPWSTR;
    usri3_workstations: LPWSTR;
    usri3_last_logon: DWORD;
    usri3_last_logoff: DWORD;
    usri3_acct_expires: DWORD;
    usri3_max_storage: DWORD;
    usri3_units_per_week: DWORD;
    usri3_logon_hours: PBYTE;
    usri3_bad_pw_count: DWORD;
    usri3_num_logons: DWORD;
    usri3_logon_server: LPWSTR;
    usri3_country_code: DWORD;
    usri3_code_page: DWORD;
    usri3_user_id: DWORD;
    usri3_primary_group_id: DWORD;
    usri3_profile: LPWSTR;
    usri3_home_dir_drive: LPWSTR;
    usri3_password_expired: DWORD;
  end;

  TUserInfo3 = _USER_INFO_3;
  {$EXTERNALSYM USER_INFO_3}
  USER_INFO_3 = _USER_INFO_3;

  PUserInfo1053 = ^TUserInfo1053;
  {$EXTERNALSYM _USER_INFO_1053}
  _USER_INFO_1053 = record
    usri1053_home_dir_drive : LPWSTR;
  end;
  TUserInfo1053 = _USER_INFO_1053;
  {$EXTERNALSYM USER_INFO_1053}
  USER_INFO_1053 = _USER_INFO_1053;

  PUserInfo1006 = ^TUserInfo1006;
  {$EXTERNALSYM _USER_INFO_1006}
  _USER_INFO_1006 = record
    usri1006_home_dir : LPWSTR;
  end;
  TUserInfo1006 = _USER_INFO_1006;
  {$EXTERNALSYM USER_INFO_1006}
  USER_INFO_1006 = _USER_INFO_1006;

// =============================================================================
//  Type PGroupInfo0
// =============================================================================

  PGroupInfo0 = ^TGroupInfo0;
  {$EXTERNALSYM _GROUP_INFO_0}
  _GROUP_INFO_0 = record
    grpi0_name: LPWSTR;
  end;
  TGroupInfo0 = _GROUP_INFO_0;
  {$EXTERNALSYM GROUP_INFO_0}
  GROUP_INFO_0 = _GROUP_INFO_0;

// =============================================================================
//  Type PGroupInfo2
// =============================================================================

  PGroupInfo2 = ^TGroupInfo2;
  {$EXTERNALSYM _GROUP_INFO_2}
  _GROUP_INFO_2 = record
    grpi2_name: LPWSTR;
    grpi2_comment: LPWSTR;
    grpi2_group_id: DWORD;
    grpi2_attributes: DWORD;
  end;
  TGroupInfo2 = _GROUP_INFO_2;
  {$EXTERNALSYM GROUP_INFO_2}
  GROUP_INFO_2 = _GROUP_INFO_2;

// =============================================================================
//  Type PLocalGroupInfo1
// =============================================================================

  PLocalGroupInfo1 = ^TLocalGroupInfo1;
  {$EXTERNALSYM _LOCALGROUP_INFO_1}
  _LOCALGROUP_INFO_1 = record
    lgrpi1_name: LPWSTR;
    lgrpi1_comment: LPWSTR;
  end;
  TLocalGroupInfo1 = _LOCALGROUP_INFO_1;
  {$EXTERNALSYM LOCALGROUP_INFO_1}
  LOCALGROUP_INFO_1 = _LOCALGROUP_INFO_1;

 // =============================================================================
//  Type PLocalGroupMembersinfo3
// =============================================================================

  PLocalGroupMembersinfo3 = ^TLocalGroupMembersinfo3;
  {$EXTERNALSYM _LOCALGROUP_MEMBERS_INFO_3}
  _LOCALGROUP_MEMBERS_INFO_3 = record
    lgrpi3_domainandname: LPWSTR;
  end;
  TLocalGroupMembersinfo3 = _LOCALGROUP_MEMBERS_INFO_3;
  {$EXTERNALSYM LOCALGROUP_MEMBERS_INFO_3}
  LOCALGROUP_MEMBERS_INFO_3 = _LOCALGROUP_MEMBERS_INFO_3;

// =============================================================================
//  Type PUserInfo0
// =============================================================================

  PUserInfo0 = ^TUserInfo0;
  {$EXTERNALSYM _USER_INFO_0}
  _USER_INFO_0  = record
    usri0_name: LPWSTR;
  end;
  TUserInfo0 = _USER_INFO_0;
  {$EXTERNALSYM USER_INFO_0}
  USER_INFO_0 = _USER_INFO_0;

// =============================================================================
//  Type PServerInfo101
// =============================================================================

  PServerInfo101 = ^TServerInfo101;
  _SERVER_INFO_101 = record
    sv101_platform_id: DWORD;
    sv101_name: LPWSTR;
    sv101_version_major: DWORD;
    sv101_version_minor: DWORD;
    sv101_type: DWORD;
    sv101_comment: LPWSTR;
  end;
  {$EXTERNALSYM _SERVER_INFO_101}
  TServerInfo101 = _SERVER_INFO_101;
  SERVER_INFO_101 = _SERVER_INFO_101;
  {$EXTERNALSYM SERVER_INFO_101}

 _SHARE_INFO_2 = record
    shi2_netname: LPWSTR;
    shi2_type: DWORD;
    shi2_remark: LPWSTR;
    shi2_permissions: DWORD;
    shi2_max_uses: DWORD;
    shi2_current_uses: DWORD;
    shi2_path: LPWSTR;
    shi2_passwd: LPWSTR;
  end;

  PSHARE_INFO_2= ^_SHARE_INFO_2;
  _SHARE_INFO_502         =  packed record
     shi502_netname:      PWideChar;
     shi502_type:         DWORD;
     shi502_remark:       PWideChar;
     shi502_permissions:  DWORD;
     shi502_max_uses:     DWORD;
     shi502_current_uses: DWORD;
     shi502_path:         LPWSTR;
     shi502_passwd:       LPWSTR;
     shi502_reserved:     DWORD;
     shi502_security_dsc: PSECURITY_DESCRIPTOR;
  end;
  SHARE_INFO_502          =  _SHARE_INFO_502;
  PSHARE_INFO_502         =  ^SHARE_INFO_502;
  LPSHARE_INFO_502        =  PSHARE_INFO_502;
  TShareInfo502           =  SHARE_INFO_502;
  PShareInfo502           =  PSHARE_INFO_502;

  TShareInfo502Array      =  Array [0..MaxWord] of TShareInfo502;
  PShareInfo502Array      =  ^TShareInfo502Array;

var
  HandleAPI : THandle;
  erreur    : integer;
  Sshare : PSHARE_INFO_2 ;
// =============================================================================
//  Liste des fonctions externes
// =============================================================================

  _NetServerEnum   : function (servername: LPCWSTR;
                               level: DWORD;
                               var bufptr: Pointer;
                               prefmaxlen: DWORD;
                               var entriesread: DWORD;
                               var totalentries: DWORD;
                               servertype: DWORD;
                               domain: LPCWSTR;
                               resume_handle: PDWORD): longword; stdcall;

  _NetUserGetInfo  : function (servername: LPCWSTR;
                               username: LPCWSTR;
                               level: DWORD;
                               var bufptr: Pointer): longword; stdcall;

  _NetUserAdd      : function (servername: LPCWSTR;
                               level: DWORD;
                               bufptr: Pointer;
                               err: PDword): longword; stdcall;

  _NetUserSetInfo  : function (servername: LPCWSTR;
                               username: LPCWSTR;
                               level: DWORD;
                               bufptr: Pointer;
                               err_param: PDWORD):longword;stdcall;

  _NetUserEnum    : function (servername: LPCWSTR;
                              level: DWORD;
                              filter: DWORD;
                              var bufptr: Pointer;
                              prefmaxlen: DWORD;
                              var entriesread: DWORD;
                              var totalentries: DWORD;
                              resume_handle: PDWORD): longword; stdcall;

  _NetGroupEnum   : function (servername: LPCWSTR;
                              level: DWORD;
                              var bufptr: Pointer;
                              prefmaxlen: DWORD;
                              var entriesread: DWORD;
                              var totalentries: DWORD;
                              resume_handle: PDWORD): longword; stdcall;

  _NetLocalGroupEnum : function (servername: LPCWSTR;
                                 level: DWORD;
                                 var bufptr: Pointer;
                                 prefmaxlen: DWORD;
                                 var entriesread: DWORD;
                                 var totalentries: DWORD;
                                 resumehandle: PDWORD): longword; stdcall;

  _NetApiBufferFree: function (Buffer: Pointer): longword; stdcall;

  _NetGroupGetUsers: function (servername: LPCWSTR;
                               groupname: LPCWSTR;
                               level: DWORD;
                               var bufptr: Pointer;
                               prefmaxlen: DWORD;
                               var entriesread: DWORD;
                               var totalentries: DWORD;
                               ResumeHandle: PDWORD): longword; stdcall;

  _NetGetJoinInformation: function (lpServer: LPCWSTR;
                                    var lpNameBuffer: LPWSTR;
                                    var BufferType: DWORD): longword; stdcall;

  _NetGetAnyDCName: function (servername: LPCWSTR;
                              domainname: LPCWSTR;
                              var bufptr: Pointer): longword; stdcall;

  _NetGetDCName: function (servername: LPCWSTR;
                           domainname: LPCWSTR;
                           var bufptr: Pointer): longword; stdcall;

  _NetServerDiskEnum: function (servername: LPWSTR;
                                level: DWORD;
                                var bufptr: Pointer;
                                prefmaxlen: DWORD;
                                var entriesread: DWORD;
                                var totalentries: DWORD;
                                resume_handle: PDWORD): LongWord ; stdcall;

  _NetUserGetGroups: function (servername: LPCWSTR;
                               username: LPCWSTR;
                               level: DWORD;
                               var bufptr: Pointer;
                               prefmaxlen: DWORD;
                               var entriesread: DWORD;
                               var totalentries: DWORD): longword; stdcall;

  _NetGroupAddUser: function (servername: LPCWSTR;
                              GroupName: LPCWSTR;
                              username: LPCWSTR): longword; stdcall;

  _NetLocalGroupAddMembers : function (servername: LPCWSTR;
                                       GroupName: LPCWSTR;
                                       level : Dword;
                                       Buffer : pointer;
                                       totalentries : Dword): longword; stdcall;

  _NetGroupDelUser: function (servername: LPCWSTR;
                              GroupName: LPCWSTR;
                              Username: LPCWSTR): longword; stdcall;

  _NetShareEnum : function   (servername: PWideChar;
                              level: DWORD;
                              bufptr: PByteArray;
                              prefmaxlen: DWORD;
                              entriesread: PDWORD;
                              totalentries: PDWORD;
                              resume_handle: PDWORD): DWORD; stdcall;

  function IcmpCreateFile : THandle; stdcall; external 'icmp.dll';
  function IcmpCloseHandle  (icmpHandle : THandle) : boolean; stdcall; external 'icmp.dll';
  function IcmpSendEcho  (IcmpHandle : THandle; DestinationAddress : IPAddr;
                          RequestData : Pointer; RequestSize : Smallint;
                          RequestOptions : pointer;
                          ReplyBuffer : Pointer;
                          ReplySize : DWORD;
                          Timeout : DWORD) : DWORD; stdcall; external 'icmp.dll';
  function NetShareAdd(servername: LPWSTR;level: DWORD;buf: pchar;parm_err: LPDWORD): LongWord; stdcall; external 'netapi32.dll';
  function NetShareDel(servername: LPWSTR;netname : LPWSTR;reserved : DWORD):LongWord; stdcall; external 'netapi32.dll';
  function NetShareGetInfo(servername: LPWSTR;netname : LPWSTR;level: DWORD;buf: pSHARE_INFO_2) : longword; stdcall;external 'netapi32.dll';

const
  SID_REVISION  = 1;

  FILENAME_ADVAPI32     = 'ADVAPI32.DLL';

  PROC_CONVERTSIDTOSTRINGSIDA   = 'ConvertSidToStringSidA';

function WinGetSidStr (Sid : PSid) : string;
var
  SidToStr      : TConvertSidToStringSidA;
  h             : LongWord;
  Buf           : array [0..MAX_PATH - 1] of char;
  p             : PAnsiChar;
begin
  h := LoadLibrary (FILENAME_ADVAPI32);
  if h <> 0 then
  try
    @SidToStr := GetProcAddress (h, PROC_CONVERTSIDTOSTRINGSIDA);
    if @SidToStr <> nil then
    begin
      FillChar (Buf, SizeOf(Buf), 0);

      if SidToStr (Sid, p) then
        Result := '{' + string(p) + '}';

      LocalFree (LongWord(p));
    end;
  finally
    FreeLibrary (h);
  end;
end;

function GetSidStr (Sid : PSid) : string;
var
  Psia          : PSIDIdentifierAuthority;
  SubAuthCount  : LongWord;
  i             : LongWord;
begin
  if IsValidSid (Sid) then
  begin
    //
    // Win 2000+ contains ConvertSidToStringSidA() in advapi32.dll so we just
    // use it if we can
    //
    if (Win32Platform = VER_PLATFORM_WIN32_NT) and
      (Win32MajorVersion >= 5) then
      Result := WinGetSidStr (Sid)
    else
    begin
      Psia := GetSidIdentifierAuthority (Sid);
      SubAuthCount := GetSidSubAuthorityCount (Sid)^;
      Result := Format('[S-%u-', [SID_REVISION]);
      if ((Psia.Value[0] <> 0) or (Psia.Value[1] <> 0)) then
        Result := Result + Format ('%.2x%.2x%.2x%.2x%.2x%.2x', [Psia.Value[0],
          Psia.Value[1], Psia.Value[2], Psia.Value[3], Psia.Value[4],
          Psia.Value[5]])
      else
        Result := Result + Format ('%u', [LongWord(Psia.Value[5]) +
          LongWord(Psia.Value[4] shl 8) + LongWord(Psia.Value[3] shl 16) +
          LongWord(Psia.Value[2] shl 24)]);

      for i := 0 to SubAuthCount - 1 do
        Result := Result + Format ('-%u', [GetSidSubAuthority(Sid, i)^]);

      Result := Result + ']';
    end;
  end;
end;

function sidtoName(sid : psid):tnomGroupe ;
var
  UserSize, DomainSize: DWORD;
  snu: SID_NAME_USE;
  User, Domain: string;
begin
  result.strnom  := '';
  result.strdomaine :='';
  LookupAccountSid(nil, sid, nil, UserSize, nil, DomainSize, snu);
  if (UserSize <> 0) and (DomainSize <> 0) then
  begin
    setLength(User, UserSize);
    SetLength(Domain, DomainSize);
    if LookupAccountSid(nil, Sid, PChar(User), UserSize,
    PChar(Domain), DomainSize, snu) then
    begin
      User := StrPas(PChar(User));
      Domain := StrPas(PChar(Domain));
      result.strdomaine  := domain;
      result.strnom  := user;
    end;
  end else
  result.strnom := getsidstr(sid);
end;

function APINetcharger(var addresseProcedure: Pointer; const strNomProcedure: string): Boolean;

//  ___________________________________________________________________________
// | Function  APINetcharger                                                   |
// |___________________________________________________________________________|
// | Permet de charger en memoire la dll et de recuperrer l'adresse de la      |
// | procedure voulue dans cette dll                                           |
// |___________________________________________________________________________|
// | Entrée | const strNomProcedure: string                                    |
// |        |   Nom de la procedure que l'on veut                              |
// |________|__________________________________________________________________|
// | Sortie | var addresseProcedure: Pointer                                   |
// |        |   pointeur sur la procedure                                      |
// |        | Result : boolean                                                 |
// |        |   true si ok si non false                                        |
// |________|__________________________________________________________________|

begin
  Result := Assigned(addresseProcedure);
  if not Result then
  begin
    if HandleAPI = 0 then
      HandleAPI := LoadLibrary('netapi32.dll');
    if HandleAPI <> 0 then
    begin
      addresseProcedure := GetProcAddress(HandleAPI, PChar(strNomProcedure));
      Result := Assigned(addresseProcedure);
    end;
  end;
end;

procedure ApiNetDecharger;

//  ___________________________________________________________________________
// | Function  ApiNetDecharger                                                 |
// |___________________________________________________________________________|
// | Permet de décharger en memoire la dll                                     |
// |___________________________________________________________________________|
// | Entrée |                                                                  |
// |        |                                                                  |
// |________|__________________________________________________________________|
// | Sortie |                                                                  |
// |        |                                                                  |
// |________|__________________________________________________________________|

begin
  if HandleAPI <> 0 then
  begin
    FreeLibrary(HandleAPI);
    HandleAPI := 0;
  end;
end;

function NetApiBufferFree(Buffer: Pointer): longword;

//  ___________________________________________________________________________
// | Function  ApiNetDecharger                                                 |
// |___________________________________________________________________________|
// | Permet de liberrer les ressources prises par un buffer                    |
// |___________________________________________________________________________|
// | Entrée | Buffer: Pointer                                                  |
// |        |   pointer sur le buffer à liberrer                               |
// |________|__________________________________________________________________|
// | Sortie | result : longword                                                |
// |        |   0 si ok                                                        |
// |________|__________________________________________________________________|

begin
  if APINetcharger(@_NetApiBufferFree, 'NetApiBufferFree') then
    Result := _NetApiBufferFree(Buffer)
  else
    Result := ERROR_CALL_NOT_IMPLEMENTED;
end;

procedure mefTableau(var tab : TTabUtilisateur0 ); overload;

//  ___________________________________________________________________________
// | procedure mefTableau                                                      |
// |___________________________________________________________________________|
// | Permet de mettre la colonne de trie d'un tableau en minuscule             |
// | pour pouvoir le trier par la suite                                        |
// |___________________________________________________________________________|
// | Entrée | tab : TTabUtilisateur0                                           |
// |        |   tableau de type TTabUtilisateur0 a mettre en forme             |
// |________|__________________________________________________________________|
// | Sortie | tab : TTabUtilisateur0                                           |
// |        |   tableau de type TTabUtilisateur0 mis en forme                  |
// |________|__________________________________________________________________|

var
  i : integer;
  nbuser : integer;
begin
  nbuser := length(tab);
  for i := 0 to nbuser - 1 do tab[i].login := LowerCase(tab[i].login );
end;

procedure mefTableau(var tab : TTabGroupe0 ); overload;

//  ___________________________________________________________________________
// | procedure mefTableau                                                      |
// |___________________________________________________________________________|
// | Permet de mettre la colonne de trie d'un tableau en minuscule             |
// | pour pouvoir le trier par la suite                                        |
// |___________________________________________________________________________|
// | Entrée | tab : TTabGroupe0                                                |
// |        |   tableau de type TTabGroupe0 a mettre en forme                  |
// |________|__________________________________________________________________|
// | Sortie | tab : TTabGroupe0                                                |
// |        |   tableau de type TTabGroupe0 mis en forme                       |
// |________|__________________________________________________________________|

var
  i : integer;
  nbuser : integer;
begin
  nbuser := length(tab);
  for i := 0 to nbuser - 1 do tab[i].nom := LowerCase(tab[i].nom );
end;

procedure mefTableau(var tab : TTabGroupeGlobaux  ); overload;

//  ___________________________________________________________________________
// | procedure mefTableau                                                      |
// |___________________________________________________________________________|
// | Permet de mettre la colonne de trie d'un tableau en minuscule             |
// | pour pouvoir le trier par la suite                                        |
// |___________________________________________________________________________|
// | Entrée | tab : TTabGroupeGlobaux                                          |
// |        |   tableau de type TTabGroupeGlobaux a mettre en forme            |
// |________|__________________________________________________________________|
// | Sortie | tab : TTabGroupeGlobaux                                          |
// |        |   tableau de type TTabGroupeGlobaux mis en forme                 |
// |________|__________________________________________________________________|

var
  i : integer;
  nbuser : integer;
begin
  nbuser := length(tab);
  for i := 0 to nbuser - 1 do tab[i].nom := LowerCase(tab[i].nom );
end;

procedure TrieTableau ( tab : TTabUtilisateur0 ); overload;

//  ___________________________________________________________________________
// | procedure TrieTableau                                                     |
// |___________________________________________________________________________|
// | Permet trier un tableau                                                   |
// |___________________________________________________________________________|
// | Entrée | tab : TTabUtilisateur0                                           |
// |        |   tableau de type TTabUtilisateur0 a trier                       |
// |________|__________________________________________________________________|
// | Sortie | tab : TTabUtilisateur0                                           |
// |        |   tableau de type TTabUtilisateur0 trié                          |
// |________|__________________________________________________________________|

var
  i , k: integer;
  nbuser : integer;
  temp : string;
  bootrie : boolean;
begin
  nbuser := length(tab);
  k := 0;
  mefTableau(tab);
  bootrie := false;
  while not bootrie do
  begin
    bootrie := true;
    for i := k to nbuser - 2 do
      if tab[i].login > tab[i+1].login then
      begin
        bootrie := false;
        temp := tab[i+1].login;
        tab[i+1].login := tab[i].login;
        tab[i].login := temp;
      end;
  end;
end;

procedure TrieTableau ( tab : TTabGroupe0 );overload;

//  ___________________________________________________________________________
// | procedure TrieTableau                                                     |
// |___________________________________________________________________________|
// | Permet trier un tableau                                                   |
// |___________________________________________________________________________|
// | Entrée | tab : TTabGroupe0                                                |
// |        |   tableau de type TTabGroupe0 a trier                            |
// |________|__________________________________________________________________|
// | Sortie | tab : TTabGroupe0                                                |
// |        |   tableau de type TTabGroupe0 trié                               |
// |________|__________________________________________________________________|

var
  i , k: integer;
  nbuser : integer;
  temp : Tgroupe0 ;
  bootrie : boolean;
begin
  nbuser := length(tab);
  k := 0;
  mefTableau(tab);
  bootrie := false;
  while not bootrie do
  begin
    bootrie := true;
    for i := k to nbuser - 2 do
      if tab[i].nom > tab[i+1].nom then
      begin
        bootrie := false;
        temp := tab[i+1];
        tab[i+1]:= tab[i];
        tab[i] := temp;
      end;
  end;
end;

procedure TrieTableau ( tab : TTabGroupeGlobaux );overload;

//  ___________________________________________________________________________
// | procedure TrieTableau                                                     |
// |___________________________________________________________________________|
// | Permet trier un tableau                                                   |
// |___________________________________________________________________________|
// | Entrée | tab : TTabGroupeGlobaux                                          |
// |        |   tableau de type TTabGroupeGlobaux a trier                      |
// |________|__________________________________________________________________|
// | Sortie | tab : TTabGroupeGlobaux                                          |
// |        |   tableau de type TTabGroupeGlobaux trié                         |
// |________|__________________________________________________________________|

var
  i , k: integer;
  nbuser : integer;
  temp : TGroupeGlobaux ;
  bootrie : boolean;
begin
  nbuser := length(tab);
  k := 0;
  mefTableau(tab);
  bootrie := false;
  while not bootrie do
  begin
    bootrie := true;
    for i := k to nbuser - 2 do
      if tab[i].nom > tab[i+1].nom then
      begin
        bootrie := false;
        temp := tab[i+1];
        tab[i+1]:= tab[i];
        tab[i] := temp;
      end;
  end;
end;


// =============================================================================
//  Liste des fonctions finalles
// =============================================================================

Function  CreateUser (const strNomServeur : string  ; Information : Tutilisateur ) : integer;
Var
  err : longword;
  errt : cardinal;
  wNomUtilisateur : WideString;
  wNomServeur     : WideString;
  Details         : USER_INFO_1  ;
begin
  wNomServeur     := strNomServeur;
  with Details do
  begin
    usri1_name := pwidechar (WideString(information.login));
    usri1_password := pwidechar (WideString(information.mot_de_passe));
    usri1_priv := Information.privilege ;
    if Information.dossier_de_base <>'' then
      usri1_home_dir := pwidechar (WideString(Information.dossier_de_base))
    else
      usri1_home_dir := nil;
    if Information.description <>'' then
      usri1_comment := pwidechar (WideString(Information.description))
    else
      usri1_comment := nil;
    usri1_flags := Information.flags;
    if Information.dossier_script <>'' then
      usri1_script_path := pwidechar (WideString(Information.dossier_script))
    else
      usri1_script_path := nil;
  end;
  if APINetcharger (@_NetUseradd, 'NetUserAdd') then
    err := _NetUseradd(pwidechar(wNomServeur ),
                       1,
                       @details,
                       @errt );
//  NetApiBufferFree(Details);
  result := err;
  erreur:=err;
end;

Function  GetinfoUtilisateur    (const strNomServeur : string ; const strNomUtilisateur : string ; var Information : Tutilisateur ) : integer;
//  ___________________________________________________________________________
// | Function  GetinfoUtilisateur                                              |
// |___________________________________________________________________________|
// |  Permet de récuperrer les information sur un utilisateur donnée           |
// | Dans une structure Tutilisateur                                           |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur a partir du quel recuperrer les info utilisateur|
// |        |   vide Machine local si non '\\nommachine'                       |
// |        | const strNomUtilisateur : string                                 |
// |        |   Nom de l'utilisateur pour qui recuperrer les informations      |
// |________|__________________________________________________________________|
// | Sortie | var Information : Tutilisateur                                   |
// |        |   structure qui va recevoir les informations de l'utilisateurs   |
// |        | Result : integer                                                 |
// |        |   -1 en cas d'erreur                                             |
// |________|__________________________________________________________________|


var
  buffer          : pointer;
  wNomUtilisateur : WideString;
  wNomServeur     : WideString;
  Details         : PUserInfo3 ;
  Err             : longword;

begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  result := -1;
  err    := $FFFF ;
  wNomUtilisateur := strNomUtilisateur;
  wNomServeur     := strNomServeur;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger (@_NetUserGetInfo, 'NetUserGetInfo') then
  err := _NetUserGetInfo( pwidechar(wNomServeur ),
                         pwidechar(wNomUtilisateur ),
                         3 ,                            // utilise la tructure 3 pour les infos utilisateurs
                         buffer );
  if Err = errsuccess  then
  begin
    details := puserinfo3(buffer);
    with Information do
    begin
      login                   := details^.usri3_name;
      mot_de_passe            := details^.usri3_password;
      age_mot_de_passe        := details^.usri3_password_age;
      privilege               := details^.usri3_priv;
      dossier_de_base         := details^.usri3_home_dir;
      description             := details^.usri3_comment;
      flags                   := details^.usri3_flags;
      dossier_script          := details^.usri3_script_path;
      flags_privilege         := details^.usri3_auth_flags;
      nom_complet             := details^.usri3_full_name;
      commentaire_utilisateur := details^.usri3_usr_comment;
      parametre               := details^.usri3_parms;
      ordinateurs             := details^.usri3_workstations;
// ****************************************************************************
// * Les dates sont renvoyé en nombre de seconde depuis 01/01/1970            *
// ****************************************************************************
      date_connexion          := details^.usri3_last_logon / 86400 + 25569;
      date_deconnexion        := details^.usri3_last_logoff / 86400 + 25569;
      date_expiration         := details^.usri3_acct_expires / 86400 + 25569;
      stockage_maxi           := details^.usri3_max_storage;
      units_per_week          := details^.usri3_units_per_week;
      logon_hours             := details^.usri3_logon_hours;
      nb_tentative_login      := details^.usri3_bad_pw_count;
      nb_login                := details^.usri3_num_logons;
      serveur_de_login        := details^.usri3_logon_server;
      code_region             := details^.usri3_country_code;
      code_page               := details^.usri3_code_page;
      id_utilisateur          := details^.usri3_user_id;
      id_groupe_principal     := details^.usri3_primary_group_id;
      profile                 := details^.usri3_profile;
      lettre_dossier_de_base  := details^.usri3_home_dir_drive;
      mot_passe_expire        := details^.usri3_password_expired <> 0 ;
    end;
    result := 0;
  end;
  NetApiBufferFree(buffer);
end;

Function SetUserHomeFolder (const strNomServeur : string ; const strNomUtilisateur : string ; const strLecteur : string;const strPartage : string ):integer;
var
  buffer          : pointer;
  wNomUtilisateur : WideString;
  wNomServeur     : WideString;
  Details         : TUserInfo1006 ;
  Err             : longword;

begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  result := -1;
  err    := $FFFF ;
  wNomUtilisateur := strNomUtilisateur;
  wNomServeur     := strNomServeur;
  buffer := nil;
  begin
    if APINetcharger (@_NetUserSetInfo, 'NetUserSetInfo') then
    begin
      GetMem ( buffer,sizeof(TUserInfo1053));
      puserinfo1053(buffer)^.usri1053_home_dir_drive := PWideChar(WideString(strLecteur));
      err := _NetUserSetInfo( pwidechar(wNomServeur ),
                              pwidechar(wNomUtilisateur ),
                              1053 ,
                              buffer,
                              nil );
      freemem(buffer);
      GetMem ( buffer,sizeof(TUserInfo1006));
      puserinfo1006(buffer)^.usri1006_home_dir := PWideChar(WideString(strPartage));

      err := _NetUserSetInfo( pwidechar(wNomServeur ),
                              pwidechar(wNomUtilisateur ),
                              1006 ,
                              buffer,
                              nil );
      freemem(buffer);
    end;
    if err = errSuccess then result := 0;
  end;
end;

Function  GetListeComptes  (const strNomServeur : string; fltType : longword ; var tabUtilisateur : TTabUtilisateur ;posi : tprogressbar =nil) : integer;
//  ___________________________________________________________________________
// | Function  GetListeComptes                                                 |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des comptes d'une machine donnée           |
// | Dans un tableau de type Tutilisateur                                      |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur sur le quel lister les utilisateurs vide pour   |
// |        |   Machine local si non '\\nommachine'                            |
// |        | fltType : longword ;                                             |
// |        |   filtre sur le type de compte à chercher                        |
// |________|__________________________________________________________________|
// | Sortie | var tabUtilisateur : TTabUtilisateur                             |
// |        |   tableau qui va recevoir la liste des utilisateurs              |
// |        | Result : integer                                                 |
// |        |   Nombre d'utilisateur trouvés -1 en cas d'erreur                |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  Details     : PUserInfo3 ;
  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;
  pc , pco : integer;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// * Vidage du tableau                                                        *
// ****************************************************************************

  result := -1;
  err := $FFFF;
  setlength(tabUtilisateur , 0 );
  wNomServeur := strNomServeur;
  pco := 0;
  if posi <> nil then
  begin
    posi.Max := 100;
    posi.Position := 0;
    Application.ProcessMessages ;
  end;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger (@_NetUserEnum, 'NetUserEnum') then
  err := _NetUserEnum(PWideChar(wNomServeur),
                     3 ,                          // utilise la tructure 3 pour les infos utilisateurs
                     fltType ,                    // filtre de recherche
                     buffer,
                     LongueurMaxi ,
                     EntriesRead,
                     TotalEntries,
                     nil);

// ****************************************************************************
// * Recherche Ok on transfere dans le tableau                                *
// ****************************************************************************

  if Err = errSuccess then
  begin
    details := puserinfo3(buffer);
    setlength(tabUtilisateur , EntriesRead );
    for I := 0 to EntriesRead - 1 do
    with tabutilisateur[i] do
    begin
      login                   := details^.usri3_name;
      mot_de_passe            := details^.usri3_password;
      age_mot_de_passe        := details^.usri3_password_age;
      privilege               := details^.usri3_priv;
      dossier_de_base         := details^.usri3_home_dir;
      description             := details^.usri3_comment;
      flags                   := details^.usri3_flags;
      dossier_script          := details^.usri3_script_path;
      flags_privilege         := details^.usri3_auth_flags;
      nom_complet             := details^.usri3_full_name;
      commentaire_utilisateur := details^.usri3_usr_comment;
      parametre               := details^.usri3_parms;
      ordinateurs             := details^.usri3_workstations;
// ****************************************************************************
// * Les dates sont renvoyé en nombre de seconde depuis 01/01/1970            *
// ****************************************************************************
      date_connexion          := details^.usri3_last_logon / 86400 + 25569;
      date_deconnexion        := details^.usri3_last_logoff / 86400 + 25569;
      date_expiration         := details^.usri3_acct_expires / 86400 + 25569;
      stockage_maxi           := details^.usri3_max_storage;
      units_per_week          := details^.usri3_units_per_week;
      logon_hours             := details^.usri3_logon_hours;
      nb_tentative_login      := details^.usri3_bad_pw_count;
      nb_login                := details^.usri3_num_logons;
      serveur_de_login        := details^.usri3_logon_server;
      code_region             := details^.usri3_country_code;
      code_page               := details^.usri3_code_page;
      id_utilisateur          := details^.usri3_user_id;
      id_groupe_principal     := details^.usri3_primary_group_id;
      profile                 := details^.usri3_profile;
      lettre_dossier_de_base  := details^.usri3_home_dir_drive;
      mot_passe_expire        := details^.usri3_password_expired <> 0 ;
      Inc(Details);
      if posi <> nil then
      begin
        pc := (i-1)*100 div EntriesRead;
        if pc <> pco then
        begin
          posi.Position := pc;
          pco := pc;
          Application.ProcessMessages ;
        end
      end;
    end;
    result := EntriesRead
  end;
  NetApiBufferFree(buffer);
end;

function  GetListeGroupeUtilisateur ( const strNomServeur : string; const strNomUtilisateur : string ; var tabGroupe : TTabGroupe0) :integer;

//  ___________________________________________________________________________
// | Function  GetListeGroupeUtilisateur                                       |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des groupes dont un utilisateur est membre |
// | Dans un tableau de type TTabGroupe0                                       |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur sur le quel faire la recherche                  |
// |        | const strNomUtilisateur : string                                 |
// |        |   Nom de l'utilisateur                                           |
// |________|__________________________________________________________________|
// | Sortie | var tabGroupe : TTabGroupe0                                      |
// |        |   tableau qui va recevoir la liste des groupes                   |
// |        | Result : integer                                                 |
// |        |   Nombre de groupe trouvés -1 en cas d'erreur                    |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  wNomUtilisateur : WideString ;
  Details     : PGroupInfo0  ;
  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// * Vidage du tableau                                                        *
// ****************************************************************************

  result := -1;
  err := $FFFF;
  setlength(tabGroupe , 0 );
  wNomServeur := strNomServeur;
  wNomUtilisateur := strNomUtilisateur ;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger (@_NetUserGetGroups, 'NetUserGetGroups') then
  err := _NetUserGetGroups(PWideChar(wNomServeur),
                     pwidechar(wNomUtilisateur ),
                     0 ,                          // utilise la tructure 3 pour les infos utilisateurs
                     buffer,
                     LongueurMaxi ,
                     EntriesRead,
                     TotalEntries);

// ****************************************************************************
// * Recherche Ok on transfere dans le tableau                                *
// ****************************************************************************

  if Err = errSuccess then
  begin
    details := PGroupInfo0(buffer);
    setlength(tabGroupe , EntriesRead );
    for I := 0 to EntriesRead - 1 do
    with tabGroupe[i] do
    begin
      nom := details^.grpi0_name ;
      Inc(Details);
    end;
    TrieTableau(tabGroupe );
    result := EntriesRead ;
  end;
  NetApiBufferFree(buffer);
end;

Function  GetListeGroupeGlobaux(const strNomServeur : string ; var tabGroupeGlobaux : TTabGroupeGlobaux ) : integer;

//  ___________________________________________________________________________
// | Function  GetListeGroupeGlobaux                                           |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des Groupes globaux d'un serveur donné     |
// | Dans un tableau de type TGroupeGlobaux                                    |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur sur le quel lister les Groupes vide pour        |
// |        |   Machine local si non '\\nommachine'                            |
// |________|__________________________________________________________________|
// | Sortie | var tabGroupeGlobaux : TTabGroupeGlobaux                         |
// |        |   tableau qui va recevoir la liste des groupes                   |
// |        | Result : integer                                                 |
// |        |   Nombre de groupe trouvés -1 en cas d'erreur                    |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  Details     : PGroupInfo2  ;
  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// * Vidage du tableau                                                        *
// ****************************************************************************

  result := -1;
  setlength(tabGroupeGlobaux , 0 );
  wNomServeur := strNomServeur;
  err := $FFFF;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger (@_NetGroupEnum, 'NetGroupEnum') then
  Err := _NetGroupEnum(PWideChar(wNomServeur),
                      2,                        // utilise la tructure 2 pour les infos groupes globaux
                      Buffer,
                      LongueurMaxi ,
                      EntriesRead,
                      TotalEntries,
                      nil);


// ****************************************************************************
// * Recherche Ok on transfere dans le tableau                                *
// ****************************************************************************

  if Err = errSuccess then
  begin
    Details := PGroupInfo2(Buffer);

// ****************************************************************************
// * Evite la remonté d'un groupe inexistant                                  *
// ****************************************************************************

    if (EntriesRead <> 1) or (Details^.grpi2_name <> 'None') then
    begin
      setlength(tabGroupeGlobaux , EntriesRead );
      for I := 0 to EntriesRead - 1 do
      with tabGroupeGlobaux[i] do
      begin
        nom         := Details^.grpi2_name;
        commentaire := Details^.grpi2_comment ;
        id_group    := Details^.grpi2_group_id;
        attributs   := Details^.grpi2_attributes;
        Inc(Details);
      end;
      Result := EntriesRead ;
      TrieTableau(tabGroupeGlobaux);
    end else
      Result := 0 ;
  end
  else
    RaiseLastOSError;
  NetApiBufferFree(Buffer);

end;

Function  GetListeGroupeLocaux  (const strNomServeur : string ; var tabGroupeLocaux : TTabGroupeLoCaux ) : integer;

//  ___________________________________________________________________________
// | Function  GetListeGroupeLocaux                                            |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des Groupes globaux d'un serveur donné     |
// | Dans un tableau de type TGroupeGlobaux                                    |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur sur le quel lister les Groupes vide pour        |
// |        |   Machine local si non '\\nommachine'                            |
// |________|__________________________________________________________________|
// | Sortie | var tabGroupeCocaux : TTabGroupeLoCaux                           |
// |        |   tableau qui va recevoir la liste des groupes                   |
// |        | Result : integer                                                 |
// |        |   Nombre de groupe trouvés -1 en cas d'erreur                    |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  Details     : PLocalGroupInfo1  ;
  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;

begin

// ****************************************************************************
// * Initialisation des variables                                             *
// * Vidage du tableau                                                        *
// ****************************************************************************

  wNomServeur := strNomServeur;
  result := -1;
  err := $FFFF;
  setlength(tabGroupeLocaux ,0);

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger (@_NetLocalGroupEnum, 'NetLocalGroupEnum') then
  Err := _NetLocalGroupEnum(PWideChar(wNomServeur),
                           1,
                           Buffer,
                           LongueurMaxi ,
                           EntriesRead,
                           TotalEntries,
                           nil);


// ****************************************************************************
// * Recherche Ok on transfere dans le tableau                                *
// ****************************************************************************

  if Err = errsuccess then
  begin
    Details := PLocalGroupInfo1(Buffer);
    setlength(tabGroupeLocaux ,EntriesRead );
    for I := 0 to EntriesRead - 1 do
    with tabGroupeLocaux[i] do
    begin
      nom         := Details^.lgrpi1_name ;
      commentaire := Details^.lgrpi1_comment ;
      Inc(Details);
    end;
    result := EntriesRead ;
  end;
  NetApiBufferFree(Buffer);
end;

Function  GetListeUtilisateursGroupeGlobal (const strNomServeur : string ;const strNomGroupe : string; var tabUtilisateurs : TTabUtilisateur0 ) : integer;

//  ___________________________________________________________________________
// | Function  GetListeUtilisateursGroupeGlobal                                |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des utilisateurs d'un Groupes global       |
// | Dans un tableau de type TTabUtilisateur0                                  |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur sur le quel lister les Groupes vide pour        |
// |        |   Machine local si non '\\nommachine'                            |
// |        | const strNomGroupe : string                                      |
// |        |   Nom du groupe global don on veux conaitre la liste des ut      |
// |________|__________________________________________________________________|
// | Sortie | var tabUtilisateurs : TTabUtilisateur0                           |
// |        |   tableau qui va recevoir la liste des utilisateurs              |
// |        | Result : integer                                                 |
// |        |   Nombre d'utilisateur trouvés -1 en cas d'erreur                |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  wNomGroupe  : WideString ;
  Details     : PUserInfo0 ;
  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;

begin

// ****************************************************************************
// * Initialisation des variables                                             *
// * Vidage du tableau                                                        *
// ****************************************************************************

  result := -1;
  setlength(tabUtilisateurs , 0 );
  wNomServeur := strNomServeur;
  wNomGroupe  := strNomGroupe;
  err := errmax;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger(@_NetGroupGetUsers, 'NetGroupGetUsers') then
  err := _NetGroupGetUsers(PWideChar(wNomServeur),
                     PWideChar(wNomGroupe),
                     0 ,                      // utilise la tructure 0 pour les infos utilisateurs
                     buffer,
                     LongueurMaxi ,
                     EntriesRead,
                     TotalEntries,
                     nil);

  if Err = errsuccess then
  begin
    Details := Puserinfo0(Buffer);
    setlength(tabUtilisateurs ,EntriesRead );
    for I := 0 to EntriesRead - 1 do
    with tabUtilisateurs[i] do
    begin
      login         := Details^.usri0_name ;
      Inc(Details);
    end;
    result := EntriesRead ;
  end;
  NetApiBufferFree(Buffer);

end;

Function  GetListeUtilisateursGroupeGlobal (const strNomServeur : string ;const strNomGroupe : string; var tabUtilisateurs : TTabUtilisateur;posi : tprogressbar = nil) : integer;

//  ___________________________________________________________________________
// | Function  GetListeUtilisateursGroupeGlobal                                |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des utilisateurs d'un Groupes global       |
// | Dans un tableau de type TTabUtilisateur                                   |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur : string                                     |
// |        |   Nom du serveur sur le quel lister les Groupes vide pour        |
// |        |   Machine local si non '\\nommachine'                            |
// |        | const strNomGroupe : string                                      |
// |        |   Nom du groupe global don on veux conaitre la liste des ut      |
// |        | posi : tprogressbar = nil                                        |
// |        |  pointeur sur une progressbar                                    |
// |________|__________________________________________________________________|
// | Sortie | var tabUtilisateurs : TTabUtilisateur                            |
// |        |   tableau qui va recevoir la liste des utilisateurs              |
// |        | Result : integer                                                 |
// |        |   Nombre d'utilisateur trouvés -1 en cas d'erreur                |
// |________|__________________________________________________________________|

var
  t1 : TTabUtilisateur0 ;
  nbuser : integer;
  i : integer;
  pc, opc : integer;
begin
  result := -1;
  nbuser := GetListeUtilisateursGroupeGlobal(strNomServeur,strNomGroupe,t1);
  if nbuser > 0 then
  begin
    TrieTableau( t1 );
    opc :=0;
    if posi <> nil then
    begin
      posi.Max := 100;
      posi.Position :=0;
    end;
    setlength(tabUtilisateurs , nbuser);
    for i := 0 to nbuser - 1 do
    begin
      GetinfoUtilisateur(strNomServeur,t1[i].login,tabutilisateurs[i] );
      if posi <> nil then
      begin
        pc := ((i+1)*100 div nbuser);
        if pc <> opc then
        begin
          posi.Position := pc;
          opc:=pc;
          Application.ProcessMessages ;
        end;
      end;
    end;
  end;
  result := nbuser;
end;

function  GetNomDomaine ( var nomDomaine : string ;var  EtatIntegration : longword) : integer;
//  ___________________________________________________________________________
// | Function  GetNomDomaine                                                   |
// |___________________________________________________________________________|
// |  Permet de récuperrer le nom du domaine de la machine local               |
// |___________________________________________________________________________|
// | Sortie | var nomDomaine : string                                          |
// |        |   nom du domaine de la machine local                             |
// |        | var  EtatIntegration : longword                                  |
// |        |   indique si la machine est intégrée au domaine ou pas           |
// |        |   0 indeterminé                                                  |
// |        |   1 l'odinateur n'est pas intégré dans un domaine                |
// |        |   2 l'ordinateur apartien un workgroup                           |
// |        |   3 l'ordinateur est intégré au domaine                          |
// |        | Result : integer                                                 |
// |        |   0 si pas d'erreur                                              |
// |________|__________________________________________________________________|

var
  Err         : longword;
  wnomdomaine : PWideChar;
  typedomaine : cardinal;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger(@_NetGetJoinInformation, 'NetGetJoinInformation') then
    Err := _NetGetJoinInformation (nil,wnomdomaine,typedomaine);
  if Err = errSuccess then
  begin
    Nomdomaine := wnomdomaine ;
    EtatIntegration := typedomaine ;
    result := 0;
  end;
  NetApiBufferFree(wnomdomaine );
end;

function  GetNomTousPdc ( strServeur : string;strDomaine : string ; var listepdc : string ):integer;
//  ___________________________________________________________________________
// | Function  GetNomTousPdc                                                   |
// |___________________________________________________________________________|
// |  Permet de récuperrer le nom d'un des controleurs du domaine              |
// |___________________________________________________________________________|
// | Entrée | strServeur : string                                              |
// |        |   nom de l'ordinateur pour qui faire la recherche                |
// |        | vstrDomaine : string                                             |
// |        |   Nom du domaine à chercher                                      |
// |________|__________________________________________________________________|
// | Sortie | var listepdc : string                                            |
// |        |   Nom du premier PDC trouvé                                      |
// |        | Result : integer                                                 |
// |        |   0 si pas d'erreur                                              |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  wNomDomaine : WideString;
  Err         : longword;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  wNomServeur := strServeur ;
  wNomDomaine := strDomaine;
  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger(@_NetGetAnyDCName, 'NetGetAnyDCName') then
  err := _NetGetAnyDCName (PWideChar( wNomServeur),
                          PWideChar( wNomDomaine),
                          buffer);
  if Err = errsuccess then
  begin
    listepdc := pwidechar(Buffer);
    result := 0;
  end;
  NetApiBufferFree(Buffer);
end;

function  GetNomPdc ( strServeur : string;strDomaine : string ; var listepdc : string ):integer;
//  ___________________________________________________________________________
// | Function  GetNomPdc                                                       |
// |___________________________________________________________________________|
// |  Permet de récuperrer le nom du des controleurs principal du domaine      |
// |___________________________________________________________________________|
// | Entrée | strServeur : string                                              |
// |        |   nom de l'ordinateur pour qui faire la recherche                |
// |        | vstrDomaine : string                                             |
// |        |   Nom du domaine à chercher                                      |
// |________|__________________________________________________________________|
// | Sortie | var listepdc : string                                            |
// |        |   Nom du PDC                                                     |
// |        | Result : integer                                                 |
// |        |   0 si pas d'erreur                                              |
// |________|__________________________________________________________________|

var
  buffer      : pointer;
  wNomServeur : WideString;
  wNomDomaine : WideString;
  Err         : longword;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  wNomServeur := strServeur ;
  wNomDomaine := strDomaine;
  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger(@_NetGetDCName, 'NetGetDCName') then
    err := _NetGetDCName (PWideChar( wNomServeur),
                       PWideChar( wNomDomaine),
                       buffer);
  if Err = errsuccess then
  begin
    listepdc := pwidechar(Buffer);
    result := 0;
  end;
  NetApiBufferFree(Buffer);
end;




function  GetListeServeur (const strNomdomaine :String;srvType : longword ;var tabServeur : TTabServeur ) : integer;
//  ___________________________________________________________________________
// | Function  GetListeServeur                                                 |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des serveurs en fonction du type           |
// |___________________________________________________________________________|
// | Entrée | const strNomdomaine :String                                      |
// |        |   nom du domaine sur le quel chercher les serveurs               |
// |        | srvType : integer                                                |
// |        |   Type de serveur à chercher                                     |
// |________|__________________________________________________________________|
// | Sortie | var tabServeur : TTabServeur                                     |
// |        |   Liste des serveur                                              |
// |        | Result : integer                                                 |
// |        |   -1 en cas d'erreur si non nombre de serveur trouvé             |
// |________|__________________________________________________________________|

var

  buffer      : pointer;
  wNomDomaine : WideString;
  Details     : PServerInfo101;
//  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;

begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  result := -1;
 // erreur    := $FFFF ;
  wNomDomaine := strNomdomaine ;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger (@_NetServerEnum, 'NetServerEnum') then
    erreur := _NetServerEnum(nil,
                             101,
                             buffer,
                             LongueurMaxi ,
                             entriesread,
                             totalentries,
                             srvType ,
                             pwidechar(wNomDomaine ),
                             nil);
  if erreur = errSuccess then
  begin
    SetLength (tabServeur , EntriesRead );
    details := pserverinfo101(buffer);
    for i := 0 to EntriesRead -1 do
    with tabServeur[i] do
    begin
      platformId   := details^.sv101_platform_id  ;
      nom          := details^.sv101_name  ;
      versionMajor := details^.sv101_version_major;
      versionMinor := details^.sv101_version_minor;
      typeServeur  := details^.sv101_type;
      commentaire  := details^.sv101_comment;
      inc(details);
    end;
    result := EntriesRead ;
  end;
  NetApiBufferFree(Buffer);
end;


function  GetListeDisque (const strNomServeur : string ; Var ListeDisque : TTabDisque ) : integer;
//  ___________________________________________________________________________
// | Function  GetListeDisque                                                  |
// |___________________________________________________________________________|
// |  Permet de récuperrer la liste des disques sur un serveur                 |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur :String                                      |
// |        |   nom du serveur sur le quel chercher les disques                |
// |________|__________________________________________________________________|
// | Sortie | var ListeDisque : TTabString                                     |
// |        |   Liste des disques                                              |
// |        | Result : integer                                                 |
// |        |   -1 en cas d'erreur si non nombre de disuqe trouvé              |
// |________|__________________________________________________________________|

var

  buffer      : pointer;
  wNomServeur : WideString;
  Details     : Pdisque  ;
  Err         : longword;
  EntriesRead, TotalEntries : Cardinal;
  i : integer;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  wNomServeur := strNomServeur  ;
  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la recherche                                                *
// ****************************************************************************

  if APINetcharger(@_NetServerDiskEnum, 'NetServerDiskEnum') then
    err := _NetServerDiskEnum( pwidechar(wnomServeur),
                               0,
                               buffer,
                               LongueurMaxi ,
                               EntriesRead ,
                               TotalEntries ,
                               nil);
  if err = errSuccess then
  begin
    SetLength (ListeDisque , EntriesRead );
    details := Pdisque(buffer);
    for i := 0 to EntriesRead -1 do
    begin
      ListeDisque[i] := details^;
      inc(details);
    end;
    result :=EntriesRead ;
  end;
  NetApiBufferFree(Buffer);
end;

function  AjouteUtilisateurGroupeLocal (const strNomServeur : string; const strNomUtilisateur : string ;const strNomGoupe : string  ) : integer;

//  ___________________________________________________________________________
// | Function  AjouteUtilisateurGroupe                                         |
// |___________________________________________________________________________|
// |  Permet d'ajouter un utilisateur comme membre d'un groupe                 |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur :String                                      |
// |        |   nom du serveur sur le quel faires l'operation                  |
// |        | const strNomUtilisateur : string                                 |
// |        |   Nom de l'utilisateur                                           |
// |        | const strNomGoupe : string                                       |
// |        |   Nom du groupe dans le quel ajouter l'utilisateur               |
// |________|__________________________________________________________________|
// | Sortie | Result : integer                                                 |
// |        |   Liste des disques                                              |
// |        | Result : integer                                                 |
// |        |   0 en cas de reussite                                           |
// |________|__________________________________________________________________|

var
  wNomServeur :  WideString;
  wNomGroupe : WideString ;
  wNomUtilisateur : wideString;
  Err         : longword;
  Buffer : LOCALGROUP_MEMBERS_INFO_3;

begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  wNomServeur     := strNomServeur ;
  wNomGroupe      := strNomGoupe ;
  wNomUtilisateur := '\' + strnomutilisateur;
  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la modification                                             *
// ****************************************************************************
  Buffer.lgrpi3_domainandname := pwidechar(wNomUtilisateur );
  if APINetcharger(@_NetLocalGroupAddMembers, 'NetLocalGroupAddMembers') then
    err := _NetLocalGroupAddMembers(pwidechar(wnomServeur),
                                     pwidechar(wNomGroupe ),
                                     3,
                                     @buffer,
                                     1);
  result := err;
  erreur := err;
end;

function  AjouteUtilisateurGroupe (const strNomServeur : string; const strNomUtilisateur : string ;const strNomGoupe : string  ) : integer;

//  ___________________________________________________________________________
// | Function  AjouteUtilisateurGroupe                                         |
// |___________________________________________________________________________|
// |  Permet d'ajouter un utilisateur comme membre d'un groupe                 |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur :String                                      |
// |        |   nom du serveur sur le quel faires l'operation                  |
// |        | const strNomUtilisateur : string                                 |
// |        |   Nom de l'utilisateur                                           |
// |        | const strNomGoupe : string                                       |
// |        |   Nom du groupe dans le quel ajouter l'utilisateur               |
// |________|__________________________________________________________________|
// | Sortie | Result : integer                                                 |
// |        |   Liste des disques                                              |
// |        | Result : integer                                                 |
// |        |   0 en cas de reussite                                           |
// |________|__________________________________________________________________|

var
  wNomServeur :  WideString;
  wNomUtilisateur : WideString ;
  wNomGroupe      : widestring ;
  Err         : longword;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  wNomServeur := strNomServeur ;
  wNomUtilisateur := strNomUtilisateur ;
  wNomGroupe      := strNomGoupe ;
  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la modification                                             *
// ****************************************************************************

  if APINetcharger(@_NetGroupAddUser, 'NetGroupAddUser') then
    err := _NetGroupAddUser( pwidechar(wnomServeur),
                             pwidechar(wNomGroupe ),
                             pwidechar(wNomUtilisateur ));
  result := err;
  erreur := err;
end;

function  SupprimeUtilisateurGroupe (const strNomServeur : string; const strNomUtilisateur : string ;const strNomGoupe : string  ) : integer;

//  ___________________________________________________________________________
// | Function  SupprimeUtilisateurGroupe                                       |
// |___________________________________________________________________________|
// |  Permet de supprimer un utilisateur de la liste des membres d'un groupe   |
// |___________________________________________________________________________|
// | Entrée | const strNomServeur :String                                      |
// |        |   nom du serveur sur le quel faire l'operation                   |
// |        | const strNomUtilisateur : string                                 |
// |        |   Nom de l'utilisateur                                           |
// |        | const strNomGoupe : string                                       |
// |        |   Nom du groupe dans le quel ajouter l'utilisateur               |
// |________|__________________________________________________________________|
// | Sortie | Result : integer                                                 |
// |        |   Liste des disques                                              |
// |        | Result : integer                                                 |
// |        |   0 en cas de reussite                                           |
// |________|__________________________________________________________________|

var
  wNomServeur : WideString;
  wNomUtilisateur : WideString ;
  wNomGroupe      : widestring ;
  Err         : longword;
begin

// ****************************************************************************
// * Initialisation des variables                                             *
// ****************************************************************************

  wNomServeur := strNomServeur ;
  wNomUtilisateur := strNomUtilisateur ;
  wNomGroupe      := strNomGoupe ;
  result := -1;
  err := errMax;

// ****************************************************************************
// * Execution de la modification                                             *
// ****************************************************************************

  if APINetcharger(@_NetGroupDelUser, 'NetGroupDelUser') then
    err := _NetGroupDelUser( pwidechar(wnomServeur),
                             pwidechar(wNomGroupe ),
                             pwidechar(wNomUtilisateur ));
  result := err;
end;

function StringVersAdresseInet(AIP: string): TIPAddr;
var
  phe: PHostEnt;
  pac: PChar;
  wskInitialisationDonnees: TWSAData;
  tmpIpAddr: TIPAddr;
begin

  // Initialisation du winsock
  WSAStartup($101, wskInitialisationDonnees);
  try

    // Ici, comme la fonction l'indique on fait la résolution de nom
    // C'est une fonction de winsock
    // struct hostent* FAR gethostbyname(
    //      const char* name
    // );
    // Le retour de la fonction est un pointeur sur une structure appelée hostent
    //      typedef struct hostent {
    //        char FAR* h_name;
    //        char FAR  FAR** h_aliases;
    //        short h_addrtype;
    //        short h_length;
    //        char FAR  FAR** h_addr_list;
    //      } hostent;
    // C'est justement h_addr_list qui nous intéresse:
    // h_addr_list est une chaîne à 0 terminal avec la liste d'adresses retournée
    // par la résolution de nom. On va donc déréférencer les informations qui
    // nous sont utiles et les encapsuler dans le record TIPAddr, utile pour la fonction
    // de 'icmp.dll' qui est IcmpSendEcho

    phe := GetHostByName(PChar(AIP));
    if Assigned(phe) then
    begin
      pac := phe^.h_addr_list^;
      if Assigned(pac) then
      begin
        with TIPAddr(tmpIpAddr).S_un_b do begin
          s_b1 := Byte(pac[0]);
          s_b2 := Byte(pac[1]);
          s_b3 := Byte(pac[2]);
          s_b4 := Byte(pac[3]);
        end;
      end
      else
      begin
        // Une exception histoire de dire ce qui se passe quand ça marche pas
        raise Exception.Create('Erreur: Pas de résolution Hostname vers adresse IP');
      end;
    end
    else
    begin
      // Une deuxième exception si ça marche vraiment pas
      raise Exception.Create('Erreur: la recherche HostName a échouée');
    end;
  except

    // Voilà, la je rempli de noir (rien) caractère #0 tout ce qui n'a pas été renseigné
    // histoire de ne pas faire d'erreur en lançant la requête qui nous intéresse
    // plus tard
    FillChar(tmpIpAddr, SizeOf(tmpIpAddr), #0);
  end;

  // Winsock: on te fout la paix et on fait le ménage
  WSACleanup;
  Result:= tmpIpAddr;
end;

function booPing(InetAddress : string) : boolean;
var
 Handle : THandle;
 AdresseInet : IPAddr;
 DW : DWORD;
 rep : array[1..128] of byte;
begin
  result := false;
  // D'abord la fonction qui permet de mettre en route une requête Echo ICPM
  Handle := IcmpCreateFile;

  // Si ça marche pas on quitte.
  if Handle = INVALID_HANDLE_VALUE then
   Exit;

  // Je traduit l'adresse sous forme de chaîne vers une forme reconnue pour la requête:
  // le célèbre TIPAddr
  AdresseInet:= StringVersAdresseInet(InetAddress);

  // Enfin la requête. Elle retourne un numéro (ICMP_ECHO_REPLY) qui correspond au pointeur de la structure
  // des données (ReplyBuffer) qui ont été renvoyées. Si elle retourne 0, alors c'est que ça marche pas
  // Je ne m'étends pas plus sur le sujet mais on peut connaître la nature de l'erreur
  // en bidouillant un peu.
  DW := IcmpSendEcho (Handle, AdresseInet, nil, 0, nil, @rep, 128, 0);
  Result := (DW <> 0);
  // Bon j'ai le résulat, maintenant je ferme tout! bye!
  IcmpCloseHandle (Handle);
end;

function PartageToChemin(strServeur : string ; strPartage : string ; var strChemin : string) : boolean;

//  ___________________________________________________________________________
// | Function  PartageToChemin                                                 |
// |___________________________________________________________________________|
// |  Permet de trouver le chemin d'un partage sur un serveur                  |
// |___________________________________________________________________________|
// | Entrée | strServeur : string                                              |
// |        |   nom du serveur sur le quel faire l'operation                   |
// |        | strPartage : string                                              |
// |        |   Nom du partage                                                 |
// |________|__________________________________________________________________|
// | Sortie | Result : boolean                                                 |
// |        |   Vrai si ok sinon voir getufbreseauerreur pour l'erreur         |
// |        | var strChemin : string                                           |
// |        |   Chemin du partage                                              |
// |________|__________________________________________________________________|

var
  Share: PSHARE_INFO_2;
begin
  result := false;
  strChemin := '';
  erreur := NetShareGetInfo(PWideChar(WideString(strServeur)) ,
                        PWideChar(WideString(strPartage)) ,2,@share);
  result := erreur =0;
  strChemin := share^.shi2_path ;
end;

function ArretPartage(strserveur : string; strpartage : string):boolean;
var
  pserveur , ppartage : PWideChar ;
begin
  if strserveur ='' then pserveur := nil else pserveur := PWideChar(WideString(strserveur));
  ppartage := PWideChar(WideString(strpartage));
  erreur := NetShareDel(pserveur,ppartage,0);
  result := erreur = 0;
end;

function CreerPartage(strServeur :string; strChemin : string; strPartage : string) : boolean;

//  ___________________________________________________________________________
// | Function  CreerPartage                                                    |
// |___________________________________________________________________________|
// |  Permet de créer un partage sur un serveur                                |
// |___________________________________________________________________________|
// | Entrée | strServeur : string                                              |
// |        |   nom du serveur sur le quel faire l'operation                   |
// |        | strChemin : string                                               |
// |        |   Chemin du partage                                              |
// |        | strPartage : string                                              |
// |        |   Nom du partage                                                 |
// |________|__________________________________________________________________|
// | Sortie | Result : boolean                                                 |
// |        |   Vrai si ok sinon voir getufbreseauerreur pour l'erreur         |
// |________|__________________________________________________________________|

var
  Share: _SHARE_INFO_2;
  ParamErr: lpdword;
begin
  result := false;
  FillChar(Share, SizeOf(Share), 0);
//  new(sshare);
  Share.shi2_netname := PWideChar(WideString(strPartage));  //nom de partage du dossier
  Share.shi2_type := STYPE_DISKTREE;                        // disk drive
  Share.shi2_remark := '';                                  // zone commentaire
  Share.shi2_permissions := ACCESS_ALL;                     //definition des droits
  Share.shi2_max_uses := $FFFFFFFF;                         // nb max de users simultane
  Share.shi2_current_uses := 0;
  Share.shi2_path := PWideChar(WideString(strChemin));     //ici le chemin du dossier a partager
  Share.shi2_passwd := Nil;                                // nil si pas password
  erreur := NetShareAdd(PWideChar(WideString(strServeur)), 2, @SHARE, ParamErr);
//  dispose(sshare);
  result :=  erreur = 0;
end;

function getufbreseauerreur : string;
begin
  result := inttostr(erreur)+' :' + SysErrorMessage (erreur );
end;

function  ListerPartage(strserveur : string;var tabListePartage : TTabPartage ):boolean;
type
  TEAArray = Array [0..0] of EXPLICIT_ACCESS;
  PEAArray = ^TEAArray;
var
  p:             PShareInfo502Array;
  res,
  er, tr,
  resume,
  i,j:             DWORD;
  pwszServer:    PWideChar;
  lpwszServer:   Array [0..255] of WideChar;
  lpbDaclPresent : Bool;
  lpbDaclDefaulted : bool;
  pdacl: pacl;
  plistentries : PEAArray;
  nNbEntries : cardinal;
  index : integer;
  pentry : pchar;
  strtemp : tnomGroupe ;
begin
  setlength(tabListePartage ,0);
  index := 0;
  er:=0;
  tr:=0;
  resume:=0;
  if strserveur ='' then pwszServer := nil else
  begin
    if (Pos('\\', strserveur) <> 1) then strserveur:='\\'+strserveur;
     StringToWideChar(strServeur, @lpwszServer, SizeOf(lpwszServer));
     pwszServer:=@lpwszServer;
  end;

// ****************************************************************************
// * Recherche des partage present sur le serveur                             *
// ****************************************************************************

  if APINetcharger( @_NetShareEnum,'NetShareEnum') then
  repeat
    res := _NetShareEnum(pwszserver,502,@p,dword(-1),@er,@tr,@resume);
    if (res = ERROR_SUCCESS) or (res = ERROR_MORE_DATA) then
    begin
      for i:=1 to Pred(er) do
      begin
        index := index+1;
        setlength( tabListePartage , index);
        tabListePartage[index-1].strnom := WideCharToString(p^[i].shi502_netname);
        tabListePartage[index-1].strchemin := WideCharToString(p^[i].shi502_path);
        tabListePartage[index-1].intType := p^[i].shi502_type;
        setlength(tabListePartage[index-1].tabacl,0);

// ****************************************************************************
// * Chargement du descripteur de securité si present                         *
// ****************************************************************************

        if p^[i].shi502_security_dsc <> nil then
        begin
        GetSecurityDescriptorDacl(p^[i].shi502_security_dsc,lpbDaclPresent,pdacl ,lpbDaclDefaulted);
          IF lpbDaclPresent then
          if pdacl <> nil then
          begin
            if GetExplicitEntriesFromAcl(pDacl^, nNbEntries, @pListEntries) = ERROR_SUCCESS then
            begin

// ****************************************************************************
// * Enumeration des securité                                                 *
// ****************************************************************************

            try
              setlength(tabListePartage[index-1].tabacl,pDacl^.AceCount);
              for j := 0 to pDacl^.AceCount - 1 do
              begin
                strtemp :=sidtoName(pListEntries[j].Trustee.ptstrName);
                pEntry := (PChar(pListEntries) + j * sizeof(EXPLICIT_ACCESS)) ;
                tabListePartage[index-1].tabacl[j].accesmode := PEXPLICIT_ACCESS_(pEntry)^.grfAccessMode;
                tabListePartage[index-1].tabacl[j].accesrigth := PEXPLICIT_ACCESS_(pEntry)^.grfAccessPermissions;
                tabListePartage[index-1].tabacl[j].strDomain :=  strtemp.strdomaine;
                tabListePartage[index-1].tabacl[j].strNom := strtemp.strnom;
              end;
            finally
              LocalFree(HLOCAL(pListEntries));
            end;
            end;
          end;
        end;
      end;
      NetApiBufferFree(p);
    end else
     erreur:= res;
  until (res<> ERROR_MORE_DATA);
end;

function accessmodetostr(accessmode : ACCESS_MODE ):string;
begin
  case accessmode of
    NOT_USED_ACCESS   : result := 'NOT_USED_ACCESS';
    GRANT_ACCESS      : result := 'autoriser acces';
    SET_ACCESS        : result := 'établir acces';
    DENY_ACCESS       : result := 'refuser acces';
    REVOKE_ACCESS     : result := 'retirer acces';
    SET_AUDIT_SUCCESS : result := 'SET_AUDIT_SUCCESS';
    SET_AUDIT_FAILURE : result := 'SET_AUDIT_FAILURE';
  end;
end;
function accesspermissiontoStr(accesspermission : cardinal ) : string;
begin
  case accesspermission of
    $1f01ff : result := 'Accès complet';
    $1200a9 : result := 'Accès lecture';
    $1301bf : result := 'Accès modification';
  else result := inttostr( accesspermission);
  end;
end;

initialization

erreur := 0;

finalization
  ApiNetDecharger ;

end.







