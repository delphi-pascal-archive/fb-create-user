unit UfrmMaitre;

// |===========================================================================|
// | unit UfrmMaitre                                                           |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité gerant la fenêtre principale du programme fbcreateuser              |
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
// |   1.0.0.0 Création de l'unité                                             |
// |===========================================================================|

  {$DEFINE MESSAGERIE}
interface

uses
  OutlookXP,Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids,uuser, ExtCtrls,Contnrs, xmldom, XMLIntf, msxmldom,
  XMLDoc,ufrmSites,ufbfichier,dateutils,ufrminfo,ufbadsi,ugestdroit, Buttons,
  OleServer;

type
  TfrmMaitre = class(TForm)
    strgrid: TStringGrid;
    Timer1: TTimer;
    SaveDialog1: TSaveDialog;
    OpenDialog1: TOpenDialog;
    fichierUser: TXMLDocument;
    cmdSupprimer: TBitBtn;
    cmdAjouter: TBitBtn;
    cmdImporter: TBitBtn;
    cmdCharger: TBitBtn;
    cmdEnregistrer: TBitBtn;
    cmdCreer: TBitBtn;
    cmdEnvoyerMail: TBitBtn;
    cmdEditerSites: TBitBtn;
    OutlookApplication: TOutlookApplication;
    cmdEditerBanque: TBitBtn;
    fichierIni: TXMLDocument;
    procedure cmdEditerBanqueClick(Sender: TObject);
    procedure cmdEnvoyerMailClick(Sender: TObject);
    procedure cmdCreerClick(Sender: TObject);
    procedure cmdImporterClick(Sender: TObject);
    procedure cmdEditerSitesClick(Sender: TObject);
    procedure cmdChargerClick(Sender: TObject);
    procedure cmdEnregistrerClick(Sender: TObject);
    procedure strgridDblClick(Sender: TObject);
    procedure cmdSupprimerClick(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure strgridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure strgridDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure cmdAjouterClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    listeGroupe : TListBox ;
    ligneselectionne : integer;
    listeUser :Tobjectlist;
    XLMutilisateurs : IXMLNode ;
    XLMGroupe : IXMLNode ;
    user : tuser;
    strLdapRacine : string;
    strDomaine : string;
    StrSrvMessagerie : string;
    strMessage : string;
    strFormatDateCSV : String;
    strCSVNom : string;
    strCSVPrenom : string;
    strCSVAdresseMail : string;
    strCSVDateDepart : string;
    strCSVUid : string;
    strCSVNomSociete : string;
    strCSVLogin : string;
    strCSVNomEntite : string;
    strCSVNomSite : string;
    strCSVTelephone : string;
    strCSVFax : string;
    strCSVNomDepartement : string;
    strCSVMatricule : string;
    strCSVMatricule2 : string;
    strChangerMotPasse : string;
    function csvtodate(csvdate : string) : tdatetime;
    procedure chargeFichierIni;
    Procedure initialisestrGrid;
    Procedure RedimGrid;
    procedure ReafficheUser;
    function  createP( vuser : tuser ): boolean;
    procedure createUser( vuser : tuser );
    function  userExiste( vuser : tuser;index : integer =-1) : integer;
  public
    { Déclarations publiques }
    procedure ecritlog(strtext : string);
  end;

var
  frmMaitre: TfrmMaitre;
  strNomFichierLog : string;

implementation

{$R *.dfm}
uses
  ufrmUser,
  Types,
  ActiveDs_TLB,
  tsuserexlib_tlb,
  {$IFDEF MESSAGERIE}
  CDOEXM_TLB,
  {$ENDIF}
  activex,
  ComObj,
  mapi,
  strutils,
  ufrmBanques,
  ufbreseau,
  ufonctionDiv;

const
  versionFichierIni = '1.0.0.0';
  versionFichierUtilisateur = '2.0.0.0';

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Privée                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmMaitre.chargeFichierIni;

//  ___________________________________________________________________________
// | procedure TfrmMaitre.chargeFichierIni                                     |
// | _________________________________________________________________________ |
// || permet de charger le fichier de configuration                           ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  xmlIni : IXMLNode;
  strNomfichierConfiguration : string;
  boook : boolean;
begin
  fichierIni.Active := false;
  fichierIni.xml.Clear ;
  fichierIni.FileName :='';
  fichierIni.Active := true;
  strNomfichierConfiguration := FBChangeFileExt(Application.ExeName ,'.fbini');

// ****************************************************************************
// * Test l'existance du fichier de configuration                             *
// ****************************************************************************

  if not FBFileExists(strNomfichierConfiguration) then
  begin
    frmMaitre.ecritlog('Le fichier ' + FBExtractFileName (strNomfichierConfiguration) +' n''existe pas');
    frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierConfiguration) +' n''existe pas');
    exit;
  end;

// ****************************************************************************
// * Test l'intégrité du fichier de configuration                             *
// ****************************************************************************

  try
    fichierIni.LoadFromFile(strNomfichierConfiguration );
  except
    on e : exception do
    begin
      frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierConfiguration) +#10#13+
                              '      n''est pas un fichier FBINI valide' + #10#13 + e.Message);
      fichierIni.Active := false;
      frmMaitre.ecritlog('/!\ Le fichier ' + FBExtractFileName (strNomfichierConfiguration) + ' n''est pas un fichier FBINI valide');
      exit;
    end;
  end;

// ****************************************************************************
// * Test la version du fichier de configuration                              *
// ****************************************************************************

  boook := fichierIni.DocumentElement.HasAttribute('Version');
  if boook then boook := fichierIni.DocumentElement.GetAttributeNS('Version','') = versionFichierIni;
  if not boook then
  begin
    frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierConfiguration) +#10#13+
                            '      n''est pas à la bonne version');
    fichierIni.Active := false;
    frmMaitre.ecritlog('/!\ Le fichier ' + FBExtractFileName (strNomfichierConfiguration) + ' n''est pas à la bonne version');
    fichierIni.Active := false;
    exit;
  end;

// ****************************************************************************
// * Chargement de la configuration                                           *
// ****************************************************************************

  xmlini := fichierIni.DocumentElement.ChildNodes[0];
  try
    strFormatDateCSV := xmlini.getAttributeNS('strFormatDateCSV','');
    strCSVNom := xmlini.getAttributeNS('strCSVNom','');
    strCSVPrenom := xmlini.getAttributeNS('strCSVPrenom','');
    strCSVAdresseMail := xmlini.getAttributeNS('strCSVAdresseMail','');
    strCSVDateDepart := xmlini.getAttributeNS('strCSVDateDepart','');
    strCSVUid := xmlini.getAttributeNS('strCSVUid','');
    strCSVNomSociete := xmlini.getAttributeNS('strCSVNomSociete','');
    strCSVLogin := xmlini.getAttributeNS('strCSVLogin','');
    strCSVNomEntite := xmlini.getAttributeNS('strCSVNomEntite','');
    strCSVNomSite := xmlini.getAttributeNS('strCSVNomSite','');
    strCSVTelephone := xmlini.getAttributeNS('strCSVTelephone','');
    strCSVFax := xmlini.getAttributeNS('strCSVFax','');
    strCSVNomDepartement := xmlini.getAttributeNS('strCSVNomDepartement','');
    strCSVMatricule := xmlini.getAttributeNS('strCSVMatricule','');
    strCSVMatricule2 := xmlini.getAttributeNS('strCSVMatricule2','');
    strLdapRacine := xmlini.getAttributeNS('strLdapRacine','');
    strDomaine := xmlini.getAttributeNS('strDomaine','');
    strChangerMotPasse := xmlini.getAttributeNS('strChangerMotPasse','');
    StrSrvMessagerie := xmlini.getAttributeNS('StrSrvMessagerie','');
    strMessage := xmlini.getAttributeNS('strMessage','');
  except
    frmInfo.AfficheMessage('  /!\ Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierConfiguration) +
                   #10#13 + '      ce n''est pas un fichier fbini valide');
    frmMaitre.ecritlog('/!\ Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierConfiguration) );
  end;
  fichierIni.Active := false;
end;

procedure TfrmMaitre.ecritlog(strtext : string);

//  ___________________________________________________________________________
// | procedure ecritlog(strtext : string);                                     |
// | _________________________________________________________________________ |
// || permet d'écrire dans un fichier log                                     ||
// ||_________________________________________________________________________||
// || Entrées | strtext : string                                              ||
// ||         |   text à ajouter au fichier                                   ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  fichier : textfile;
  nb : longword;
  username : array[0..255] of char;
begin
  try
    nb := 255 ;
    getusername(username,nb);
    strtext := DateTimeToStr(Now )+ ';"'+username+'";' + strtext;
    if strNomFichierLog ='' then strNomFichierLog := fbChangeFileExt(application.ExeName,'.fblog');
    assignfile(fichier ,strNomFichierLog);
    if FileExists (strNomFichierLog) then
      append(fichier) else
      rewrite(fichier);
    writeln(fichier,strtext);
    CloseFile (fichier );
  except
  end;
end;


function SendMail1(const Subject, MessageText, MailFromName, MailFromAddress,
  MailToName, MailToAddress: String; const attachments: array of String; WinHandle: THandle = 0):boolean;

//  ___________________________________________________________________________
// | function SendMail                                                         |
// | _________________________________________________________________________ |
// || Permet la création d'un mail avec le client de messagerie par defaut    ||
// ||_________________________________________________________________________||
// || Entrées | const Subject : String                                        ||
// ||         |   Sujet du message                                            ||
// ||         | const MessageText : String                                    ||
// ||         |   Test du message                                             ||
// ||         | const MailFromName : String                                   ||
// ||         |   Nom de l'expediteur                                         ||
// ||         | const MailFromAddress : String                                ||
// ||         |   Adresse de l'expediteur                                     ||
// ||         | const MailToName : String                                     ||
// ||         |   Nom du destibataire                                         ||
// ||         | const MailToAddress : String                                  ||
// ||         |   Adresse du destinataire                                     ||
// ||         | const attachments: array of String                            ||
// ||         |   Pieces jointes                                              ||
// ||         | WinHandle: THandle = 0                                        ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  MAPIError: DWord;
  MapiMessage: TMapiMessage;
  Originator, Recipient: TMapiRecipDesc;
  Files, FilesTmp: PMapiFileDesc;
  FilesCount: Integer;
begin
  FillChar(MapiMessage, Sizeof(TMapiMessage), 0);
  MapiMessage.ulReserved := 0;
  MapiMessage.lpszSubject := PChar(Subject);
  MapiMessage.lpszNoteText := PChar(MessageText);
  MapiMessage.lpszMessageType := nil;
  FillChar(Originator, Sizeof(TMapiRecipDesc), 0);
  Originator.lpszName := PChar(MailFromName);
  Originator.lpszAddress := PChar(MailFromAddress);
  MapiMessage.lpOriginator := @Originator;
  MapiMessage.nRecipCount := 0;  //
  FillChar(Recipient, Sizeof(TMapiRecipDesc), 0);
  Recipient.ulRecipClass := MAPI_TO;
  Recipient.lpszName := PChar(MailToName);
  Recipient.lpszAddress := PChar(MailToAddress);
  MapiMessage.lpRecips := @Recipient;
  MapiMessage.nFileCount := High(attachments) - Low(attachments) + 1;
  Files := AllocMem(SizeOf(TMapiFileDesc) * MapiMessage.nFileCount);
  MapiMessage.lpFiles := Files;
  FilesTmp := Files;
  for FilesCount := Low(attachments) to High(attachments) do
  begin
    FilesTmp.nPosition := $FFFFFFFF;
    FilesTmp.lpszPathName := PChar(attachments[FilesCount]);
    Inc(FilesTmp)
  end;
  try
    MAPIError := MapiSendMail(0, 0,
      MapiMessage, MAPI_DIALOG {or MAPI_LOGON_UI or MAPI_NEW_SESSION}, 0);
    result:= MAPIError = 0;
  finally
    FreeMem(Files)
  end
end;

Procedure TfrmMaitre.RedimGrid ;

//  ___________________________________________________________________________
// | Procedure TfrmMaitre.RedimGrid                                            |
// | _________________________________________________________________________ |
// || Permet de redimensionner les colonnes en fonction de leurs contenus     ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j,k : integer;
  largeur : integer;
begin

// ****************************************************************************
// * Scan de toutes les colonnes                                              *
// ****************************************************************************

  for j := 0 to strgrid.ColCount -1 do
  begin
    largeur := 0;

// ****************************************************************************
// * Scan de toutes les lignes                                                *
// ****************************************************************************

    for i := 0 to strgrid.RowCount -1 do
    begin

// ****************************************************************************
// * Mise à jour en fonction du contenu de la cellule                         *
// ****************************************************************************

      if strgrid.Canvas.TextWidth(strgrid.Cells[j,i])+5 > largeur then
      begin
        largeur := strgrid.Canvas.TextWidth(strgrid.Cells[j,i])+5;
        strgrid.ColWidths [j] := largeur;
      end;

// ****************************************************************************
// * Mise à jour en fonction des groupes                                      *
// ****************************************************************************

      if (j=20) and(i>0) and ( listeUser.Count > 0) then
      for k := 0 to tuser(listeuser[i-1]).listeGroupes.Count-1 do
        if listeGroupe.Canvas.TextWidth ( tuser(listeuser[i-1]).listeGroupes[k])+15 > largeur then
        begin
          largeur := listeGroupe.Canvas.TextWidth ( tuser(listeuser[i-1]).listeGroupes[k])+15;
          strgrid.ColWidths [j] := largeur;
        end;
    end;
  end;
end;

procedure TfrmMaitre.ReafficheUser ;

//  ___________________________________________________________________________
// | procedure TfrmMaitre.ReafficheUser                                        |
// | _________________________________________________________________________ |
// || Rafraichi l'affichage des utilisateurs                                  ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j : integer;
  usertemp : tuser;
begin
  initialisestrGrid;
  if listeUser.Count = 0 then  strgrid.RowCount := 2 else strgrid.RowCount := listeUser.Count +1;
  for i := 0 to  listeUser.Count -1 do
  begin
    j:=i+1;
    usertemp := tuser(listeUser.items[i]);
    strgrid.cells[ 1,j] := usertemp.strNomComplet ;
    strgrid.cells[ 2,j] := usertemp.strNom;
    strgrid.cells[ 3,j] := usertemp.strPrenom;
    strgrid.cells[ 4,j] := usertemp.strLogin;
    strgrid.cells[ 5,j] := usertemp.strUid ;
    strgrid.cells[ 6,j] := usertemp.strAppartenance ;
    strgrid.cells[ 7,j] := usertemp.strSIte ;
    strgrid.cells[ 8,j] := usertemp.strSociete ;
    strgrid.cells[ 9,j] := datetostr(usertemp.dtDateExpiration );
    strgrid.cells[10,j] := usertemp.strDescription;
    strgrid.cells[11,j] := usertemp.strPassword ;
    if usertemp.booPasswordNeverExpire then
      strgrid.cells[12,j] := 'X'
    else
      strgrid.cells[12,j] := '';
    strgrid.Cells[13,j] := usertemp.strProfileTse ;
    strgrid.cells[14,j] := usertemp.strVille;
    strgrid.cells[15,j] := usertemp.strCompagnie ;
    strgrid.cells[16,j] := usertemp.strLogonScript ;
    strgrid.cells[17,j] := usertemp.strTel;
    strgrid.cells[18,j] := usertemp.strFax ;
    strgrid.cells[19,j] := usertemp.strDepartement;
    case usertemp.intEtat of
      0   : strgrid.cells[0,j] := 'New';
      1   : strgrid.cells[0,j] := 'New';
      2   : strgrid.cells[0,j] := '/!\';
      4   : strgrid.cells[0,j] := '/!\';
      $F0 : strgrid.cells[0,j] := '\/';
    end;
  end;
  RedimGrid;
end;

Procedure TfrmMaitre.initialisestrGrid ;

//  ___________________________________________________________________________
// | Procedure TfrmMaitre.initialisestrGrid                                    |
// | _________________________________________________________________________ |
// || Permet de réinitialiser la stringgrille                                 ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j : integer;
begin
  strgrid.ColCount := 21;
  for i := 0 to strgrid.RowCount -1 do
    for j := 0 to strgrid.ColCount -1 do
      strgrid.Cells [j,i] :='';
  strgrid.RowCount := 2;
  strgrid.Cells[0,0] := '/!\';
  strgrid.cells[1,0] := 'Nom complet';
  strgrid.cells[2,0] := 'Nom';
  strgrid.cells[3,0] := 'Prénom';
  strgrid.cells[4,0] := 'Login';
  strgrid.cells[5,0] := 'UID';
  strgrid.cells[6,0] := 'Appartenance';
  strgrid.cells[7,0] := 'Site';
  strgrid.cells[8,0] := 'Nom Société';
  strgrid.cells[9,0] := 'date départ';
  strgrid.cells[10,0] := 'Description';
  strgrid.cells[11,0] := 'Mot de passe';
  strgrid.Cells[12,0] := 'NE';
  strgrid.cells[13,0] := 'Profile TSE';
  strgrid.cells[14,0] := 'Ville';
  strgrid.cells[15,0] := 'Compagnie';
  strgrid.cells[16,0] := 'Logon Script';
  strgrid.cells[17,0] := 'N° tel';
  strgrid.cells[18,0] := 'N° Fax';
  strgrid.cells[19,0] := 'Département';
  strgrid.Cells[20,0] := 'Groupes';
  RedimGrid;
end;

function  TfrmMaitre.userExiste( vuser : tuser;index : integer =-1 ) : integer;

//  ___________________________________________________________________________
// | function  TfrmMaitre.userExiste                                           |
// | _________________________________________________________________________ |
// || permet de savoir si un utilisateur est present dans la liste via son    ||
// || login ou le nom complet                                                 ||
// ||_________________________________________________________________________||
// || Entrées | vuser : tuser                                                 ||
// ||         |   utilisateur à tester                                        ||
// ||         | index : integer =-1                                           ||
// ||         |   index de l'utilisateur à eviter de tester                   ||
// ||_________|_______________________________________________________________||
// || Sorties | result : integer                                              ||
// ||         |   -1 si pas trouvé si non index de l'utilisateur trouvé       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  result := -1;
  for i := 0 to listeUser.Count-1 do
    if (index =-1) or (i<>index) then
    if (lowercase(tuser(listeuser.Items[i]).strLogin) = lowercase(vuser.strLogin) ) or
       (lowercase(tuser(listeuser.Items[i]).strNomComplet) = lowercase(vuser.strNomComplet )) then
      result := i;
end;

function TfrmMaitre.csvtodate(csvdate : string) : tdatetime;

//  ___________________________________________________________________________
// | function TfrmMaitre.csvtodate                                             |
// | _________________________________________________________________________ |
// || Permet de convertir une date 20100101 en date classic                   ||
// ||_________________________________________________________________________||
// || Entrées | csvdate : string                                              ||
// ||         |   date a convertir au format string                           ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : tdatetime                                            ||
// ||         |   date convertie                                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  y,d,m ,ly: integer;
begin
  ly := 4;
  y := pos('yyyy',strFormatDateCSV);
  if y = 0 then
  begin
    y := pos('yy',strFormatDateCSV);
    ly := 2 ;
  end;
  d := pos('dd',strFormatDateCSV);
  m := pos('mm',strFormatDateCSV);
  result := -1;
  if (y = 0)or(d = 0) or(m = 0) then exit;
  y := strtoint(copy( csvdate,y,ly));
  d := strtoint(copy( csvdate,d,2));
  m := strtoint(copy( csvdate,m,2));
  result := EncodeDate(y,m,d);
end;

function TfrmMaitre.createP( vuser : tuser ) : boolean;

//  ___________________________________________________________________________
// | function TfrmMaitre.createP                                               |
// | _________________________________________________________________________ |
// || Permet de créer le repertoire personnel d'un user sur un serveur        ||
// ||_________________________________________________________________________||
// || Entrées | vuser : tuser                                                 ||
// ||         |   utilisateur pour le quel créer le répertoire                ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |  true si ok                                                   ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  err : integer;
  boook : boolean;
  strCheminReel : string;
  strCheminUNC : string;
  SID : PSID;
begin
  result := false;
  frmInfo.AfficheMessage('  Création du P sur ' + vuser.strSRV + ' ...' );
  strCheminUNC := '\\' + vuser.strSRV +'\' + vuser.strSRVPartage;
  if not PartageToChemin(vuser.strSRV ,vuser.strSRVPartage  ,strCheminReel )then
  begin
    frmInfo.AfficheMessage('    /!\ Erreur lors de la recuperration des information sur le serveur');
    frmInfo.AfficheMessage('    /!\ ' + GetUfbReseauErreur );
    vuser.intEtat := 4;
    vuser.strErreur :=  vuser.strErreur + '/!\ Création P Erreur lors de la recuperration des information sur le serveur : ' + GetUfbReseauErreur ;
    exit;
  end;
  if vuser.strSRVRepertoire <>'' then
  begin
    strCheminReel := strCheminReel + '\' +vuser.strSRVRepertoire ;
    strCheminUNC := strCheminUNC + '\' +vuser.strSRVRepertoire ;
  end;
  strCheminUNC := strCheminUNC + '\' + vuser.strLogin;
  strCheminReel := strCheminReel +'\' + vuser.strLogin ;
  while pos('\\',strCheminReel ) <> 0 do strCheminReel := StringReplace(strCheminReel,'\\','\', [rfReplaceAll]); 
  if not FBDirectoryCreate(strCheminUNC ) then
  begin
    err := getlasterror;
    frmInfo.AfficheMessage('    /!\ Erreur lors de la  tentative de création du répertoire sur le serveur');
    frmInfo.AfficheMessage('    /!\ ' + strCheminUNC);
    frmInfo.AfficheMessage('    /!\ ' + strCheminReel);
    frmInfo.AfficheMessage('    /!\ ' +inttostr(err)+' :' + SysErrorMessage (err ));
    vuser.intEtat := 4;
    vuser.strErreur :=  vuser.strErreur + '/!\ Création P Erreur lors de la  création du répertoire : '+SysErrorMessage (err ) ;
    exit;
  end;
  GetAccountSID(vuser.strLogin,SID);
  if not SetFileObjectAccessRights( strCheminUNC ,
                                   SID,
                                   GENERIC_READ + GENERIC_WRITE + GENERIC_EXECUTE+_DELETE,
                                   false) then
  begin
    err := getlasterror;
    frmInfo.AfficheMessage('    /!\ Erreur lors de la  tentative d''attribution des droits sur le répertoire ');
    frmInfo.AfficheMessage('    /!\ Suppression du répertoire sur le serveur');
    RemoveDir(strCheminUNC  );
    vuser.intEtat := 4;
    vuser.strErreur :=  vuser.strErreur + '/!\ Création P Erreur lors de l''attribution des droits : '+SysErrorMessage (err ) ;
    exit;
  end;
  if SID <> nil then
  begin
    HeapFree(GetProcessHeap(), 0, Sid);
    SID:=nil;
  end;
  boook := CreerPartage('\\'+vuser.strSRV ,strCheminReel  , vuser.strLogin + '$');
  if not boook then
  begin
    frmInfo.AfficheMessage('    /!\ Erreur lors de la tentative de création du partage ');
    frmInfo.AfficheMessage('    /!\ ' + strCheminUNC);
    frmInfo.AfficheMessage('    /!\ ' + strCheminReel);
    frmInfo.AfficheMessage('    /!\ ' + GetUfbReseauErreur );
    frmInfo.AfficheMessage('    /!\ Suppression du répertoire sur le serveur');
    RemoveDir(strCheminUNC );
    vuser.intEtat := 4;
    vuser.strErreur :=  vuser.strErreur + '/!\ Création P Erreur lors de la création du partage : '+GetUfbReseauErreur ;
    exit;
  end;
  frmInfo.AfficheMessage('  Création du P OK');
  result := true;
end;

function booUserExist (strLogin : string) : boolean;

//  ___________________________________________________________________________
// | function  booUserExist                                                    |
// | _________________________________________________________________________ |
// || Permet de s'avoir si l'utilisateur existe dans l'ad                     ||
// ||_________________________________________________________________________||
// || Entrées | strLogin : string                                             ||
// ||         |   login de l'utilisateur                                      ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |  vrai si  trouvé                                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  ReferencedDomain: pointer;
  cbSid: dword;
  cchReferencedDomain: dword;
  peUse: SID_NAME_USE;
  bSuccess: BOOL;
  waccountname : Pwidechar;
  SID : PSID;
begin
  Result:=FALSE;
  bSuccess:=FALSE;
  ReferencedDomain:=nil;
  cbSid:=128;
  cchReferencedDomain:=16;
  waccountname := PWideChar(WideString(strLogin ));
  try
    SID:=HeapAlloc(GetProcessHeap(),0,cbSid);
    if SID = nil then Exit;
    ReferencedDomain:=HeapAlloc(GetProcessHeap(),0,cchReferencedDomain*2);
    if ReferencedDomain = nil then Exit;
    while NOT LookupAccountNameW(nil,waccountname,Sid,cbSid,ReferencedDomain,cchReferencedDomain,peUse) do
    begin
      if GetLastError() = ERROR_INSUFFICIENT_BUFFER then
      begin
        Sid:=HeapReAlloc(GetProcessHeap(),0,Sid,cbSid);
        if SID = nil then Exit;
        ReferencedDomain:=HeapReAlloc(GetProcessHeap(),0,ReferencedDomain,cchReferencedDomain*2);
        if ReferencedDomain = nil then Exit;
      end else
        Exit;
    end;
    bSuccess := IsValidSid( SID);
  finally
    HeapFree(GetProcessHeap(), 0, ReferencedDomain);
  end;
  if(SID <> nil)then
  begin
    HeapFree(GetProcessHeap(), 0, Sid);
    SID:=nil;
  end;
  Result:=bSuccess;
end;

procedure  TfrmMaitre.createUser( vuser : tuser );

//  ___________________________________________________________________________
// | procedure  TfrmMaitre.createUser                                          |
// | _________________________________________________________________________ |
// || Permet de créer un utilisateur tuser dans l'AD                          ||
// ||_________________________________________________________________________||
// || Entrées | vuser : tuser                                                 ||
// ||         |   Utilisateur à créer                                         ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  adComp : IADsContainer;
  adUser : IADsUser;
  adGrp : IADsGroup ;
  TSrvUsr : IADsTSUserEx;
  tt : ttime;
  h,m,s,c : word;
  boook : boolean;
  borelance : boolean;
  i : integer;
  userAccountControl : integer;
  strLDAPUser : string;
  strLdapGroupe : string;
  v:variant;
  {$IFDEF MESSAGERIE}
  mailbox : IMailboxStore;
  emailrec : IMailRecipient;
  {$ENDIF}
Begin
  frmInfo.AfficheMessage('');
  frmInfo.AfficheMessage('****************************************************');
  frmInfo.AfficheMessage('Création du compte : ' + vuser.strNomComplet );
  frmInfo.AfficheMessage('  dans l''OU : ' + vuser.strOU);
  ADsGetObject(vuser.strOU,IADsContainer,adcomp);

// ****************************************************************************
// * tentative de création de l'utilisateur                                   *
// ****************************************************************************

  vuser.intEtat := $F0;
  vuser.strErreur := 'Création OK';
  try
    adUser := adComp.Create('user', 'CN=' + trim(vuser.strNomComplet) ) as IADsUser;
    adUser.Put('samAccountName', trim(vuser.strLogin));
    adUser.Put('userPrincipalName', trim(vuser.strLogin)+'@'+strDomaine);
    adUser.SetInfo;
  except
    on e: exception do
    begin
      frmInfo.AfficheMessage('  /!\ Erreur lors de la création du compte : ' +e.Message);
      vuser.intEtat := 2;
      vuser.strErreur := e.Message;
      ecritlog('"' + vuser.strNomComplet + '";"' + vuser.strErreur +'"');
      exit;
    end;
  end;

// ****************************************************************************
// * Mise à jour des informations complementaires                             *
// ****************************************************************************

  frmInfo.AfficheMessage('  Création compte OK mise à jour des informations ...');
  adUser.Put('givenName',trim(vuser.strPrenom));
  adUser.Put('sn',trim(vuser.strNom));
  adUser.Put('DisplayName',trim(vuser.strNomComplet));
  adUser.Put('Description',trim(vuser.strDescription));
  aduser.Put('UID',trim(vuser.strUid));
  aduser.SetPassword(trim(vuser.strPassword ));
  if vuser.booPasswordNeverExpire then
  begin
    userAccountControl :=aduser.Get('userAccountControl');
    userAccountControl := userAccountControl or ADS_UF_DONT_EXPIRE_PASSWD;
    aduser.Put('userAccountControl', userAccountControl );
  end else
    aduser.Put('pwdLastSet',0);
  if trim(vuser.strLogonScript) <>'' then
    adUser.Put('scriptPath',trim(vuser.strLogonScript));
  if trim(vuser.strDepartement )<>'' then
    aduser.Put('Department',trim(vuser.strDepartement ));
  if trim(vuser.strCompagnie ) <>'' then
    aduser.Put('Company',trim(vuser.strCompagnie));
  if trim(vuser.strAddresseMail ) <>'' then
    aduser.Put('EmailAddress',trim(vuser.strAddresseMail));
  if trim(vuser.strVille ) <>'' then
    aduser.Put('l',trim(vuser.strVille ));    // ville
  if vuser.dtDateExpiration > 0 then
    aduser.AccountExpirationDate := IncDay ( vuser.dtDateExpiration,1) ;
  adUser.SetInfo;
  adUser.QueryInterface(IID_IADsTSUserEx, TSrvUsr);
  if trim (vuser.strProfileTse) <>'' then
    TSrvUsr.TerminalServicesProfilePath := trim (vuser.strProfileTse);
  adUser.SetInfo;
  if trim (vuser.strTel)<>'' then
    aduser.Put('telephoneNumber',trim (vuser.strTel));
  if trim (vuser.strFax)<>'' then
    aduser.Put('facsimileTelephoneNumber',trim (vuser.strFax));
  aduser.AccountDisabled := false;
  aduser.SetInfo;
  frmInfo.AfficheMessage('  Mise à jour des informations OK');

// ***************************************************************************
// * Attente replication du compte pour le p et l'ajout des groupes          *
// ***************************************************************************

  repeat
    tt := gettime;
    tt := IncMinute(tt,1);
    boook := false;
    frmInfo.AfficheMessage('  Attente réplication du compte (max 1Min) ...');
    while (not booUserExist  (vuser.strLogin ))and (gettime<tt) do
    begin
      Application.ProcessMessages ;
      Sleep(1000);
    end;
    if booUserExist  (vuser.strLogin ) then
    begin
      tt := incminute(tt,-1);
      decodetime(gettime-tt,h,m,s,c);
      frmInfo.AfficheMessage('  Réplication du compte OK en ' + format('%.2D : %.2D : %.2D , %.2D' ,[h,m,s,c]));
      boook := true;
    end ;
    if not boook then boook := MessageDlg('Réplication NOK Voulez vous attendre 1 min de plus ?' ,mtConfirmation ,[mbyes,mbno],0)=mrno;
  until boook ;
  boook := booUserExist(vuser.strLogin );
  if not boook then
  begin
    frmInfo.AfficheMessage('  /!\ Réplication du compte NOK');
    frmInfo.AfficheMessage('Fin création du compte : ' + vuser.strNomComplet);
    frmInfo.AfficheMessage('****************************************************');
    frmInfo.AfficheMessage('');
    vuser.intEtat := 2 ;
    vuser.strErreur := 'Réplication du compte NOK';
    ecritlog('"' + vuser.strNomComplet + '";"' + vuser.strErreur +'"');
    exit;
  end;

// ****************************************************************************
// * Ajout dans les groupes                                                   *
// ****************************************************************************

  strLDAPUser := 'LDAP://' + adUser.Get('distinguishedName');
  for i := 0 to vuser.listeGroupes.Count -1 do
  begin
    ADsRechercheUtilisateur(vuser.listeGroupes[i],strLdapRacine,strLdapGroupe);
    if strLdapGroupe <> '' then
    begin
      ADsGetObject(StrLdapGroupe,iadsgroup,adgrp);
      try
        adgrp.Add(strLDAPUser );
        adgrp.SetInfo;
        frmInfo.AfficheMessage('  Ajout au groupe : ' + vuser.listeGroupes[i] + ' OK');
      except
        on e : exception do
        begin
          frmInfo.AfficheMessage('  /!\ Erreur lors de l''ajout de : '+ vuser.listeGroupes[i] + ' ' + e.Message  );
          vuser.intEtat := 4;
          vuser.strErreur := vuser.strErreur + 'Erreur lors de l''ajout de : '+ vuser.listeGroupes[i] + ' ' + e.Message + '#10#13';
        end;
      end;
    end else
      frmInfo.AfficheMessage('  /!\ Groupe '+ vuser.listeGroupes[i] + ' non trouvé' );
  end;

// ****************************************************************************
// * Création du P                                                            *
// ****************************************************************************

  if vuser.strLecteur <> '' then
    if createP(vuser) then
    try
      aduser.Put('HomeDirectory',(WideString(vuser.strRepertoirePerso )));
      aduser.put('homeDrive',vuser.strLecteur );
      adUser.SetInfo;
      frmInfo.AfficheMessage('  Ajout du P au profil OK');
    except
      on e: exception do
      begin
        frmInfo.AfficheMessage('  /!\ lors de l''ajout du P au profil : ' + e.Message );
        vuser.intEtat := 4 ;
        vuser.strErreur :=  vuser.strErreur + ' Erreur lors de l''ajout de du P au profil' +#13#10;
      end
    end;

  {$IFDEF MESSAGERIE}
  if vuser.strBanque <> '' then
  try

// ****************************************************************************
// * Création de la messagerie                                                *
// ****************************************************************************

    mailbox := aduser as IMailboxStore ;
    mailbox.CreateMailbox(vuser.strBanque);

// ****************************************************************************
// * Ajout du matricule dans l'extentionAttribute1                            *
// ****************************************************************************

    if vuser.strMatricule <>'' then
      aduser.Put('extensionAttribute1',vuser.strMatricule);
    aduser.SetInfo ;

// ****************************************************************************
// * Attente création de la bal par le serveur                                *
// ****************************************************************************

    repeat
      tt := gettime;
      tt := IncMinute(tt,2);
      boook := false;
      frmInfo.AfficheMessage('  Attente Réplication bal (max 2 min)');
      while (not boook )and (gettime <tt)  do
      begin
        try
          if ADsGetObject(strLDAPUser ,iadsuser,aduser)= s_ok then
          begin
            v := aduser.Get('ProxyAddresses');  // si bal ok pas d'exception
            boook := true;
            tt := incminute(tt,-1);
            decodetime(gettime-tt,h,m,s,c);
            frmInfo.AfficheMessage('  Réplication de la bal OK en ' + format('%.2D : %.2D : %.2D , %.2D' ,[h,m,s,c]));
          end else
            boook := false;
        except
          boook := false; // bal nom créée encore
        end;
        Application.ProcessMessages ;
        Sleep(1000);
      end;
      borelance := not boook;
      if not boook then borelance := MessageDlg('Réplication NOK Voulez vous attendre 2 min de plus ?' ,mtConfirmation ,[mbyes,mbno],0)=mryes;
    until not borelance ;
    if not boook then
    begin
      frmInfo.AfficheMessage('  /!\ Réplication de la bal NOK');
      frmInfo.AfficheMessage('Fin création du compte : ' + vuser.strNomComplet);
      frmInfo.AfficheMessage('****************************************************');
      frmInfo.AfficheMessage('');
      vuser.intEtat := 2 ;
      vuser.strErreur := 'Réplication de la BAL NOK';
      ecritlog('"' + vuser.strNomComplet + '";"' + vuser.strErreur +'"');
      exit;
    end;

// ****************************************************************************
// * Passe les adresses générées en non principale  'SMTP:' -> 'smtp:'        *
// ****************************************************************************

    mailbox := aduser as IMailboxStore ;
    emailrec := mailbox as MailRecipient;
    v := emailrec.ProxyAddresses;
    for i := VarArrayLowBound(v, 1) To VarArrayHighBound(v, 1) do
    begin
      v[i]:= StringReplace (v[i],'SMTP:','smtp:',[rfIgnoreCase]) ;
    end;

// ****************************************************************************
// * Ajout de l'adresse principale                                            *
// ****************************************************************************

    VarArrayRedim(v,VarArrayHighBound(v, 1)+1);
    v[VarArrayHighBound(v, 1)] := 'SMTP:'+ vuser.strAddresseMail ;
    emailrec.ProxyAddresses := v;
    aduser.SetInfo ;
    frmInfo.AfficheMessage('  Ajout adresse ok');

// ****************************************************************************
// * Désaction de l'option de mise a jour automatique de l'adresse            *
// ****************************************************************************

    emailrec.AutoGenerateEmailAddresses := false;
    aduser.SetInfo ;
    frmInfo.AfficheMessage('  création bal ok');
  except
    on e:exception do
    begin
      frmInfo.AfficheMessage('  /!\ erreur création bal : '+ e.Message);
      vuser.intEtat := 4 ;
      vuser.strErreur :=  vuser.strErreur + '  Erreur lors de la création de la BAL ' + e.Message + #10#13;
    end;
  end;
  {$ENDIF}

  frmInfo.AfficheMessage('Fin création du compte : ' + vuser.strNomComplet);
  frmInfo.AfficheMessage('****************************************************');
  frmInfo.AfficheMessage('');
  ecritlog('"' + vuser.strNomComplet + '";"' + vuser.strErreur +'"');
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Evénementiel                                                       #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmMaitre.FormCreate(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.FormCreate                                           |
// | _________________________________________________________________________ |
// ||                                                                         ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  frmInfo.AfficheMessage('Démarrage');
  chargeFichierIni;
  Application.ProcessMessages;
  DateSeparator :='/';
  ShortDateFormat := 'dd/mm/yyyy';
  Caption := 'fb Create User' + ' V' + StrLectureVersion + ' ' + StrLectureCopyright ;
  ecritlog('"** Démarrage '+StrLectureNomProduit + ' V' + StrLectureVersion+'"');
  listeGroupe := TListBox.Create(strgrid);
  listeGroupe.Parent := strgrid ;
  listeGroupe.Top := - listeGroupe.Height ;
  listegroupe.Width := 30;
  ligneselectionne:=1;
  listeUser :=TObjectList.Create(true) ;
  initialisestrGrid ;
  WindowState := wsMaximized ;
end;

procedure TfrmMaitre.cmdAjouterClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdAjouterClick                                      |
// | _________________________________________________________________________ |
// || Permet d'ajouter un nouvel utilisateur                                  ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  boook : boolean;
  doublon : boolean;
begin

// ****************************************************************************
// * Création d'une nouvelle instance de tuser                                *
// ****************************************************************************

  user := tUser.create ;

// ****************************************************************************
// * Passage d'an frmUser pour saisie et verification si pas de création d'un *
// * doublon avec nom complet et login                                        *
// ****************************************************************************

  repeat
    doublon := false;
    boook := TfrmUser.editerUser(user);
    if boook then doublon := userExiste(user) <> -1;
    if doublon then
      MessageDlg('Le nom complet ou le login existe déjà dans la liste' , mterror,[mbok],0);
  until not doublon;

// ****************************************************************************
// * Ajout du nouvelle user dans la liste et affichage dans la stringgrille si*
// * non liberation de l'instance créer                                       *
// ****************************************************************************

  if boook then
  begin
    user.intEtat := 1;
    listeUser.add(user);
    ReafficheUser;
  end else user.Free ;
end;

procedure TfrmMaitre.strgridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.strgridDrawCell                                      |
// | _________________________________________________________________________ |
// || Permet l'affichage de la liste des groupe dans un listebox              ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||         | ACol: Integer                                                 ||
// ||         |   N° collone de la cellule redessinée                         ||
// ||         | ARow: Intege                                                  ||
// ||         |   N° ligne de la cellule redessinée                           ||
// ||         | Rect: TRect                                                   ||
// ||         |   coordonée de la cellule redessinée                          ||
// ||         | State: TGridDrawState                                         ||
// ||         |   etat                                                        ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin

// ****************************************************************************
// * Si la collone n'est pas la derniere on masque la liste des groupes       *
// ****************************************************************************

  if (acol <> 20) then listeGroupe.Visible := false;
  if (acol = 20)and(arow = ligneselectionne ) then
  begin

// ****************************************************************************
// * Si la collone est la derniere et que la ligne est celle en cours on      *
// * affiche la liste des groupes                                             *
// ****************************************************************************

    listeGroupe.left := rect.Left ;
    listegroupe.Width := rect.Right - rect.Left ;
    if rect.Top + listeGroupe.Height > strgrid.ClientHeight then
      listegroupe.Top := rect.Bottom - listegroupe.Height
    else
      listeGroupe.Top := rect.Top;
    listeGroupe.Clear ;

// ****************************************************************************
// * Chargement de la liste des groupe du user en cours                       *
// ****************************************************************************

    if listeUser.Count > 0 then
    begin
      for i := 0 to tuser(listeUser.items[ligneselectionne -1]).listeGroupes.count-1 do
        listeGroupe.Items.Add(tuser(listeUser.items[ligneselectionne -1]).listeGroupes[i]);
      listeGroupe.Visible := true;
    end;
  end;
end;

procedure TfrmMaitre.FormDestroy(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.FormDestroy                                          |
// | _________________________________________________________________________ |
// || Permet la liberration des objets créés au demarrage                     ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  listeGroupe.Free;
  listeUser.Free;
  ecritlog('"** Arrêt **"');
end;

procedure TfrmMaitre.strgridSelectCell(Sender: TObject; ACol, ARow: Integer;var CanSelect: Boolean);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.strgridSelectCell                                    |
// | _________________________________________________________________________ |
// || Permet l'affichage de la liste des groupes                              ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||         | ACol: Integer                                                 ||
// ||         |   N° collone de la cellule redessinée                         ||
// ||         | ARow: Intege                                                  ||
// ||         |   N° ligne de la cellule redessinée                           ||
// ||_________|_______________________________________________________________||
// || Sorties | var CanSelect: Boolean                                        ||
// ||         |   true pour autoriser la selection                            ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  timer1.Enabled := false;
  if acol = 20 then
  begin
    canselect := false;
    exit;
  end;
  listeGroupe.Visible := false;
  ligneselectionne := Arow;
  timer1.Enabled := true;
end;

procedure TfrmMaitre.Timer1Timer(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.Timer1Timer                                          |
// | _________________________________________________________________________ |
// || relance le dessin de la stringgrille pour afficher la liste des groupes ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  timer1.Enabled := false;
  strgrid.Repaint;
end;

procedure TfrmMaitre.strgridDblClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.strgridDblClick                                      |
// | _________________________________________________________________________ |
// || Permet lors du double clic d'éditer un utilisateur                      ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  userTemp : tuser;
  boook : boolean;
  doublon : boolean;
begin
  if listeUser.Count > 0 then
  begin

// ****************************************************************************
// * Création d'un instance temporaire de tuser et copie des informations de  *
// * L'utilisateur en cours dedans                                            *
// ****************************************************************************

    userTemp := tuser.create;
    usertemp.copy ( tuser(listeUser.items[ligneselectionne -1]));
    repeat
      doublon := false;
      boook := TfrmUser.editerUser(userTemp);
      if boook then doublon := userExiste(userTemp,ligneselectionne -1) <> -1;
      if doublon then
        MessageDlg('Le nom complet ou le login existe déjà dans la liste' , mterror,[mbok],0);
    until not doublon;

// ****************************************************************************
// * Modification sans doublon OK on copie les information de l'instance      *
// * temporaire dans le user en cours                                         *
// ****************************************************************************

    if boook then tuser(listeUser.items[ligneselectionne -1]).copy(usertemp);
    ReafficheUser ;

// ****************************************************************************
// * Liberration de l'instance temporaire                                     *
// ****************************************************************************

    usertemp.Free ;
  end;
end;

procedure TfrmMaitre.cmdSupprimerClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdSupprimerClick                                    |
// | _________________________________________________________________________ |
// || Pemret la suppression du user en cours                                  ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if listeUser.Count = 0 then exit;
  if ligneselectionne >0 then
    if MessageDlg('Etes vous sur de vouloir supprimer' + #10#13 +
                  tuser(listeUser.items[ligneselectionne -1]).strNomComplet,mtConfirmation ,[mbyes,mbno],0)=mryes then
    begin
      listeUser.Delete(ligneselectionne -1);
      ReafficheUser;
    end;
end;

procedure TfrmMaitre.cmdEnregistrerClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdEnregistrerClick                                  |
// | _________________________________________________________________________ |
// || permet l'enregistrement dans un fichier XML des utilisateurs            ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j : integer;
  userTemp : tuser;
  strtemp : string;
begin
  if listeUser.Count >0 then
    if SaveDialog1.Execute then
    begin
      fichierUser.Active := false;
      fichierUser.xml.Clear ;
      fichierUser.FileName :='';
      fichierUser.Active := true;

// ****************************************************************************
// * Ajout des infos de version XML                                           *
// ****************************************************************************

      fichierUser.Version :='1.0';
      fichierUser.StandAlone :='no';
      fichierUser.encoding := 'UTF-8';
      fichierUser.DocumentElement := fichierUser.CreateElement('Utilisateurs','');
      fichierUser.DocumentElement.SetAttributeNS('Version','',versionFichierUtilisateur);
      for i := 0 to listeUser.Count -1 do
      begin

// ****************************************************************************
// * Sauvegarde d'un user dans un noeud XML                                   *
// ****************************************************************************

        userTemp := tuser(listeUser.Items[i]);
        XLMutilisateurs := fichierUser.DocumentElement.AddChild('Utilisateur');
        XLMutilisateurs.SetAttributeNS('NomComplet','',userTemp.strNomComplet);
        XLMutilisateurs.SetAttributeNS('Nom','',userTemp.strNom);
        XLMutilisateurs.SetAttributeNS('Prenom','',userTemp.strPrenom);
        XLMutilisateurs.SetAttributeNS('Login','',userTemp.strLogin);
        XLMutilisateurs.SetAttributeNS('Description','',userTemp.strDescription);
        XLMutilisateurs.SetAttributeNS('Uid','',userTemp.strUid);
        XLMutilisateurs.SetAttributeNS('Password','',userTemp.strPassword);
        XLMutilisateurs.SetAttributeNS('NE','',userTemp.booPasswordNeverExpire);
        XLMutilisateurs.SetAttributeNS('ChangePassword','',userTemp.booChangePassword);
        XLMutilisateurs.SetAttributeNS('LogonScript','',userTemp.strLogonScript);
        XLMutilisateurs.SetAttributeNS('Departement','',userTemp.strDepartement);
        XLMutilisateurs.SetAttributeNS('Apartenance','',userTemp.strAppartenance);
        XLMutilisateurs.SetAttributeNS('Site','',userTemp.strSIte);
        XLMutilisateurs.SetAttributeNS('Societe','',userTemp.strSociete);
        XLMutilisateurs.SetAttributeNS('Mail','',userTemp.strAddresseMail);
        XLMutilisateurs.SetAttributeNS('Ville','',userTemp.strVille);
        XLMutilisateurs.SetAttributeNS('Compagnie','',userTemp.strCompagnie);
        strtemp := DateTimeToStr(userTemp.dtDateExpiration);
        XLMutilisateurs.SetAttributeNS('DateExpiration','',strtemp);
        XLMutilisateurs.SetAttributeNS('RepertoirePerso','',userTemp.strRepertoirePerso);
        XLMutilisateurs.SetAttributeNS('Lecteur','',userTemp.strLecteur);
        XLMutilisateurs.SetAttributeNS('ProfileTse','',userTemp.strProfileTse);
        XLMutilisateurs.SetAttributeNS('tel','',userTemp.strTel);
        XLMutilisateurs.SetAttributeNS('fax','',userTemp.strFax);
        XLMutilisateurs.SetAttributeNS('OU','',userTemp.strOU);
        XLMutilisateurs.SetAttributeNS('Matricule','',usertemp.strMatricule );

// ****************************************************************************
// * Sauvegarde des groupes dans un sous noeud XML                            *
// ****************************************************************************

        for j := 0 to userTemp.listeGroupes.Count -1 do
        begin
          XLMGroupe := XLMutilisateurs.AddChild('Groupe');
          XLMGroupe.SetAttributeNS('Nom','',userTemp.listeGroupes[j]);
        end;
        XLMutilisateurs.SetAttributeNS('SRV','', user.strSRV );
        XLMutilisateurs.SetAttributeNS('SRVPartage','', user.strSRVPartage );
        XLMutilisateurs.SetAttributeNS('SRVRepertoire','',user.strSRVRepertoire );
        XLMutilisateurs.SetAttributeNS('Banque','',user.strBanque );
      end;

// ****************************************************************************
// * Sauvegarde dans le fichier                                               *
// ****************************************************************************

      fichierUser.SaveToFile(SaveDialog1.FileName );
    end;
end;

procedure TfrmMaitre.cmdChargerClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdChargerClick                                      |
// | _________________________________________________________________________ |
// || permet de charger les user depuis un fichier XML                        ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j : integer;
  nbuser : integer;
  strtemp : string;
  boook : boolean;
begin
  OpenDialog1.Title := 'Ouvrir un fichier fbUser';
  OpenDialog1.Filter := 'Fichier fbUser|*.fbuser';
  OpenDialog1.FilterIndex :=1;
  if OpenDialog1.Execute then
  begin
    listeUser.Clear;
    fichierUser.Active := false;
    fichierUser.xml.Clear ;
    fichierUser.FileName :='';
    try
    fichierUser.LoadFromFile(OpenDialog1.FileName );
    except
      on e : exception do
      begin
        MessageDlg('Le fichier ' + FBExtractFileName (OpenDialog1.FileName) +#10#13+
                   'n''est pas un fichier FBUSER valide' + #10#13 + e.Message , mtError ,[mbok],0 );
        fichierUser.Active := false;
        exit;
      end;
    end;
    boook := fichierUser.DocumentElement.HasAttribute('Version');
    if boook then boook := fichierUSer.DocumentElement.GetAttributeNS('Version','') = versionFichierUtilisateur;
    if not boook then
    begin
      MessageDlg('Le fichier ' + FBExtractFileName (OpenDialog1.FileName) +#10#13+
                 'n''est pas un fichier FBUSER valide pas la bonne version trouvée' , mtError ,[mbok],0 );
      fichierUser.Active := false;
      exit;
    end;
    try
      nbuser := fichierUser.DocumentElement.ChildNodes.Count;
      for i := 0 to nbuser  - 1 do
      begin
        XLMutilisateurs := fichierUser.DocumentElement.ChildNodes[i] ;
        user := tuser.create ;
        user.strNomComplet := XLMutilisateurs.GetAttributeNS('NomComplet','');
        user.strNom := XLMutilisateurs.GetAttributeNS('Nom','');
        user.strPrenom := XLMutilisateurs.GetAttributeNS('Prenom','');
        user.strLogin := XLMutilisateurs.GetAttributeNS('Login','');
        user.strDescription := XLMutilisateurs.GetAttributeNS('Description','');
        user.strUid := XLMutilisateurs.GetAttributeNS('Uid','');
        user.strPassword := XLMutilisateurs.GetAttributeNS('Password','');
        user.booPasswordNeverExpire := XLMutilisateurs.GetAttributeNS('NE','');
        user.booChangePassword := XLMutilisateurs.GetAttributeNS('ChangePassword','');
        user.strLogonScript := XLMutilisateurs.GetAttributeNS('LogonScript','');
        user.strDepartement := XLMutilisateurs.GetAttributeNS('Departement','');
        user.strAppartenance := XLMutilisateurs.GetAttributeNS('Apartenance','');
        user.strSIte := XLMutilisateurs.GetAttributeNS('Site','');
        user.strSociete := XLMutilisateurs.GetAttributeNS('Societe','');
        user.strAddresseMail := XLMutilisateurs.GetAttributeNS('Mail','');
        user.strVille := XLMutilisateurs.GetAttributeNS('Ville','');
        user.strCompagnie := XLMutilisateurs.GetAttributeNS('Compagnie','');
        strtemp := XLMutilisateurs.GetAttributeNS('DateExpiration','');
        user.dtDateExpiration := StrToDateDef(strtemp,0);
        user.strRepertoirePerso := XLMutilisateurs.GetAttributeNS('RepertoirePerso','');
        user.strLecteur := XLMutilisateurs.GetAttributeNS('Lecteur','');
        user.strProfileTse := XLMutilisateurs.GetAttributeNS('ProfileTse','');
        user.strTel := XLMutilisateurs.GetAttributeNS('tel','');
        user.strFax := XLMutilisateurs.GetAttributeNS('fax','');
        user.strOU := XLMutilisateurs.GetAttributeNS('OU','');
        user.strMatricule := XLMutilisateurs.GetAttributeNS('Matricule','');
        for j := 0 to XLMutilisateurs.ChildNodes.Count -1 do
        begin
          XLMGroupe := XLMutilisateurs.ChildNodes[j];
          user.listeGroupes.Add(XLMGroupe.GetAttributeNS('Nom',''));
        end;
        user.strSRV := XLMutilisateurs.GetAttributeNS('SRV','');
        user.strSRVPartage := XLMutilisateurs.GetAttributeNS('SRVPartage','');
        user.strSRVRepertoire := XLMutilisateurs.GetAttributeNS('SRVRepertoire','');
        user.strBanque := XLMutilisateurs.GetAttributeNS('Banque','');
        user.intEtat := 1 ;
        listeUser.Add(user);
      end;
    except
      MessageDlg('Erreur lors du traitement du fichier ' + FBExtractFileName (OpenDialog1.FileName) +
                  #10#13 + 'ce n''est pas un fichier fbuser valide' ,mtError ,[mbok],0);
      user.free;
    end;
    ReafficheUser;
    fichierUser.Active := false;
  end;
end;

procedure TfrmMaitre.cmdEditerSitesClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdEditerSitesClick                                  |
// | _________________________________________________________________________ |
// || Permet d'editer les entités                                             ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if MessageDlg('Attention toutes erreur dans les modifications des entités' + #10#13 +
                 'engendrent des erreur lors de la création des utilisateurs' +#10#13 +
                 'estes vous sur de vouloir continuer ?' ,mtConfirmation ,[mbyes,mbno],0 ) = mryes then
    frmSites.ShowModal ;
end;

procedure TfrmMaitre.cmdEditerBanqueClick(Sender: TObject);
begin
  if MessageDlg('Attention toutes erreur dans les modifications des banques' + #10#13 +
                 'de messagerie engendrent des erreur lors de la création des utilisateurs' +#10#13 +
                 'estes vous sur de vouloir continuer ?' ,mtConfirmation ,[mbyes,mbno],0 ) = mryes then
    tfrmBanque.EditBanques;
end;


procedure TfrmMaitre.cmdImporterClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdImporterClick                                     |
// | _________________________________________________________________________ |
// || Permet d'importer des user depuis un export de l'ACD                    ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet appelant                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  fichiercsv : tstrings;
  listeCol : tstrings;
  intNom : integer;
  intPrenom : integer;
  intMail : integer;
  intdate : integer;
  intUID : integer;
  intSociete : integer;
  intLogin : integer;
  intapartenance : integer;
  intSite : integer;
  intTelephone : integer;
  intFax : integer;
  intDepartement : integer;
  intMatricule : integer;
  i : integer;
  booErreur : boolean;
  boook : boolean;
  doublon : boolean;
begin
  OpenDialog1.Title := 'Ouvrir un fichier';
  OpenDialog1.Filter := 'Fichier |*.*';
  OpenDialog1.FilterIndex :=1;
  if OpenDialog1.Execute then
  begin

// ****************************************************************************
// * repérage des colonnes indispensables                                     *
// ****************************************************************************

    booerreur := false;
    fichiercsv := tstringlist.Create ;
    listecol := tstringlist.Create ;
    fichiercsv.LoadFromFile(opendialog1.FileName );
    listecol.Delimiter :=';';
    fichiercsv.Strings[0] := '"' + StringReplace(fichiercsv.Strings[0],';','";"',[rfReplaceAll] )+'"';
    listeCol.DelimitedText := fichiercsv.Strings[0];
    intNom := listecol.IndexOf(strCSVNom) ;
    if intnom =-1 then booerreur:=true;
    intPrenom := listecol.IndexOf(strCSVPrenom);
    if intPrenom =-1 then booerreur:=true;
    intMail := listecol.IndexOf(strCSVAdresseMail);
    if intMail =-1 then booerreur:=true;
    intdate := listecol.IndexOf(strCSVDateDepart);
    if intdate =-1 then booerreur:=true;
    intUID := listecol.IndexOf(strCSVUid);
    if intUID =-1 then booerreur:=true;
    intSociete := listecol.IndexOf(strCSVNomSociete);
    intLogin := listecol.IndexOf(strCSVLogin);
    if intLogin =-1 then booerreur:=true;
    intapartenance := listecol.IndexOf(strCSVNomEntite);
    if intapartenance =-1 then booerreur:=true;
    intSite := listecol.IndexOf(strCSVNomSite);
    if intSite =-1 then booerreur:=true;
    intTelephone  := listecol.IndexOf(strCSVTelephone);
    if intTelephone =-1 then booerreur := true;
    intFax := listecol.IndexOf(strCSVFax);
    if intFax =-1 then booerreur := true;
    intDepartement := listecol.IndexOf(strCSVNomDepartement);
    // les exterieurs n'ont pas de departement
    intMatricule := listecol.IndexOf(strCSVMatricule);
    if intMatricule =-1 then intMatricule := listecol.IndexOf(strCSVMatricule2);
    if intMatricule =-1 then booerreur := true;

// ****************************************************************************
// * Si il manque une colonnes le fichier n'est pas bon                       *
// ****************************************************************************

    if booErreur then MessageDlg('Le fichier ' + FBExtractFileName( opendialog1.FileName )+' n''est pas conforme',mtError ,[mbok],0 );
    if not booerreur then
    for i := 1 to fichiercsv.Count -1 do
    begin

// ****************************************************************************
// * Pour chaque ligne du fichier on edite le User                            *
// ****************************************************************************

      fichiercsv.Strings[i] := '"' + StringReplace(fichiercsv.Strings[i],';','";"',[rfReplaceAll] )+'"';
      listecol.DelimitedText:= fichiercsv.Strings[i];
      user := tuser.create ;
      user.strNom := listecol.Strings[intnom];
      user.strPrenom := listecol.Strings[intprenom];
      user.strUid := listecol.Strings[intuid];
      user.strLogin  := listecol.Strings[intlogin];
      user.strAddresseMail := listecol.Strings[intmail];
      user.strAppartenance := listecol.Strings[intapartenance];
      user.strSIte := listecol.Strings[intsite];
      if intSociete = -1 then user.strSociete :='' else user.strSociete := listecol.Strings[intSociete];
      if YearOf(csvtodate(listecol.Strings[intdate]))< 2099 then
        user.dtDateExpiration := csvtodate(listecol.Strings[intdate])
      else
        user.dtDateExpiration :=0;
      user.strTel := listeCol.Strings [intTelephone];
      user.strFax := listeCol.Strings [intfax];
      user.strMatricule := listeCol.Strings [intmatricule];
      if intDepartement <> - 1 then
        user.strDepartement := listeCol.Strings [intdepartement]
      else
        user.strDepartement :='';

// ****************************************************************************
// * Evite les doublons de nom complet et de login                            *
// ****************************************************************************

      repeat
        doublon := false;
        boook := TfrmUser.editerUser(user);
        if boook then doublon := userExiste(user) <> -1;
        if doublon then
          MessageDlg('Le nom complet ou le login existe déjà dans la liste' , mterror,[mbok],0);
      until not doublon;

      if boook then
      begin
        user.intEtat := 1 ;
        listeUser.Add(user);
      end
      else
        user.free;
    end;
    listecol.Free;
    fichiercsv.Free ;
    ReafficheUser ;
  end;
end;

procedure TfrmMaitre.cmdCreerClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmMaitre.cmdCreerClick                                        |
// | _________________________________________________________________________ |
// || lance la création de utilisateur dans l'AD                              ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var i : integer;
begin
  try
  if listeUser.Count =0 then exit;
  frmInfo.AfficheMessage('****************************************************');
  frmInfo.AfficheMessage('*  Démarrage création des utilisateurs             *');
  frmInfo.AfficheMessage('****************************************************');
  Application.ProcessMessages ;
  for i:= 0 to  listeUser.Count - 1 do
  begin
    ecritlog(inttostr(i));
    createUser(tuser(listeuser.Items[i]));
    ReafficheUser ;
  end;
  except
    on e : exception do
    begin
      showmessage(e.Message );
      frmInfo.AfficheMessage('****************************************************');
      frmInfo.AfficheMessage('*  /!\ erreur lors de la création des utilisateurs *');
    end;
  end;
  frmInfo.AfficheMessage('****************************************************');
  frmInfo.AfficheMessage('*  Fin de la création des utilisateurs             *');
  frmInfo.AfficheMessage('****************************************************');
end;

procedure TfrmMaitre.cmdEnvoyerMailClick(Sender: TObject);
var
  i : integer;
  usertemp : tuser;
  strtemp : string;
  strtmp : string;
var
  EspaceDeNom: NameSpace;
  UnMail : MailItem;

begin
  if listeUser.Count =0 then exit;
  OutlookApplication.Connect ;
  EspaceDeNom := OutlookApplication.GetNamespace('MAPI');
  EspaceDeNom.Logon('','',false,false);
  for i := 0 to listeUser.Count -1 do
  begin
    usertemp := tuser(listeUser.Items[i]);
    strtmp := 'Création du compte de ' + usertemp.strPrenom + ' '+ usertemp.strNom ;
    strtemp := StringReplace(strMessage,'%login%',usertemp.strLogin ,[rfReplaceAll,rfIgnoreCase]);
    strtemp := StringReplace(strtemp,'%nom%',usertemp.strNom ,[rfReplaceAll,rfIgnoreCase]) ;
    strtemp := StringReplace(strtemp,'%prenom%',usertemp.strPrenom ,[rfReplaceAll,rfIgnoreCase]) ;
    strtemp := StringReplace(strtemp,'%srvMessagerie%',StrSrvMessagerie ,[rfReplaceAll,rfIgnoreCase]) ;
    strtemp := StringReplace(strtemp,'%mail%',usertemp.strAddresseMail ,[rfReplaceAll,rfIgnoreCase]) ;
    strtemp := StringReplace(strtemp,'%password%',usertemp.strPassword ,[rfReplaceAll,rfIgnoreCase]) ;
    if usertemp.booPasswordNeverExpire then
      strtemp := StringReplace(strtemp,'%changermotpasse%','',[rfReplaceAll,rfIgnoreCase])
    else
      strtemp := StringReplace(strtemp,'%changermotpasse%',strChangerMotPasse,[rfReplaceAll,rfIgnoreCase]);
    UnMail:=OutlookApplication.CreateItem(olMailItem) as MailItem;
    UnMail.Subject := strtmp;
    UnMail.Body := strtemp;
    unMail.Display (false) ;
  end;
  EspaceDeNom.Logoff ;
end;

end.
