unit ufrmSites;

// |===========================================================================|
// | unit ufrmSites                                                            |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité permettant la gestion des entités et des sites pour le programme    |
// | fbcreateuser                                                              |
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

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, xmldom, XMLIntf, msxmldom, XMLDoc,ufbfichier,Contnrs,uentite
  ,ufrmGroupe,ubanques;

type
  TfrmSites = class(TForm)
    Label1: TLabel;
    cmdAjouterEntite: TButton;
    cmdSupprimerEntite: TButton;
    lstEntite: TListBox;
    edtNomEntite: TEdit;
    Label2: TLabel;
    lstSite: TListBox;
    Label4: TLabel;
    cmdEnregistrerEntite: TButton;
    edtNomSite: TEdit;
    edtOUSite: TEdit;
    edtCity: TEdit;
    edtSrvP: TEdit;
    lstGroupeGroupe: TListBox;
    lstGroupeExterieur: TListBox;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    cmdAjouterSite: TButton;
    cmdSupprimerSite: TButton;
    cmdEnregistrerSite: TButton;
    edtProfileTse: TEdit;
    Label11: TLabel;
    edtNomComplet: TEdit;
    Label12: TLabel;
    edtPassword: TEdit;
    cbxNeverExpire: TCheckBox;
    Label13: TLabel;
    cmdAjouterGrpGroupe: TButton;
    cmdAjouterGRPExterieur: TButton;
    cmdSupprimerGRPGroupe: TButton;
    cmdSupprimerGRPExterieur: TButton;
    Label14: TLabel;
    edtLoginScript: TEdit;
    edtSRVPartage: TEdit;
    Label15: TLabel;
    edtLecteurP: TEdit;
    Label16: TLabel;
    edtsrvRepertoire: TEdit;
    Label17: TLabel;
    fichierSites: TXMLDocument;
    lstBanqueExchangeGroupe: TListBox;
    lstBanqueExchangeExterieur: TListBox;
    Label3: TLabel;
    Label18: TLabel;
    cmdSupprimerBanqueGroupe: TButton;
    cmdAjouteBanqueGroupe: TButton;
    cmdSupprimerbanqueExterieur: TButton;
    cmdAjouterBanqueExterieur: TButton;
    procedure cmdSupprimerbanqueExterieurClick(Sender: TObject);
    procedure cmdAjouterBanqueExterieurClick(Sender: TObject);
    procedure cmdSupprimerBanqueGroupeClick(Sender: TObject);
    procedure cmdAjouteBanqueGroupeClick(Sender: TObject);
    procedure cmdSupprimerSiteClick(Sender: TObject);
    procedure cmdSupprimerGRPExterieurClick(Sender: TObject);
    procedure cmdAjouterGRPExterieurClick(Sender: TObject);
    procedure cmdSupprimerGRPGroupeClick(Sender: TObject);
    procedure cmdAouterGrproupeClick(Sender: TObject);
    procedure lstSiteClick(Sender: TObject);
    procedure cmdEnregistrerSiteClick(Sender: TObject);
    procedure cmdAjouterSiteClick(Sender: TObject);
    procedure lstEntiteClick(Sender: TObject);
    procedure cmdEnregistrerEntiteClick(Sender: TObject);
    procedure cmdSupprimerEntiteClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure cmdAjouterEntiteClick(Sender: TObject);
  private
    { Déclarations privées }
    entite : tentite;
    site   : tsite;
    listeBanques : TObjectList ;
    xmlEntite : IXMLNode;
    xmlSite   : IXMLNode;
    xmlGrpGroupe : IXMLNode ;
    xmlgrpExterieur : IXMLNode ;
    procedure chargerFichier ;
    procedure enregistrerFichier;
    procedure activerDetailEntite ( boo : boolean );
    procedure viderDetailEntite ;
    procedure viderDetailSite ;
    procedure activerSite ( boo : boolean );
    procedure activerDetailSite ( boo : boolean );
  public
    { Déclarations publiques }
    procedure supprimerIDBanque ( id : integer);
  end;

var
  frmSites : TfrmSites;

implementation

{$R *.dfm}
uses ufrmmaitre,ufrmBanques,ufrminfo;
const
  versionFichierUtilisateur = '2.0.0.0';

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Privée                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmSites.activerDetailEntite ( boo : boolean );

//  ___________________________________________________________________________
// | procedure TfrmSites.activerDetailEntite                                   |
// | _________________________________________________________________________ |
// || Permet d'activé ou non la partiedétail                                  ||
// ||_________________________________________________________________________||
// || Entrées | boo : boolean                                                 ||
// ||         |   vrai on active                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  edtNomEntite.Enabled := boo;
  edtProfileTse.Enabled := boo;
  edtNomComplet.Enabled := boo;
  edtpassword.Enabled := boo;
  edtLoginScript.Enabled := boo;
  cbxNeverExpire.Enabled := boo;
end;

procedure TfrmSites.activerDetailSite ( boo : boolean );

//  ___________________________________________________________________________
// | procedure TfrmSites.activerDetailSite                                     |
// | _________________________________________________________________________ |
// || PPermet d'activé la partie détail                                       ||
// ||_________________________________________________________________________||
// || Entrées | boo : boolean                                                 ||
// ||         |   vrai on active                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  edtNomSite.Enabled := boo;
  edtOUSite.Enabled := boo;
  edtCity.Enabled := boo;
  edtsrvP.Enabled := boo;
  edtSRVPartage.Enabled := boo;
  edtsrvRepertoire.Enabled := boo;
  edtlecteurP.Enabled := boo;
  lstGroupeGroupe.Enabled := boo;
  lstGroupeExterieur.Enabled := boo;
  cmdAjouterGrpGroupe.Enabled := boo;
  cmdSupprimerGRPGroupe.Enabled := boo;
  cmdAjouterGRPExterieur.Enabled := boo;
  cmdSupprimerGRPExterieur.Enabled := boo;
  cmdAjouterSite.Enabled := boo;
  cmdEnregistrerSite.Enabled := boo;
  cmdSupprimerSite.Enabled := boo;
end;

procedure TfrmSites.activerSite ( boo : boolean );

//  ___________________________________________________________________________
// | procedure TfrmSites.activerSite                                           |
// | _________________________________________________________________________ |
// || Permet d'activé la partie site                                          ||
// ||_________________________________________________________________________||
// || Entrées | boo : boolean                                                 ||
// ||         |   vrai on active                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstSite.Enabled := boo;
  cmdAjouterSite.Enabled := boo;
  cmdEnregistrerSite.Enabled := boo;
  cmdSupprimerSite.Enabled := boo;
  activerDetailSite (boo);
end;

procedure TfrmSites.viderDetailEntite ;

//  ___________________________________________________________________________
// | procedure TfrmSites.viderDetailEntite                                     |
// | _________________________________________________________________________ |
// || permet de vider tous les champs de la partie detail entité              ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  edtNomEntite.text := '';
  edtProfileTse.Text := '';
  edtNomComplet.Text := '%societe% %nom% %prenom% (%appartenance%)';
  edtpassword.Text := 'password+1';
  edtLoginScript.Text := 'logonscript.vbs';
  cbxNeverExpire.Checked := false;
end;

procedure TfrmSites.viderDetailSite ;

//  ___________________________________________________________________________
// | procedure TfrmSites.viderDetailSite                                       |
// | _________________________________________________________________________ |
// || permet de vider les champs de la partie detail Site                     ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  edtNomSite.Text :='';
  edtOUSite.text :='';
  edtCity.Text :='';
  edtsrvP.Text :='';
  edtSRVPartage.text := '';
  edtsrvRepertoire.text := '';
  edtlecteurP.Text :='';
  lstGroupeGroupe.Items.Clear ;
  lstGroupeExterieur.Items.clear;
  lstBanqueExchangeGroupe.Clear;
  lstBanqueExchangeExterieur.Clear;
end;

procedure TfrmSites.enregistrerFichier;

//  ___________________________________________________________________________
// | procedure TfrmSites.enregistrerFichier                                    |
// | _________________________________________________________________________ |
// || Permet d'enregistrer le fichier des entités                             ||
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
  entiteTemp : tentite;
  siteTemp : tsite;
begin
  fichierSites.Active := false;
  fichierSites.xml.Clear ;
  fichierSites.FileName :='';
  fichierSites.Active := true;

// ****************************************************************************
// * Ajout des infos de version XML                                           *
// ****************************************************************************

  fichierSites.Version :='1.0';
  fichierSites.StandAlone :='no';
  fichierSites.encoding := 'UTF-8';
  fichierSites.DocumentElement := fichierSites.CreateElement('Entites','');
  fichierSites.DocumentElement.SetAttributeNS('Version','',versionFichierUtilisateur);
  for i := 0 to lstEntite.Count -1 do
  begin

// ****************************************************************************
// * Enregistrement de chaque entité                                          *
// ****************************************************************************

    entiteTemp := tentite(lstEntite.items.Objects[i]);
    xmlEntite := fichierSites.DocumentElement.AddChild('Entite');
    xmlEntite.SetAttributeNS('Nom','',entiteTemp.strNom );
    xmlEntite.SetAttributeNS('ProfileTse','',entiteTemp.strProfileTSE );
    xmlEntite.SetAttributeNS('NomComplet','',entiteTemp.strNomComplet );
    xmlEntite.SetAttributeNS('Password','',entiteTemp.strPassword );
    xmlEntite.SetAttributeNS('NE','',entiteTemp.booNeverExpire );
    xmlEntite.SetAttributeNS('LoginScript','',entiteTemp.strLoginScript );
    for j := 0 to entiteTemp.listeSite.Count -1 do
    begin

// ****************************************************************************
// * Enregistrement de chaques site d'une  entité                             *
// ****************************************************************************

      sitetemp := tSite (entiteTemp.listeSite[j]);
      xmlSite := xmlEntite.AddChild('Site');
      xmlsite.SetAttributeNS('Nom','',sitetemp.strNom );
      xmlsite.SetAttributeNS('OU','',sitetemp.strOU );
      xmlsite.SetAttributeNS('Ville','',sitetemp.strVille );
      xmlsite.SetAttributeNS('SRV','',siteTemp.strSRV );
      xmlsite.SetAttributeNS('SRVPartage','',siteTemp.strSRVPartage );
      xmlsite.SetAttributeNS('SRVRepertoire','',siteTemp.strSRVRepertoire );
      xmlsite.SetAttributeNS('SRVLecteur','',siteTemp.strSRVLecteur );

// ****************************************************************************
// * Enregistrement des groupes par defaut                                    *
// ****************************************************************************

      for k := 0 to sitetemp.lstGrpGroupe.Count -1 do
      begin
        xmlGrpGroupe := xmlsite.AddChild('GroupeGroupe');
        xmlGrpGroupe.SetAttributeNS('Nom','',sitetemp.lstGrpGroupe[k]);
      end;
      for k := 0 to siteTemp.lstGrpExterieur.Count -1 do
      begin
        xmlgrpExterieur := xmlSite.AddChild('GroupeExterieur');
        xmlgrpExterieur.SetAttributeNS('Nom','',siteTemp.lstGrpExterieur[k]);
      end;

// ****************************************************************************
// * Enregistrement des banques par defaut                                    *
// ****************************************************************************

      for k := 0 to sitetemp.lstbanqueGroupe.Count -1 do
      begin
        xmlGrpGroupe := xmlsite.AddChild('BanqueGroupe');
        xmlGrpGroupe.SetAttributeNS('ID','',tbanque(sitetemp.lstbanqueGroupe.Items[k]).intId );
      end;
      for k := 0 to siteTemp.lstBanqueExterieur.Count -1 do
      begin
        xmlgrpExterieur := xmlSite.AddChild('BanqueExterieur');
        xmlgrpExterieur.SetAttributeNS('ID','',tbanque(sitetemp.lstBanqueExterieur.Items[k]).intId);
      end;
    end;
  end;
  fichierSites.SaveToFile(FBChangeFileExt(Application.ExeName ,'.fbsite') );
  fichierSites.Active := false;
end;

procedure TfrmSites.chargerFichier ;

//  ___________________________________________________________________________
// | procedure TfrmSites.chargerFichier                                        |
// | _________________________________________________________________________ |
// || Permet de charger le fichier des entités                                 ||
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
  boook : boolean;
  strNomfichierEntite : string;
begin

// ****************************************************************************
// * Effacement des entités                                                   *
// ****************************************************************************

  strNomfichierEntite := FBChangeFileExt(Application.ExeName ,'.fbsite');
  while lstEntite.Count > 0 do
  begin
    tentite(lstEntite.items.Objects[0]).Free;
    lstEntite.Items.Delete(0);
  end;
  lstsite.Clear ;
  viderDetailEntite;
  viderDetailSite ;
  activerDetailSite (false);
  fichierSites.Active := false;
  fichierSites.xml.Clear ;
  fichierSites.FileName :='';

// ****************************************************************************
// * Test l'existance du fichier des entités                                  *
// ****************************************************************************

  if not FBFileExists(strNomfichierEntite) then
  begin
    frmMaitre.ecritlog('Le fichier ' + FBExtractFileName (strNomfichierEntite) +' n''existe pas');
    frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierEntite) +' n''existe pas');
    exit;
  end;

// ****************************************************************************
// * Test l'intégrité du fichier des entités                                  *
// ****************************************************************************

  try
    fichierSites.LoadFromFile(strNomfichierEntite );
  except
    on e : exception do
    begin
      frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierEntite) +#10#13+
                              '      n''est pas un fichier FBSITE valide' + #10#13 + e.Message);
      fichierSites.Active := false;
      frmMaitre.ecritlog('/!\ Le fichier ' + FBExtractFileName (strNomfichierEntite) + ' n''est pas un fichier FBSITE valide');
      exit;
    end;
  end;

// ****************************************************************************
// * Test la version du fichier des entités                                   *
// ****************************************************************************

  boook := fichierSites.DocumentElement.HasAttribute('Version');
  if boook then boook := fichierSites.DocumentElement.GetAttributeNS('Version','') = versionFichierUtilisateur;
  if not boook then
  begin
    frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierEntite) +#10#13+
                            '      n''est pas à la bonne version');
    fichierSites.Active := false;
    frmMaitre.ecritlog('/!\ Le fichier ' + FBExtractFileName (strNomfichierEntite) + ' n''est pas à la bonne version');
    fichierSites.Active := false;
    exit;
  end;

// ****************************************************************************
// * Chargement des entités                                                   *
// ****************************************************************************

  try
    for i := 0 to fichierSites.DocumentElement.ChildNodes.Count -1 do
    begin
      entite := tEntite.create ;
      xmlEntite := fichierSites.DocumentElement.ChildNodes[i];
      entite.strNom := xmlEntite.GetAttributeNS('Nom','');
      entite.strProfileTSE := xmlEntite.GetAttributeNS('ProfileTse','' );
      entite.strNomComplet := xmlEntite.GetAttributeNS('NomComplet','' );
      entite.strPassword := xmlEntite.GetAttributeNS('Password','' );
      entite.booNeverExpire:= xmlEntite.GetAttributeNS('NE','' );
      entite.strLoginScript := xmlEntite.GetAttributeNS('LoginScript','' );
      for j := 0 to xmlentite.ChildNodes.Count -1 do
      begin
        site := tSite.create ;
        xmlSite := xmlEntite.ChildNodes[j];
        site.strNom := xmlsite.GetAttributeNS('Nom','' );
        site.strOU := xmlsite.GetAttributeNS('OU','' );
        site.strVille := xmlsite.GetAttributeNS('Ville','' );
        site.strSRV := xmlsite.GetAttributeNS('SRV','' );
        site.strSRVPartage := xmlsite.GetAttributeNS('SRVPartage','' );
        site.strSRVRepertoire := xmlsite.GetAttributeNS('SRVRepertoire','' );
        site.strSRVLecteur:= xmlsite.GetAttributeNS('SRVLecteur','' );
        for k := 0 to xmlsite.ChildNodes.Count -1 do
        begin
          if xmlSite.ChildNodes[k].NodeName = 'GroupeGroupe' then
          begin
            xmlGrpGroupe := xmlSite.ChildNodes[k];
            site.lstGrpGroupe.Add(xmlGrpGroupe.GetAttributeNS('Nom','') );
          end else
          if xmlSite.ChildNodes[k].NodeName = 'GroupeExterieur' then
          begin
            xmlgrpExterieur := xmlSite.ChildNodes[k];
            site.lstGrpExterieur.Add(xmlgrpExterieur.GetAttributeNS('Nom','') );
          end else
          if xmlSite.ChildNodes[k].NodeName = 'BanqueGroupe' then
          begin
            xmlgrpExterieur := xmlSite.ChildNodes[k];
            if frmBanque.findBanques(xmlgrpExterieur.GetAttributeNS('ID',''), listeBanques )then
              site.lstbanqueGroupe.Add(listebanques[0]);
          end else
          if xmlSite.ChildNodes[k].NodeName = 'BanqueExterieur' then
          begin
            xmlgrpExterieur := xmlSite.ChildNodes[k];
            if frmBanque.findBanques(xmlgrpExterieur.GetAttributeNS('ID',''), listeBanques )then
              site.lstBanqueExterieur.Add(listebanques[0]);
          end;
        end;
        entite.listeSite.Add(site);
      end;
      lstEntite.AddItem(entite.strNom ,entite);
    end;
    frmMaitre.ecritlog('Chargement du fichier ' + FBExtractFileName (strNomfichierEntite) + ' OK');
    frmInfo.AfficheMessage('  Chargement du fichier ' + FBExtractFileName (strNomfichierEntite) + ' OK');
  except
    frmInfo.AfficheMessage('  /!\ Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierEntite) +
                   #10#13 + '      ce n''est pas un fichier fbsite valide');
    entite.free;
    frmMaitre.ecritlog('/!\ Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierEntite) );
  end;
  fichierSites.Active := false;
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Public                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmSites.supprimerIDBanque(id : integer);

//  ___________________________________________________________________________
// | procedure TfrmSites.supprimerIDBanque                                     |
// | _________________________________________________________________________ |
// || Permet de supprimer toutes Références à une banque                      ||
// ||_________________________________________________________________________||
// || Entrées | ID : integer                                                  ||
// ||         |   N° de la banque                                             ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  entiteTemp : tEntite ;
  sitetemp : tsite;
  i,j,k : integer;
begin

// ****************************************************************************
// * On scan chaque entite                                                    *
// ****************************************************************************

  for i := 0 to lstEntite.Count-1 do
  begin
    entiteTemp := tentite(lstEntite.Items.Objects[i]);

// ****************************************************************************
// * On scan chaque site                                                      *
// ****************************************************************************

    for j:= 0 to entitetemp.listeSite.Count -1 do
    begin
      sitetemp := tsite (entiteTemp.listeSite[j]);

// ****************************************************************************
// * Si on trouve la banque on la supprime                                    *
// ****************************************************************************

      for k := sitetemp.lstbanqueGroupe.Count -1 downto 0 do
        if tbanque(sitetemp.lstbanqueGroupe[k]).intId = id then
           sitetemp.lstbanqueGroupe.Delete(k);
      for k := sitetemp.lstBanqueExterieur.Count -1 downto 0 do
        if tbanque(sitetemp.lstBanqueExterieur[k]).intId = id then
           sitetemp.lstBanqueExterieur.Delete(k);
    end;
  end;

// ****************************************************************************
// * On Enregistre le fichier des entites                                     *
// ****************************************************************************

  enregistrerFichier ;
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Evenementiel                                                       #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmSites.cmdAjouterEntiteClick( Sender: TObject );

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdAjouterEntiteClick                                 |
// | _________________________________________________________________________ |
// || Permet d'ajouter une nouvelle entité                                    ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstEntite.ItemIndex := -1;
  lstSite.Clear ;
  viderDetailEntite;
  viderDetailSite;
  activerDetailsite(false);
end;

procedure TfrmSites.FormCreate(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.FormCreate                                            |
// | _________________________________________________________________________ |
// || Permet lors de la création de la form de charger les entités            ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  frmInfo.AfficheMessage('Chargement liste des Entités');
  viderDetailEntite;
  viderDetailSite ;
  activerSite(false);
  listeBanques := TObjectList.Create (false);
  chargerFichier;
  frmInfo.fermer;
end;

procedure TfrmSites.FormDestroy(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.FormDestroy                                           |
// | _________________________________________________________________________ |
// || Libere les objets a la fermeture de la form                             ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
   listeBanques.Free ;
  while lstEntite.Count > 0 do
  begin
    tentite(lstEntite.items.Objects[0]).Free;
    lstEntite.Items.Delete(0);
  end;
end;

procedure TfrmSites.cmdSupprimerEntiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdSupprimerEntiteClick                               |
// | _________________________________________________________________________ |
// || Permet de supprimer une entité                                          ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if lstEntite.ItemIndex =-1 then exit;
  if MessageDlg('Etes vous sur de vouloir supprimer l''entité ' + tentite(lstEntite.items.Objects[lstEntite.ItemIndex]).strnom,
     mtConfirmation ,[mbyes,mbno],0) = mryes then
  begin
    tentite(lstEntite.items.Objects[lstEntite.ItemIndex]).Free;
    lstEntite.DeleteSelected ;
    lstSite.Clear ;
    viderDetailSite ;
    viderDetailEntite ;
    activerDetailsite(false);
    enregistrerFichier ;
  end;
end;

procedure TfrmSites.cmdEnregistrerEntiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdEnregistrerEntiteClick                             |
// | _________________________________________________________________________ |
// || Permet d'enregistrer les modifications de l'netité                      ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if edtNomEntite.Text = '' then exit;

// ****************************************************************************
// * Test doublon de nom                                                      *
// ****************************************************************************

  if (lstEntite.ItemIndex =-1)and(lstEntite.Items.IndexOf(edtNomEntite.Text)<>-1) then
  begin
    MessageDlg('L''entité que vous voulez ajouter existe déjà',mtError ,[mbok],0);
    exit;
  end;

// ****************************************************************************
// * C'est une nouvelle entité on l'ajoute                                    *
// ****************************************************************************

  if (lstEntite.ItemIndex =-1) then
  begin
    entite :=tentite.create;
    entite.strNom := edtNomEntite.Text ;
    entite.strProfileTSE := edtProfileTse.Text ;
    entite.strNomComplet := edtNomComplet.Text ;
    entite.strPassword := edtPassword.Text  ;
    entite.strLoginScript := edtLoginScript.Text ;
    entite.booNeverExpire := cbxNeverExpire.Checked ;
    lstEntite.AddItem(edtNomEntite.Text,entite);
  end else

// ****************************************************************************
// * C'est une modification d'entité on l'enregistre                          *
// ****************************************************************************

    if (lstEntite.Items.IndexOf(edtNomEntite.Text)=lstEntite.ItemIndex) or (lstEntite.Items.IndexOf(edtNomEntite.Text)=-1) then
    begin
      lstEntite.Items[lstEntite.ItemIndex] :=  edtNomEntite.Text;
      with tentite(lstEntite.items.Objects[lstEntite.ItemIndex]) do
      begin
        strNom := edtNomEntite.Text;
        strProfileTSE := edtProfileTse.Text ;
        strNomComplet := edtNomComplet.Text ;
        strPassword := edtPassword.Text  ;
        strLoginScript := edtLoginScript.Text ;
        booNeverExpire := cbxNeverExpire.Checked ;
      end;
    end else
    begin
      MessageDlg('Le nom fourni existe déjà' , mterror ,[mbok],0);
      exit;
    end;
  enregistrerFichier ;
end;

procedure TfrmSites.lstEntiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.lstEntiteClick                                        |
// | _________________________________________________________________________ |
// || permet l'affichage du détail des entitès                                ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var i : integer;
begin
  if lstEntite.ItemIndex > -1 then
  begin
    with tentite(lstEntite.items.Objects[lstEntite.ItemIndex]) do
    begin
      edtNomEntite.Text := strnom ;
      edtProfileTse.Text := strProfileTSE ;
      edtNomComplet.Text := strNomComplet ;
      edtPassword.Text := strPassword ;
      edtLoginScript.Text := strLoginScript ;
      cbxNeverExpire.Checked := booNeverExpire ;
    end;
    lstSite.Clear ;
    viderDetailSite ;
    activerSite(true);
    for i := 0 to tentite(lstEntite.items.Objects[lstEntite.ItemIndex]).listeSite.Count -1 do
    with tsite(tentite(lstEntite.items.Objects[lstEntite.ItemIndex]).listeSite[i]) do
    begin
      lstsite.AddItem(strNom ,tsite(tentite(lstEntite.items.Objects[lstEntite.ItemIndex]).listeSite[i]));
    end;
  end;
end;

procedure TfrmSites.cmdAjouterSiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdAjouterSiteClick                                   |
// | _________________________________________________________________________ |
// || Permet d'ajouter un nouveau site                                        ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstSite.ItemIndex := -1;
  viderDetailSite;
end;

procedure TfrmSites.cmdEnregistrerSiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdEnregistrerSiteClick                               |
// | _________________________________________________________________________ |
// || Permet d'enregistre un site                                             ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  if (edtNomSite.Text ='') or (lstEntite.ItemIndex =-1) then exit;

// ****************************************************************************
// * Test doublon de nom                                                      *
// ****************************************************************************

  if (lstSite.ItemIndex =-1)and(lstSite.Items.IndexOf(edtNomSite.Text)<>-1) then
  begin
    MessageDlg('Le site que vous voulez ajouter existe déjà',mtError ,[mbok],0);
    exit;
  end;

// ****************************************************************************
// * c'est un nouveau site on l'ajoute                                        *
// ****************************************************************************

  if (lstSite.ItemIndex =-1) then
  begin
    site := tsite.create ;
    site.strNom := edtNomSite.Text ;
    site.strOU  := edtOUSite.Text;
    site.strVille := edtCity.text;
    site.strSRV := edtSrvP.Text;
    site.strSRVPartage := edtSRVPartage.Text;
    site.strSRVRepertoire := edtsrvRepertoire.Text ;
    site.strSRVLecteur := edtLecteurP.text;
    for i := 0 to lstGroupeGroupe.Count -1 do
      site.lstGrpGroupe.Add(lstGroupeGroupe.Items[i]);
    for i := 0 to lstGroupeExterieur.Count -1 do
      site.lstGrpExterieur.Add(lstGroupeExterieur.Items[i]);
    for i := 0 to lstBanqueExchangeGroupe.Count -1 do
      site.lstbanqueGroupe.Add(lstBanqueExchangeGroupe.Items.Objects[i]);
    for i := 0 to lstBanqueExchangeExterieur.Count -1 do
      site.lstBanqueExterieur.Add(lstBanqueExchangeExterieur.Items.Objects[i]);
    tEntite(lstEntite.items.Objects[lstEntite.ItemIndex]).listeSite.Add(site);
    lstsite.AddItem(edtNomSite.Text,site);
  end else

// ****************************************************************************
// * C'est une modification on l'enregistre                                   *
// ****************************************************************************

  if (lstSite.Items.IndexOf(edtNomSite.Text)=lstSite.ItemIndex) or (lstSite.Items.IndexOf(edtNomSite.Text)=-1) then
  begin
    lstSite.Items[lstSite.ItemIndex] :=  edtNomSite.Text;
    with tsite( lstSite.items.Objects[lstSite.ItemIndex]) do
    begin
      strNom := edtNomSite.Text;
      strOU  := edtOUSite.Text;
      strVille := edtCity.text;
      strSRV := edtSrvP.Text;
      strSRVPartage := edtSRVPartage.Text;
      strSRVRepertoire := edtsrvRepertoire.Text ;
      strSRVLecteur := edtLecteurP.text;
      lstGrpGroupe.Clear ;
      for i := 0 to lstGroupeGroupe.Count -1 do
        lstGrpGroupe.Add(lstGroupeGroupe.Items[i]);
      lstGrpExterieur.Clear ;
      for i := 0 to lstGroupeExterieur.Count -1 do
        lstGrpExterieur.Add(lstGroupeExterieur.Items[i]);
      lstbanqueGroupe.Clear ;
      for i := 0 to lstBanqueExchangeGroupe.Count -1 do
        lstbanqueGroupe.Add(lstBanqueExchangeGroupe.Items.Objects[i]);
      lstBanqueExterieur.Clear ;
      for i := 0 to lstBanqueExchangeExterieur.Count -1 do
        lstBanqueExterieur.Add(lstBanqueExchangeExterieur.Items.Objects[i]);
    end;
  end else
  begin
    MessageDlg('Le nom du site fourni existe déjà' , mterror ,[mbok],0);
    exit;
  end;
  enregistrerFichier ;
end;

procedure TfrmSites.lstSiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.lstSiteClick                                          |
// | _________________________________________________________________________ |
// || Permet d'afficher le détail d'un site                                   ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var i : integer;
begin
  if lstSite.ItemIndex >-1 then
  begin
    with tsite(tentite(lstEntite.items.Objects[lstEntite.ItemIndex]).listeSite[lstSite.ItemIndex]) do
    begin
      edtnomSite.Text := strNom  ;
      edtOUSite.Text := strOU ;
      edtCity.Text := strVille ;
      edtSrvP.Text := strSRV ;
      edtSRVPartage.Text := strSRVPartage ;
      edtsrvRepertoire.Text := strSRVRepertoire ;
      edtLecteurP.Text := strSRVLecteur ;
      lstGroupeGroupe.Clear ;
      lstGroupeExterieur.Clear ;
      for i :=0 to lstGrpGroupe.Count -1 do
        lstGroupeGroupe.items.Add(lstGrpGroupe[i]);
      for i :=0 to lstGrpExterieur.Count -1 do
        lstGroupeExterieur.Items.Add(lstGrpExterieur[i]);
      lstBanqueExchangeGroupe.Clear ;
      lstBanqueExchangeExterieur.Clear ;
      for i :=0 to lstbanqueGroupe.Count -1 do
        lstBanqueExchangeGroupe.AddItem(tbanque(lstbanqueGroupe.Items[i]).strNom , lstbanqueGroupe.Items[i]);
      for i :=0 to lstBanqueExterieur.Count -1 do
        lstBanqueExchangeExterieur.AddItem(tbanque(lstBanqueExterieur.Items[i]).strNom , lstBanqueExterieur.Items[i]);
    end;
  end;
end;

procedure TfrmSites.cmdAouterGrproupeClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdAouterGrproupeClick                                |
// | _________________________________________________________________________ |
// || Permet d'ajouter un groupe                                              ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
  lstgrp : tstringlist;
begin
  lstgrp := frmGroupe.GetGroupe;
  for i := 0 to lstgrp.Count -1 do
  begin
    if lstGroupeGroupe.Items.IndexOf(lstgrp.Strings[i]) = -1 then
      lstGroupeGroupe.Items.Add(lstgrp.Strings[i]);
  end;
  lstgrp.Free ;
end;

procedure TfrmSites.cmdSupprimerGRPGroupeClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdSupprimerGRPGroupeClick                            |
// | _________________________________________________________________________ |
// || permet de supprimer un groupe                                           ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstGroupeGroupe.DeleteSelected ;
end;

procedure TfrmSites.cmdAjouterGRPExterieurClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdAjouterGRPExterieurClick                           |
// | _________________________________________________________________________ |
// || permet d'ajouter un groupe                                              ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
  lstgrp : tstringlist;
begin
  lstgrp := frmGroupe.GetGroupe;
  for i := 0 to lstgrp.Count -1 do
  begin
    if lstGroupeExterieur.Items.IndexOf(lstgrp.Strings[i]) = -1 then
      lstGroupeExterieur.Items.Add(lstgrp.Strings[i]);
  end;
  lstgrp.Free;
end;

procedure TfrmSites.cmdSupprimerGRPExterieurClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdSupprimerGRPExterieurClick                         |
// | _________________________________________________________________________ |
// || permet de supprimer un groupe                                           ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstGroupeExterieur.DeleteSelected ;
end;

procedure TfrmSites.cmdSupprimerSiteClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdSupprimerSiteClick                                 |
// | _________________________________________________________________________ |
// || Permet de supprimer un site                                             ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  sitetemp : tsite;
  entiteTemp : tentite;
  i : integer;
begin
  if lstSite.ItemIndex =-1 then exit;
  sitetemp := tsite(lstsite.Items.Objects[lstSite.ItemIndex]);
  if MessageDlg('Etes vous sur de vouloir supprimer le site ' +   sitetemp.strnom,
     mtConfirmation ,[mbyes,mbno],0) = mryes then
  begin
    for i := tentite(lstentite.Items.Objects[lstEntite.ItemIndex ]).listeSite.Count-1 downto 0 do
      if tsite(tentite(lstentite.Items.Objects[lstEntite.ItemIndex ]).listeSite.Items[i]).strNom = sitetemp.strNom then
      begin
        tentite(lstentite.Items.Objects[lstentite.ItemIndex ]).listeSite.Delete (i);
      end;
    lstEntiteClick(nil);
    enregistrerFichier ;
  end;
end;

procedure TfrmSites.cmdAjouteBanqueGroupeClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdAjouteBanqueGroupeClick                            |
// | _________________________________________________________________________ |
// || Permet d'ajouter une banque                                             ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  if lstSite.ItemIndex =-1 then exit;
  if frmBanque.getBanques(listeBanques) then
  for i := 0 to listeBanques.Count -1 do
  begin
    if lstBanqueExchangeGroupe.Items.IndexOf(tbanque(listeBanques.Items[i]).strNom)= -1 then
    begin
      lstBanqueExchangeGroupe.AddItem(tbanque(listeBanques.Items[i]).strNom,listeBanques.Items[i]);
    end;
  end;
end;

procedure TfrmSites.cmdSupprimerBanqueGroupeClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdSupprimerBanqueGroupeClick                         |
// | _________________________________________________________________________ |
// || permet de supprimer une banque                                          ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstBanqueExchangeGroupe.DeleteSelected;
end;

procedure TfrmSites.cmdAjouterBanqueExterieurClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdAjouterBanqueExterieurClick                        |
// | _________________________________________________________________________ |
// || permet d'ajouter une banque                                             ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  if lstSite.ItemIndex =-1 then exit;
  if frmBanque.getBanques(listeBanques) then
  for i := 0 to listeBanques.Count -1 do
  begin
    if lstBanqueExchangeExterieur.Items.IndexOf(tbanque(listeBanques.Items[i]).strNom)= -1 then
    begin
      lstBanqueExchangeExterieur.AddItem(tbanque(listeBanques.Items[i]).strNom,listeBanques.Items[i]);
    end;
  end;
end;

procedure TfrmSites.cmdSupprimerbanqueExterieurClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmSites.cmdSupprimerbanqueExterieurClick                      |
// | _________________________________________________________________________ |
// || permet de supprimer une banque                                          ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstBanqueExchangeExterieur.DeleteSelected ;
end;

end.
