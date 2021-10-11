unit ufrmBanques;

// |===========================================================================|
// | unit ufrmBanques                                                          |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité permettant la gestion des banques de messagerie exchange pour le    |
// | programme fbcreateuser                                                    |
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
  Dialogs, StdCtrls,ubanques,Contnrs,ufbfichier, xmldom, XMLIntf, msxmldom,
  XMLDoc, ExtCtrls;

type
  TfrmBanque = class(TForm)
    lstBanques: TListBox;
    fichierBanques: TXMLDocument;
    pnlGetBanques: TPanel;
    cmdAjouter: TButton;
    pnlEditBanques: TPanel;
    edtNomBanques: TEdit;
    Label1: TLabel;
    edtCheminLdap: TEdit;
    Label2: TLabel;
    cmdAjouterBanque: TButton;
    cmdEnregistrer: TButton;
    cmdSupprimer: TButton;
    procedure lstBanquesClick(Sender: TObject);
    procedure cmdSupprimerClick(Sender: TObject);
    procedure cmdEnregistrerClick(Sender: TObject);
    procedure cmdAjouterBanqueClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    banque : tbanque;
    xmlEntite : IXMLNode;
    procedure chargerFichierBanque;
    procedure SauvegarderFichierBanque;
    function getNewID : integer;
  public
    { Déclarations publiques }
    class function getBanques(var listeBanque : TObjectlist ) : boolean;
    class function findBanques(ID : integer ;var listeBanque : TObjectlist ) : boolean;
    class procedure EditBanques;
  end;

var
  frmBanque: TfrmBanque;

implementation

uses UfrmMaitre,ufrmInfo,ufrmSites;

{$R *.dfm}
const
  versionFichierBanque ='1.0.0.0';

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Privée                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

function TfrmBanque.getNewID : integer;

//  ___________________________________________________________________________
// | function TfrmBanque.getNewID                                              |
// | _________________________________________________________________________ |
// || Permet d'avoir un nouvelle ID                                           ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties | result : integer                                              ||
// ||         |   nouvelle ID                                                 ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  result := 0;
  for i := 0 to lstBanques.Count -1 do
    if tbanque(lstBanques.Items.Objects[i]).intId > result then result := tbanque(lstBanques.Items.Objects[i]).intId;
  result := result +1;
end;

procedure TfrmBanque.SauvegarderFichierBanque ;

//  ___________________________________________________________________________
// | procedure TfrmBanque.SauvegarderFichierBanque                             |
// | _________________________________________________________________________ |
// || Permet d'enregistrer les modifications dans le fichiers des banques     ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  strNomfichierBanque : string;
  i : integer;
  tempBanque : tbanque;
begin
  strNomfichierBanque := FBChangeFileExt(Application.ExeName ,'.fbbanque');
  fichierBanques.Active := false;
  fichierBanques.xml.Clear ;
  fichierBanques.FileName :='';
  fichierBanques.Active := true;

// ****************************************************************************
// * Ajout des infos de version XML                                           *
// ****************************************************************************

  fichierBanques.Version :='1.0';
  fichierBanques.StandAlone :='no';
  fichierBanques.encoding := 'UTF-8';
  fichierBanques.DocumentElement := fichierBanques.CreateElement('Banques','');
  fichierBanques.DocumentElement.SetAttributeNS('Version','',versionFichierBanque);
  for i := 0 to lstBanques.Count -1 do
  begin

// ****************************************************************************
// * Enregistrement de chaque entité                                          *
// ****************************************************************************

    tempBanque  := tBanque(lstBanques.items.Objects[i]);
    xmlEntite := fichierBanques.DocumentElement.AddChild('Banque');
    xmlEntite.SetAttributeNS('ID','',tempBanque.intId );
    xmlEntite.SetAttributeNS('Nom','',tempBanque.strNom );
    xmlEntite.SetAttributeNS('Chemin','',tempBanque.strChemin );
  end;
  fichierBanques.SaveToFile(strNomfichierBanque );
  fichierBanques.Active := false;
end;

procedure TfrmBanque.chargerFichierBanque;

//  ___________________________________________________________________________
// | procedure TfrmBanque.chargerFichierBanque                                 |
// | _________________________________________________________________________ |
// || Permet de charger la liste des banques                                  ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
  boook : boolean;
  strNomfichierBanque : string;
begin
  strNomfichierBanque := FBChangeFileExt(Application.ExeName ,'.fbbanque');

// ****************************************************************************
// * On vide la liste existante                                               *
// ****************************************************************************

  while lstBanques.Count > 0 do
  begin
    tBanque(lstBanques.items.Objects[0]).Free;
    lstBanques.Items.Delete(0);
  end;

// ****************************************************************************
// * Test l'existance du fichier                                              *
// ****************************************************************************

  if not FBFileExists(strNomfichierBanque) then
  begin
    frmMaitre.ecritlog('Le fichier ' + FBExtractFileName (strNomfichierBanque) +' n''existe pas');
    frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierBanque) +' n''existe pas');
    exit;
  end;

// ****************************************************************************
// * Test la validite du fichier                                              *
// ****************************************************************************

  fichierBanques.Active := false;
  fichierBanques.xml.Clear ;
  fichierBanques.FileName :='';
  try
    fichierBanques.LoadFromFile(strNomfichierBanque );
  except
    on e : exception do
    begin
      fichierBanques.Active := false;
      frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierBanque) +#10#13+
                              '      n''est pas un fichier FBBANQUE valide' + #10#13 +
                              '      '+e.Message);
      frmMaitre.ecritlog('/!\ Le fichier ' + FBExtractFileName (strNomfichierBanque) + ' n''est pas un fichier FBBANQUE valide');
      exit;
    end;
  end;

// ****************************************************************************
// * Test la version du fichier                                               *
// ****************************************************************************

  boook := fichierBanques.DocumentElement.HasAttribute('Version');
  if boook then boook := fichierBanques.DocumentElement.GetAttributeNS('Version','') = versionFichierBanque;
  if not boook then
  begin
    frmInfo.AfficheMessage('  /!\ Le fichier ' + FBExtractFileName (strNomfichierBanque) +#10#13+
                            '      n''est pas à la bonne version');
    fichierBanques.Active := false;
    frmMaitre.ecritlog('/!\ Le fichier ' + FBExtractFileName (strNomfichierBanque) + ' n''est pas à la bonne version');
    fichierBanques.Active := false;
    exit;
  end;

// ****************************************************************************
// * Chargement de la liste des banques                                       *
// ****************************************************************************

  try
    for i := 0 to fichierBanques.DocumentElement.ChildNodes.Count -1 do
    begin
      banque := tBanque.create ;
      xmlEntite := fichierBanques.DocumentElement.ChildNodes[i];
      banque.intId := xmlEntite.GetAttributeNS('ID','');
      banque.strNom := xmlEntite.GetAttributeNS('Nom','' );
      banque.strChemin := xmlEntite.GetAttributeNS('Chemin','' );
      lstBanques.AddItem (banque.strNom , banque);
    end;
    frmMaitre.ecritlog('Chargement du fichier ' + FBExtractFileName (strNomfichierBanque) + ' OK');
    frmInfo.AfficheMessage('  Chargement du fichier ' + FBExtractFileName (strNomfichierBanque) + ' OK');
  except
    MessageDlg('Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierBanque) +
                #10#13 + 'ce n''est pas un fichier fbbanque valide' ,mtError ,[mbok],0);
    frmInfo.AfficheMessage('  /!\ Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierBanque) +
                             #10#13 + '      ce n''est pas un fichier fbbanque valide');
    banque.free;
    frmMaitre.ecritlog('/!\ Erreur lors du traitement du fichier ' + FBExtractFileName (strNomfichierBanque) );
  end;
  fichierBanques.Active := false;
end ;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Public                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

class function TfrmBanque.getBanques ( var listeBanque : TObjectlist ) : boolean;

//  ___________________________________________________________________________
// | class function TfrmBanque.getBanques                                      |
// | _________________________________________________________________________ |
// || Permet d'avoir une ou plusieurs banques en les selectionant             ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties | var listeBanque : TObjectlist                                 ||
// ||         |   liste d'objet rempli avec la selection                      ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  listeBanque.Clear ;
  if not Assigned( frmBanque ) then exit;
  result := false;
  frmBanque.pnlEditbanques.Visible := false;
  frmBanque.pnlGetBanques.Visible := true;
  if frmBanque.ShowModal = mrok then
  begin
    result := frmBanque.lstBanques.selcount> 0;
    if result then
    for i := 0 to frmBanque.lstBanques.count -1 do
      if frmBanque.lstBanques.Selected[i] then
      begin
        listeBanque.Add(frmBanque.lstBanques.Items.Objects[i]);
      end;
  end;
end;

class function TfrmBanque.findBanques(ID : integer ;var listeBanque : TObjectlist ) : boolean;

//  ___________________________________________________________________________
// | class function TfrmBanque.findBanques                                     |
// | _________________________________________________________________________ |
// || Permet d'avoir une banque via son ID                                    ||
// ||_________________________________________________________________________||
// || Entrées | ID : integer                                                  ||
// ||         |   numero de la banque                                         ||
// ||_________|_______________________________________________________________||
// || Sorties | var listeBanque : TObjectlist                                 ||
// ||         |   liste d'objet rempli avec la selection                      ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var i : integer;
begin
  listeBanque.Clear ;
  result := false;
  if not Assigned( frmBanque ) then exit;
  for i := 0 to frmBanque.lstBanques.Count -1 do
    if tbanque(frmBanque.lstBanques.Items.Objects[i]).intId = id then
    begin
      result := true;
      listeBanque.Add(frmBanque.lstBanques.Items.Objects[i]);
    end;
end;

class procedure TfrmBanque.EditBanques;

//  ___________________________________________________________________________
// | class procedure TfrmBanque.EditBanques                                    |
// | _________________________________________________________________________ |
// || Pemet d'éditer la liste des banques                                     ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if not Assigned( frmBanque ) then exit;
  frmBanque.pnlEditbanques.Visible := true;
  frmBanque.pnlGetBanques.Visible := false;
  frmBanque.ShowModal
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Evenementiel                                                       #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmBanque.FormCreate(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmBanque.FormCreate                                           |
// | _________________________________________________________________________ |
// || Charge la liste des banques au demarrage                                ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  frmInfo.AfficheMessage('Chargement liste des banques exchange');
  chargerFichierBanque;
end;

procedure TfrmBanque.FormDestroy(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmBanque.FormDestroy                                          |
// | _________________________________________________________________________ |
// || permet de liberer la mémoire lors de l'arret du programme               ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  while lstBanques.Count > 0 do
  begin
    tBanque(lstBanques.items.Objects[0]).Free;
    lstBanques.Items.Delete(0);
  end;
end;

procedure TfrmBanque.cmdAjouterBanqueClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmBanque.cmdAjouterBanqueClick                                |
// | _________________________________________________________________________ |
// || Permet d'ajouter une nouvelle banque                                    ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  edtNomBanques.Text := trim(edtNomBanques.Text);
  edtCheminLdap.Text := trim( edtCheminLdap.Text);
  if (edtNomBanques.Text ='')or (edtCheminLdap.Text ='') then exit;

// ****************************************************************************
// * Evite les doublon de nom de banques                                      *
// ****************************************************************************

  if lstBanques.Items.IndexOf(edtNomBanques.Text )<> -1 then
  begin
    MessageDlg('Le nom de la banque existe déjà',mterror,[mbok],0);
    exit;
  end;

// ****************************************************************************
// * Ajout de la nouvelle banque                                              *
// ****************************************************************************

  banque := tbanque.create;
  banque.intId := getNewID;
  banque.strNom := edtnomBanques.Text;
  banque.strChemin := edtCheminLdap.Text ;
  lstBanques.Items.AddObject(banque.strnom,banque);
  SauvegarderFichierBanque ;
end;

procedure TfrmBanque.cmdEnregistrerClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmBanque.cmdEnregistrerClick                                  |
// | _________________________________________________________________________ |
// || Premet d'enregistrer les modifications faite à une banque               ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  tempBanque : tBanque ;
begin
  edtNomBanques.Text := trim(edtNomBanques.Text);
  edtCheminLdap.Text := trim( edtCheminLdap.Text);
  if (edtNomBanques.Text ='')or (edtCheminLdap.Text ='')or (lstBanques.ItemIndex =-1) then exit;

// ****************************************************************************
// * Evite les doublon de nom de banques                                      *
// ****************************************************************************

  if( lstBanques.Items.IndexOf(edtNomBanques.Text )<> -1)and(lstBanques.Items.IndexOf(edtNomBanques.Text )<>lstBanques.ItemIndex ) then
  begin
    MessageDlg('Le nom de la banque existe déjà',mterror,[mbok],0);
    exit;
  end;

// ****************************************************************************
// * Enregistrement des modifictions                                          *
// ****************************************************************************

  tempBanque := tbanque(lstBanques.Items.Objects[lstBanques.ItemIndex]);
  tempBanque.strnom := edtNomBanques.Text;
  tempBanque.strChemin := edtCheminLdap.Text;
  SauvegarderFichierBanque ;
end;

procedure TfrmBanque.cmdSupprimerClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmBanque.cmdSupprimerClick                                    |
// | _________________________________________________________________________ |
// || Permet de supprimer un banques                                          ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if (lstBanques.ItemIndex =-1) then exit;
  if MessageDlg('Etes vous sur de vouloir supprimer ' + edtNomBanques.Text ,mtConfirmation ,[mbyes,mbno],0)=mryes then
  begin
    frmSites.supprimerIDBanque(tBanque(lstBanques.items.Objects[lstBanques.itemindex]).intId);
    tBanque(lstBanques.items.Objects[lstBanques.itemindex]).Free;
    lstBanques.items.Delete(lstBanques.itemindex);
    SauvegarderFichierBanque ;
  end;
end;

procedure TfrmBanque.lstBanquesClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmBanque.lstBanquesClick                                      |
// | _________________________________________________________________________ |
// || Permet d'afficher le détail de la banque sélectionnée                   ||
// ||_________________________________________________________________________||
// || Entrées | Sender: TObject                                               ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  tempBanque : tBanque ;
begin
  if pnlEditBanques.visible then
  begin
    tempBanque := tbanque(lstBanques.Items.Objects[lstBanques.ItemIndex]);
    edtNomBanques.Text := tempBanque.strNom ;
    edtCheminLdap.Text := tempBanque.strChemin ;
  end;
end;

end.
