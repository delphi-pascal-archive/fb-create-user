unit ufrmUser;

// |===========================================================================|
// | unit ufrmUser                                                             |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité permettant la gestion des utilisateurs pour le programme            |
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
  Dialogs, StdCtrls, Buttons, ComCtrls,ufbadsi,activex,ComObj ,inifiles,
  ExtCtrls,DateUtils,ufbreseau,ufrmgroupe,mapi,registry,uuser,ufrmsites,uentite,ubanques;

type

  TfrmUser = class(TForm)
    GrpOptionel: TGroupBox;
    edtTel: TEdit;
    Label6: TLabel;
    edtFax: TEdit;
    Label8: TLabel;
    grpObligaatoire: TGroupBox;
    edtNom: TEdit;
    Label1: TLabel;
    edtPrenom: TEdit;
    Label2: TLabel;
    edtUID: TEdit;
    Label3: TLabel;
    edtLogin: TEdit;
    Label4: TLabel;
    cbxLocalisation: TComboBox;
    Label7: TLabel;
    Label9: TLabel;
    cbxApartenance: TComboBox;
    edtDepartement: TEdit;
    Label10: TLabel;
    lstGroupe: TListBox;
    Label11: TLabel;
    cmdAddGrp: TBitBtn;
    cmdSupprimer: TBitBtn;
    Label5: TLabel;
    edtMail: TEdit;
    grbAutogenere: TGroupBox;
    edtProfileTse: TEdit;
    Label14: TLabel;
    edtPassword: TEdit;
    Label12: TLabel;
    edtVille: TEdit;
    Label15: TLabel;
    edtNomComplet: TEdit;
    Label16: TLabel;
    edtDescription: TEdit;
    Label17: TLabel;
    edtSociete: TEdit;
    Label18: TLabel;
    grbDate: TGroupBox;
    dtDateDepart: TDateTimePicker;
    Label13: TLabel;
    cbxDuree: TCheckBox;
    edtCompagnie: TEdit;
    Label19: TLabel;
    edtLogonScript: TEdit;
    Label20: TLabel;
    cmdCreer: TBitBtn;
    cbxP: TCheckBox;
    cbxNeverExpire: TCheckBox;
    cmdAnnuler: TBitBtn;
    GroupBox1: TGroupBox;
    cbxBanque: TComboBox;
    procedure cmdAddGrpClick(Sender: TObject);
    procedure cmdCreerClick(Sender: TObject);
    procedure cmdSupprimerClick(Sender: TObject);
    procedure edtLoginChange(Sender: TObject);
    procedure cbxDureeClick(Sender: TObject);
    procedure edtChange(Sender: TObject);
    procedure cbxApartenanceClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Déclarations privées }
    strNomComplet : string;
    strProfileTse : string;
    strlogonscript : string;
  public
    { Déclarations publiques }
    class function editerUser (var user : tuser) : boolean;
  end;

var
  frmUser: TfrmUser;
implementation

{$R *.dfm}

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Public                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

class function TfrmUser.editerUser (var user : tuser) : boolean;

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet de modifier un user                                              ||
// ||_________________________________________________________________________||
// || Entrées | var user : tuser                                              ||
// ||         |   utilisateur à modifier                                      ||
// ||_________|_______________________________________________________________||
// || Sorties | result : boolean                                              ||
// ||         |   vrai si validation                                          ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  j,i : integer;
begin
  result := false;
  if not Assigned(frmUser ) then exit;
  with frmuser do
  begin

// ****************************************************************************
// * Réinitialise tous les champs avec les information du User                *
// ****************************************************************************

    FormCreate(nil);
    cbxApartenance.ItemIndex := cbxApartenance.Items.IndexOf(user.strAppartenance);
    cbxApartenanceClick(nil);
    Application.ProcessMessages ;
    cbxLocalisation.ItemIndex := cbxLocalisation.Items.IndexOf(user.strSIte);
    edtNom.Text := user.strNom ;
    edtPrenom.Text := user.strPrenom ;
    edtUID.Text := user.strUid;
    edtLogin.Text := user.strLogin ;
    edtMail.Text := user.strAddresseMail ;
    edtSociete.Text := user.strSociete ;
    dtDateDepart.DateTime := user.dtDateExpiration ;
    cbxDuree.Checked := (user.strSociete <>'' ) or (user.dtDateExpiration <>0 );
    edtChange(nil);
    Application.ProcessMessages ;
    if user.strNomComplet <> '' then edtNomComplet.Text := user.strNomComplet ;
    if user.strDescription <> '' then edtDescription.Text := user.strDescription;
    if user.strPassword <> '' then
    begin
      edtPassword.Text := user.strPassword;
      cbxNeverExpire.Checked := user.booPasswordNeverExpire ;
    end;
    if user.strProfileTse <> '' then edtProfileTse.Text := user.strProfileTse;
    if user.strVille <> '' then edtVille.Text := user.strVille ;
    if user.strCompagnie <> '' then edtCompagnie.Text := user.strCompagnie ;
    if user.strLogonScript <> '' then edtLogonScript.Text := user.strLogonScript ;
    if user.strNomComplet <>'' then
      cbxP.Checked := user.strRepertoirePerso <>'';
    edtTel.Text := user.strtel ;
    edtfax.Text := user.strFax ;
    edtDepartement.Text := user.strDepartement ;
    Application.ProcessMessages ;
    if user.listeGroupes.Count >0 then lstGroupe.Items := user.listeGroupes ;
    for i := 0 to cbxBanque.Items.Count -1 do
      if tbanque(cbxBanque.Items.Objects[i]).strChemin = user.strBanque then
        cbxBanque.ItemIndex := i;

// ****************************************************************************
// * Attente modifications                                                    *
// ****************************************************************************

    result := ShowModal = mrok;
    if not result then exit;

// ****************************************************************************
// * Modifications des informations du USer                                   *
// ****************************************************************************

    user.strNom := edtNom.Text;
    user.strPrenom := edtPrenom.Text;
    user.strUid := edtUID.Text;
    user.strLogin := edtLogin.Text;
    user.strAddresseMail := edtMail.Text;
    user.strAppartenance := cbxApartenance.Text ;
    user.strSIte := cbxLocalisation.text;
    user.strSociete := edtSociete.Text;
    if cbxDuree.Checked then
      user.dtDateExpiration := dtDateDepart.DateTime
    else
      user.dtDateExpiration := 0;
    user.strNomComplet := edtNomComplet.Text;
    user.strDescription := edtDescription.Text;
    user.strPassword := edtPassword.Text;
    user.booPasswordNeverExpire := cbxNeverExpire.Checked;
    user.strProfileTse := edtProfileTse.Text;
    user.strVille := edtVille.Text;
    user.strCompagnie := edtCompagnie.Text;
    user.strLogonScript := edtLogonScript.Text;
    if cbxp.Checked then
    begin
      user.strRepertoirePerso := '\\'+tsite(cbxLocalisation.Items.Objects[cbxLocalisation.ItemIndex]).strSRV +'\' +  edtLogin.Text +'$' ;
      user.strLecteur := tsite(cbxLocalisation.Items.Objects[cbxLocalisation.ItemIndex]).strSRVLecteur;
      user.strSRV := tsite(cbxLocalisation.Items.Objects[cbxLocalisation.ItemIndex]).strSRV;
      user.strSRVPartage := tsite(cbxLocalisation.Items.Objects[cbxLocalisation.ItemIndex]).strSRVPartage;
      user.strSRVRepertoire := tsite(cbxLocalisation.Items.Objects[cbxLocalisation.ItemIndex]).strSRVRepertoire;
    end else
    begin
      user.strRepertoirePerso := '';
      user.strLecteur := '';
      user.strSRV := '';
      user.strSRVPartage :='';
      user.strSRVRepertoire := '';
    end;
    user.strtel := edtTel.Text;
    user.strFax := edtfax.Text;
    user.strDepartement := edtDepartement.Text;
    user.listeGroupes.clear ;
    user.strOU := tsite(cbxLocalisation.Items.Objects[cbxLocalisation.ItemIndex]).strOU;
    for j := 0 to lstGroupe.Count -1 do
      user.listeGroupes.Add(lstGroupe.Items[j]);
    if cbxBanque.ItemIndex <> -1 then
      user.strBanque:= tbanque(cbxBanque.Items.Objects[cbxBanque.ItemIndex]).strChemin
    else
      user.strBanque:= '';
  end;
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Evenementiel                                                       #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

function AnsiFirstUpCase(const S: string): string;

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet de passer le premier caractère en majuscule                      ||
// ||_________________________________________________________________________||
// || Entrées | const S: string                                               ||
// ||         |   chaine à modifier                                           ||
// ||_________|_______________________________________________________________||
// || Sorties | result : string                                               ||
// ||         |   chaime modifée                                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  Len: Integer;
begin
  Len := Length(S);
  SetString(Result, PChar(S), Len);
  if Len > 1 then CharLowerBuff(Pointer(Result), Len);
  if Len > 0 then CharUpperBuff(Pointer(Result), 1);
end;

procedure TfrmUser.FormCreate(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmUser.FormCreate                                             |
// | _________________________________________________________________________ |
// || Permet d'initialiser la liste des entite a la création                  ||
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
  entitetemp : tEntite;
begin
  cbxApartenance.Items.Clear ;
  cbxApartenance.Text := '';
  caption := 'Edition Utilisateur';
  application.ProcessMessages ;
  dtDateDepart.Date := now;
  for i := 0 to frmSites.lstEntite.Count -1 do
  begin
    entitetemp := tentite( frmSites.lstEntite.Items.Objects[i]);
    if entitetemp.listeSite.Count > 0 then
    begin
      cbxApartenance.Items.AddObject(entitetemp.strNom ,frmSites.lstEntite.Items.Objects[i]);
    end;
  end;
  application.ProcessMessages ;
  cbxApartenanceClick(nil);
  cbxDureeClick(nil);
end;

procedure TfrmUser.cbxApartenanceClick(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmUser.cbxApartenanceClick                                    |
// | _________________________________________________________________________ |
// || Permet de pcharger la liste des site en fonction de l'appartenance      ||
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
  entiteTemp : tentite;
  sitetemp : tsite;
begin
  cbxLocalisation.Items.Clear ;
  if cbxApartenance.ItemIndex <> -1 then
  begin
    entiteTemp := tEntite (cbxApartenance.Items.Objects[cbxApartenance.ItemIndex]);
    for i := 0 to entiteTemp.listeSite.Count -1 do
    begin
      sitetemp := tsite(entiteTemp.listeSite[i]);
      cbxLocalisation.Items.AddObject (sitetemp.strNom ,tsite(entiteTemp.listeSite[i]));
    end;
    strNomComplet := entitetemp.strNomComplet ;
    strProfileTse := entiteTemp.strProfileTSE ;
    edtPassword.Text := entiteTemp.strPassword ;
    cbxNeverExpire.Checked := entiteTemp.booNeverExpire ;
    strlogonscript := entiteTemp.strLoginScript ;
    edtLogonScript.Text := strlogonscript ;
  end;
  cbxLocalisation.ItemIndex := 0;
  edtchange(nil);
  edtLoginChange(nil); 
end;

procedure TfrmUser.edtChange(Sender: TObject);

//  ___________________________________________________________________________
// | procedure TfrmUser.edtChange                                              |
// | _________________________________________________________________________ |
// || Permet de calculer les champs auto dès la modification d'un champ       ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  strtemp ,strT: string;
  sitetemp : tsite;
  entitetemp : tEntite ;
  i : integer;
begin
  if (cbxLocalisation.itemindex =-1) or (cbxApartenance.Itemindex =-1) then exit;
  sitetemp := tsite(cbxLocalisation.Items.Objects[cbxLocalisation.itemindex]);
  entitetemp := tEntite (cbxApartenance.Items.Objects[cbxApartenance.Itemindex]);

// ****************************************************************************
// * Modification du nom complet en fonction du model                         *
// ****************************************************************************

  strT := edtNom.Text;
  if pos ('%nom%',strnomcomplet)<> 0 then strT := lowercase (strT);
  if pos ('%Nom%',strnomcomplet)<> 0 then strT := AnsiFirstUpCase(strT);
  if pos ('%NOM%',strnomcomplet)<> 0 then strT := UpperCase(strT);
  strtemp := StringReplace(strnomcomplet,'%nom%',strT,[rfReplaceAll,rfIgnoreCase ]);
  strT := edtPrenom.Text ;
  if pos ('%prenom%',strnomcomplet)<> 0 then strT := lowercase (strT);
  if pos ('%Prenom%',strnomcomplet)<> 0 then strT := AnsiFirstUpCase(strT);
  if pos ('%PRENOM%',strnomcomplet)<> 0 then strT := UpperCase(strT);
  strtemp := StringReplace(strtemp,'%prenom%',strT,[rfReplaceAll,rfIgnoreCase]);
  strT := edtSociete.Text ;
  if pos ('%societe%',strnomcomplet)<> 0 then strT := lowercase (strT);
  if pos ('%Societe%',strnomcomplet)<> 0 then strT := AnsiFirstUpCase(strT);
  if pos ('%SOCIETE%',strnomcomplet)<> 0 then strT := UpperCase(strT);
  if strT ='' then
    strtemp := StringReplace(strtemp,'%societe%','',[rfReplaceAll,rfIgnoreCase]) else
    strtemp := StringReplace(strtemp,'%societe%',strT + ' -',[rfReplaceAll,rfIgnoreCase]);
  strT := cbxApartenance.Items[cbxApartenance.itemindex] ;
  if pos ('%appartenance%',strnomcomplet)<> 0 then strT := lowercase (strT);
  if pos ('%Appartenance%',strnomcomplet)<> 0 then strT := AnsiFirstUpCase(strT);
  if pos ('%APPARTENANCE%',strnomcomplet)<> 0 then strT := UpperCase(strT);
  strtemp := StringReplace(strtemp,'%appartenance%',strT,[rfReplaceAll,rfIgnoreCase]);
  edtNomComplet.Text :=trim (strtemp);

// ****************************************************************************
// * Ajout de la ville dans le champ description si exterieur on rajoute la   *
// * date de départ                                                           *
// ****************************************************************************

  edtDescription.Text := sitetemp.strVille ;
  if cbxDuree.Checked then
    edtDescription.text := edtDescription.text + ' -> ' + DateToStr (dtDateDepart.Date );
  lstgroupe.Items.clear ;

// ****************************************************************************
// * Le champ societe est vide c'est un agent groupe si non un exterieur      *
// ****************************************************************************

  if edtSociete.Text ='' then
  begin
    edtCompagnie.text := entitetemp.strNom ;
    lstgroupe.Items := sitetemp.lstGrpGroupe ;
  end else
  begin
    edtcompagnie.Text := edtSociete.Text ;
    lstgroupe.Items := sitetemp.lstGrpExterieur;
  end;
  edtVille.Text := sitetemp.strVille ;
  if edtSociete.Text <>'' then cbxDuree.Checked := true;
  cbxDuree.Enabled := edtSociete.Text ='';
  edtLogonScript.Text := strlogonscript;

// ****************************************************************************
// * Est il possible de rajouter un lecteur perso                             *
// ****************************************************************************

  cbxP.Checked := sitetemp.strSRVLecteur <>'';
  cbxp.Enabled := sitetemp.strSRVLecteur <>'';
  cbxBanque.Clear ;
  if edtSociete.Text = '' then
  begin
    for i :=0 to sitetemp.lstbanqueGroupe.count -1 do
      cbxBanque.AddItem(tbanque (sitetemp.lstbanqueGroupe[i]).strnom,sitetemp.lstbanqueGroupe[i]);
  end else
  begin
    for i :=0 to sitetemp.lstBanqueExterieur.count -1 do
      cbxBanque.AddItem(tbanque (sitetemp.lstBanqueExterieur[i]).strnom,sitetemp.lstBanqueExterieur[i]);
  end;
end;

procedure TfrmUser.cbxDureeClick(Sender: TObject);

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet d'activer la date d'expiration du compte                         ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  dtDateDepart.Enabled := cbxDuree.Checked ;
  edtChange(nil);
end;

procedure TfrmUser.edtLoginChange(Sender: TObject);

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet de modifier le profile TSE en fonction du login                  ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if edtlogin.text <>'' then
    edtProfileTse.Text := StringReplace(strProfileTse ,'%login%',edtLogin.Text,[rfReplaceAll])
  else
    edtProfileTse.text :='';
end;

procedure TfrmUser.cmdCreerClick(Sender: TObject);

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet de valider la création de l'utilisateur                          ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  strtemp : string;
begin

// ****************************************************************************
// * Verification que les champs obligatoire soient remplis                   *
// ****************************************************************************

  strtemp := '';
  if trim(edtNom.text)='' then strtemp := 'Nom ,';
  if trim(edtPrenom.Text)='' then strtemp := strtemp + 'Prénom ,';
  if trim(edtUID.Text)='' then strtemp := strtemp + 'UID ,';
  if trim(edtLogin.Text)='' then strtemp := strtemp + 'Login ,';
  if trim(edtMail.Text)='' then strtemp := strtemp + 'Mail ,';
  if trim(edtNomComplet.Text)='' then strtemp := strtemp + 'Nom complet ,';
  if trim(edtPassword.Text)='' then strtemp := strtemp + 'Mot de passe ,';
  if trim(edtDescription.Text)='' then strtemp := strtemp + 'Description ,';
  if cbxApartenance.ItemIndex =-1 then strtemp := strtemp + 'Entité ,';
  if cbxLocalisation.ItemIndex =-1 then strtemp := strtemp + 'Localisation ,';
  if strtemp <> '' then
  begin
    messagedlg('Le ou les champs suivant ne doivent pas être vides' + #10#13 + strtemp ,
               mtError ,[mbok],0);
    exit;
  end;
  ModalResult := mrok;
end;

procedure TfrmUser.cmdSupprimerClick(Sender: TObject);

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet d'enlever le groupe séléctionné                                  ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  lstgroupe.DeleteSelected ;
end;

procedure TfrmUser.cmdAddGrpClick(Sender: TObject);

//  ___________________________________________________________________________
// | function AnsiFirstUpCase                                                  |
// | _________________________________________________________________________ |
// || Permet d'ajouter un ou des groupes                                      ||
// ||_________________________________________________________________________||
// || Entrées | sender : tobject                                              ||
// ||         |   objet générant l'événement                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i,j : integer;
  booexist : boolean;
  lstgrp : tstringlist;
begin
  lstgrp := frmGroupe.GetGroupe;
  for i := 0 to lstgrp.Count -1 do
  begin
    booexist := false;
    for j:=0 to lstgroupe.Count -1 do
      if LowerCase(lstGroupe.Items[j])= lowercase(lstgrp.Strings[i]) then booexist := true;
    if not booexist then lstGroupe.Items.Add(lstgrp.Strings[i]);
  end;
  lstgrp.Free ;
end;

end.
