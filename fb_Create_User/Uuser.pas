unit Uuser;

// |===========================================================================|
// | unit Uuser                                                                |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité contenant la class tuser                                            |
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
uses classes,Contnrs;

type
  tUser = class(TObject)
  private
    fetat : integer;
    fstrNom : string;
    fstrPrenom : string;
    fstrLogin : string;
    fstrDescription : string;
    fstrNomComplet : string;
    fstrUid : string;
    fstrPassword : string;
    fbooPasswordNeverExpire : boolean;
    fbooChangePassword : boolean;
    fstrLogonScript : string;
    fstrAppartenace : string;
    fstrSite : string;
    fstrDepartement : string;
    fstrSociete : string;
    fstrAddresseMail : string;
    fstrVille : string;
    fstrCompagnie : string;
    fdtDateExpiration : TDateTime ;
    fstrRepertoirePerso : string;
    fstrLecteur : string;
    fstrProfileTse : string;
    flisteGroupes : tstrings;
    fstrTel : string;
    fstrFax : string;
    fstrOU : string;
    fstrMatricule : string;
    fstrErreur : string;
    fstrSRV : string;
    fstrSRVPartage : string;
    fstrSRVRepertoire : string;
    fstrBanque : string;
  public
    property intEtat :  integer read fetat write fetat;
    property strErreur : string read fstrErreur write fstrErreur;
    property strNom : string read fstrnom write fstrnom;
    property strPrenom : string read fstrPrenom write fstrPrenom;
    property strLogin : string read fstrLogin write fstrLogin;
    property strDescription : string read fstrDescription write fstrDescription;
    property strNomComplet : string read fstrNomComplet write fstrNomComplet;
    property strUid : string read fstrUid write fstrUid;
    property strPassword : string read fstrPassword write fstrPassword;
    property booPasswordNeverExpire : boolean read fbooPasswordNeverExpire write fbooPasswordNeverExpire;
    property booChangePassword : boolean read fbooChangePassword write fbooChangePassword;
    property strLogonScript : string read fstrLogonScript write fstrLogonScript;
    property strDepartement : string read fstrDepartement write fstrDepartement;
    property strAppartenance : string read fstrAppartenace  write fstrAppartenace;
    property strSIte : string read fstrSite  write fstrSite;
    property strSociete : string read fstrSociete write fstrSociete;
    property strAddresseMail : string read fstrAddresseMail write fstrAddresseMail;
    property strVille : string read fstrVille write fstrVille;
    property strCompagnie : string read fstrCompagnie write fstrCompagnie;
    property dtDateExpiration : TDateTime read fdtDateExpiration write fdtDateExpiration;
    property strRepertoirePerso : string read fstrRepertoirePerso write fstrRepertoirePerso;
    property strLecteur : string read fstrLecteur write fstrLecteur;
    property strProfileTse : string read fstrProfileTse write fstrProfileTse;
    property listeGroupes : tstrings read flisteGroupes write flisteGroupes;
    property strTel : string read fstrTel write fstrTel;
    property strFax : string read fstrFax write fstrFax;
    property strOU : string read fstrou write fstrou;
    property strMatricule : string read fstrMatricule write fstrMatricule ;
    property strSRV : string read fstrSRV write fstrSRV ;
    property strSRVPartage : string read fstrSRVPartage write fstrSRVPartage ;
    property strSRVRepertoire : string read fstrSRVRepertoire write fstrSRVRepertoire ;
    property strBanque : string read fstrBanque write fstrBanque ;
    constructor create ;
    destructor Destroy; override;
    procedure copy(user : tuser);
  end;

implementation

constructor tuser.create ;

//  ___________________________________________________________________________
// | constructor tuser.create                                                  |
// | _________________________________________________________________________ |
// || constructeur de la class tuser                                          ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  inherited create;
  fetat := 0;
  flisteGroupes := tstringlist.Create ;
end;

destructor tuser.Destroy ;

//  ___________________________________________________________________________
// | destructor tuser.Destroy                                                  |
// | _________________________________________________________________________ |
// || destructeur de la class tuser                                           ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  flisteGroupes.free;
  inherited destroy;
end;

procedure tuser.copy(user : tuser);

//  ___________________________________________________________________________
// | procedure tuser.copy                                                      |
// | _________________________________________________________________________ |
// || permet la copy des informations d'un user                               ||
// ||_________________________________________________________________________||
// || Entrées | user : tuser                                                  ||
// ||         |   tuser à copier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  fetat := user.fetat ;
  fstrErreur := user.fstrErreur;
  fstrNom := user.fstrNom ;
  fstrPrenom := user.fstrPrenom ;
  fstrLogin := user.fstrLogin ;
  fstrDescription := user.fstrDescription;
  fstrNomComplet := user.fstrNomComplet;
  fstrUid := user.fstrUid;
  fstrPassword := user.fstrPassword;
  fbooPasswordNeverExpire := user.fbooPasswordNeverExpire;
  fbooChangePassword := user.fbooChangePassword;
  fstrLogonScript := user.fstrLogonScript;
  fstrAppartenace := user.fstrAppartenace;
  fstrSite := user.fstrSite;
  fstrDepartement := user.fstrDepartement;
  fstrSociete := user.fstrSociete;
  fstrAddresseMail := user.fstrAddresseMail;
  fstrVille := user.fstrVille;
  fstrCompagnie := user.fstrCompagnie;
  fdtDateExpiration := user.dtDateExpiration;
  fstrRepertoirePerso := user.fstrRepertoirePerso;
  fstrLecteur := user.fstrLecteur;
  fstrProfileTse := user.fstrProfileTse;
  flisteGroupes.Clear ;
  for i:= 0 to user.flisteGroupes.Count -1 do
    flisteGroupes.Add ( user.flisteGroupes[i]);
  fstrTel := user.fstrTel;
  fstrFax := user.fstrFax;
  fstrOU := user.fstrOU;
  fstrMatricule := user.fstrMatricule ;
  fstrSRV := user.fstrSRV;
  fstrSRVPartage  := user.fstrSRVPartage;
  fstrSRVRepertoire := user.fstrSRVRepertoire ;
  fstrBanque := user.fstrBanque;
end;

end.
