unit uEntite;

// |===========================================================================|
// | unit uEntite                                                              |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité contenant la class tentite et tsite                                 |
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
uses Contnrs,classes,dialogs;
type
  tEntite = class(TObject)
  private
    fstrNom : string;
    fstrProfileTSE : string;
    fstrNomComplet : string;
    fstrPassword : string;
    fbooNeverExpire : boolean;
    fstrLoginScript : string;
    flisteSite : Tobjectlist;
  public
    property strNom : string read fstrnom write fstrNom ;
    property strProfileTSE : string read fstrProfileTSE write fstrProfileTSE ;
    property strNomComplet : string read fstrNomComplet write fstrNomComplet ;
    property strPassword : string read fstrPassword write fstrPassword ;
    property booNeverExpire : boolean read fbooNeverExpire write fbooNeverExpire ;
    property strLoginScript : string read fstrLoginScript write fstrLoginScript ;
    property listeSite : TObjectList read flisteSite write flisteSite ;
    constructor create ;
    destructor Destroy; override;
  end;
  tSite = class (tobject)
  private
    fstrNom : string;
    fstrOU : string;
    fstrVille : string;
    fstrSRV : string;
    fstrSRVPartage : string;
    fstrSRVRepertoire : string;
    fstrSRVLecteur : string;
    flstGrpGroupe : tstringlist;
    flstGrpExterieur : tstringlist;
    flstBanqueGroupe : TObjectList;
    flstbanqueExterieur : TObjectList;
  public
    property strNom : string read fstrnom write fstrNom ;
    property strOU : string read fstrOU write fstrOU;
    property strVille : string read fstrVille write fstrVille ;
    property strSRV : string read fstrSRV write fstrSRV ;
    property strSRVPartage : string read fstrSRVPartage write fstrSRVPartage ;
    property strSRVRepertoire : string read fstrSRVRepertoire write fstrSRVRepertoire ;
    property strSRVLecteur : string read fstrSRVLecteur write fstrSRVLecteur ;
    property lstGrpGroupe : tstringlist read flstGrpGroupe write flstGrpGroupe ;
    property lstGrpExterieur : tstringlist read flstGrpExterieur write flstGrpExterieur ;
    property lstbanqueGroupe : TObjectList read flstBanqueGroupe write flstBanqueGroupe ;
    property lstBanqueExterieur : TObjectList read flstbanqueExterieur write flstbanqueExterieur ;
    constructor create ;
    destructor Destroy ; override;
  end;
implementation

constructor tEntite.create ;

//  ___________________________________________________________________________
// | constructor tEntite.create                                                |
// | _________________________________________________________________________ |
// || constructeur de la class tentite                                        ||
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
  flisteSite := TObjectList.Create(true);
end;

destructor tentite.destroy ;

//  ___________________________________________________________________________
// | destructor tEntite.destroy                                                |
// | _________________________________________________________________________ |
// || destructeur de la class tentite                                         ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  flisteSite.Free;
  inherited destroy;
end;

constructor tSite.create ;

//  ___________________________________________________________________________
// | constructor tSite.create                                                  |
// | _________________________________________________________________________ |
// || constructeur de la class tSite                                          ||
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
  flstGrpGroupe := tstringlist.Create ;
  flstGrpExterieur := TStringList.Create ;
  flstBanqueGroupe := TObjectList.Create(false);
  flstbanqueExterieur := TObjectList.Create(false);
end;

destructor tSite.destroy;

//  ___________________________________________________________________________
// | destructor tSite.destroy                                                  |
// | _________________________________________________________________________ |
// || destructeur de la class tSite                                           ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  flstGrpGroupe.Free ;
  flstGrpExterieur.Free;
  flstBanqueGroupe.Free;
  flstbanqueExterieur.Free ;
  inherited destroy;
end;

end.
