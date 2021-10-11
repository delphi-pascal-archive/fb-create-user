unit uBanques;

interface
type
  tBanque = class(TObject)
  private
    fId : integer;
    fNom : String;
    fChemin : string;
  public
    property intId : integer read fid write fId ;
    property strNom : string read fNom write fNom ;
    property strChemin : string read fChemin write fChemin ;
  end;

implementation

end.
