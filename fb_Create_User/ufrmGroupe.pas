unit ufrmGroupe;

// |===========================================================================|
// | unit ufrmGroupe                                                           |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité permettant la sélection de groupes dans l'AD pour le programme      |
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
  Dialogs, StdCtrls,ActiveDs_TLB,ufbadsi, ExtCtrls;

type
  TfrmGroupe = class(TForm)
    lstGroupe: TListBox;
    cmdOK: TButton;
    pnlInfo: TPanel;
  private
    { Déclarations privées }
    function recherchegroupe :integer ;
  public
    { Déclarations publiques }
    procedure ChargerGroupes;
    function GetGroupe: tstringlist;
  end;

var
  frmGroupe: TfrmGroupe;

implementation

{$R *.dfm}

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Privée                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

function TfrmGroupe.recherchegroupe :integer ;

//  ___________________________________________________________________________
// | function  TfrmMaitre.recherchegroupe                                      |
// | _________________________________________________________________________ |
// || Permet de récupérer la liste compléte des groupes                       ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties | result : integer                                              ||
// ||         |   Nombre de groupes trouvés -1 en cas d'erreur                ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  search : IDirectorySearch;
  p : array[0..0] of PWideChar;
  ptrResult : THandle;
  col : ads_search_column ;
  hr : HResult;
  opt : array[0..0] of ads_searchpref_info;
  dwErr : DWord;
  szErr : array[0..255] of WideCHar;
  szName : array[0..255] of WideChar;
  strtemp : string;
  strgroupe : string;
  dd : IADs  ;
begin
  result := 0;
  try

// ****************************************************************************
// * Reccuperation du nom du domaine courant                                  *
// ****************************************************************************

    AdsGetObject('LDAP://RootDSE',IADs,dd );
    strgroupe := dd.get('defaultNamingContext');
    AdsGetObject('LDAP://'+ strgroupe , IDirectorySearch, search);
    p[0] := StringToOleStr('adspath');
    opt[0].dwSearchPref :=ADS_SEARCHPREF_PAGESIZE;
    opt[0].vValue.dwType := ADSTYPE_INTEGER;
    opt[0].vValue.Integer := 7000;
    hr := search.SetSearchPreference(@opt[0],1);
    if (hr <> 0) then
    begin
      ADsGetLastError(dwErr, @szErr[0], 254, @szName[0], 254);
      ShowMessage(WideCharToString(szErr));
      Exit;
    end;

// ****************************************************************************
// * Lancement recherche des groupes                                          *
// ****************************************************************************

    search.ExecuteSearch('(objectCategory=Group)',@p[0], 1, ptrResult);
    hr := search.GetNextRow(ptrResult);
    while (hr <> S_ADS_NOMORE_ROWS) do
    begin
      strtemp := '';
      hr := search.GetColumn(ptrResult, p[0],col);
      if Succeeded(hr) then
      begin
        if col.pADsValues <> nil then
        begin
          lstGroupe.Items.Add( copy ((col.pAdsvalues^.ClassName),11,pos(',',(col.pAdsvalues^.ClassName))-11));
          if lstgroupe.Count mod 100 = 1 then application.ProcessMessages ;
        end;
        search.FreeColumn(col);
      end;
      Hr := search.GetNextRow(ptrResult);
    end;
    result := lstgroupe.Items.Count ;
  except
    result := -1;
  end;
end;

// #===========================================================================#
// #===========================================================================#
// #                                                                           #
// # Partie Public                                                             #
// #                                                                           #
// #===========================================================================#
// #===========================================================================#

procedure TfrmGroupe.ChargerGroupes;

//  ___________________________________________________________________________
// | class procedure TfrmGroupe.ChargerGroupes                                 |
// | _________________________________________________________________________ |
// || Permet de charger la liste compléte des groupes                         ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if not Assigned(frmGroupe ) then exit;
  frmGroupe.lstGroupe.Visible := false;
  frmGroupe.pnlInfo.Visible := true;
  frmGroupe.cmdOK.Visible := false;
  frmGroupe.ClientHeight := frmGroupe.pnlinfo.Height;
  frmGroupe.Visible := true;
  Application.ProcessMessages ;
  frmGroupe.recherchegroupe;
  frmGroupe.Visible := false;
  frmGroupe.lstGroupe.Visible := true;
  frmGroupe.pnlInfo.Visible := false;
  frmGroupe.cmdOK.Visible := true;
  frmGroupe.Height := 534;
end;

function TfrmGroupe.GetGroupe : tstringlist;

//  ___________________________________________________________________________
// | class function TfrmGroupe.GetGroupe                                       |
// | _________________________________________________________________________ |
// || Permet de récupérer la sélection d'un ou plusieurs groupes              ||
// ||_________________________________________________________________________||
// || Entrées |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// || Sorties | result : tstringlist                                          ||
// ||         |   liste des groupes sélectionnés                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  i : integer;
begin
  if not Assigned(frmGroupe ) then exit;
  result := TStringList.Create ;
  if frmGroupe.lstGroupe.Count = 0 then ChargerGroupes;
  if frmGroupe.ShowModal = mrok then
  begin
    if frmgroupe.lstgroupe.selcount> 0 then
    for i := 0 to frmgroupe.lstgroupe.count -1 do
      if frmgroupe.lstgroupe.Selected[i] then
        result.Add(frmgroupe.lstgroupe.Items[i] );
  end;
end;

end.
