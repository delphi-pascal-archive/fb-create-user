unit UFBADSI;

// |===========================================================================|
// | unit UFBADSI                                                              |
// | 2006 F.BASSO                                                              |
// |___________________________________________________________________________|
// | Unité d'interface de l'ADSI                                               |
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
// || 1.0  \__________________________________________________________________ |
// || Création de l'unité avec les fonctions suivantes                        ||
// ||   function  ADsRechercheUtilisateur                                     ||
// || Fonctions externes                                                      ||
// ||   procedure ADsEnumerateMembers                                         ||
// ||   function  ADsGetObject                                                ||
// ||   function  ADsBuildEnumerator                                          ||
// ||   function  ADsFreeEnumerator                                           ||
// ||   function  ADsEnumerateNext                                            ||
// ||   function  ADsBuildVarArrayStr                                         ||
// ||   function  ADsBuildVarArrayInt                                         ||
// ||   function  ADsOpenObject                                               ||
// ||   function  ADsGetLastError                                             ||
// ||   procedure ADsSetLastError                                             ||
// ||   procedure AllocADsMem                                                 ||
// ||   function  FreeADsMem                                                  ||
// ||   function  ReallocADsMem                                               ||
// ||   function  AllocADsStr                                                 ||
// ||   function  FreeADsStr                                                  ||
// ||   function  ReallocADsStr                                               ||
// ||   function  ADsEncodeBinaryData                                         ||
// ||   function  PropVariantToAdsType                                        ||
// ||   function  AdsTypeToPropVariant                                        ||
// ||   Procedure AdsFreeAdsValues                                            ||
// ||_________________________________________________________________________||
// |===========================================================================|

interface
uses Windows, ActiveX, Classes, OleServer,ActiveDs_TLB, SysUTILS,variants;

const

// =============================================================================
//  Constantes de retour de la fonction GetnextRow
// =============================================================================

S_ADS_NOMORE_ROWS   = $00005012;
MAX_ADS_ENUM_COUNT  = 100;


// =============================================================================
//  Déclaration des fonctions
// =============================================================================
function  ADsRechercheUtilisateur(strLogin : string ; strLDAPDomaine : string ; out strLDAPUser :string ):integer;
function  ADsRechercheUtilisateur1(strLogin : string ; strLDAPDomaine : string ; out strLDAPUser :string ):integer;

function  ADsGetObject(lpszPathName:WideString ; const riid:TGUID;out obj):HRESULT;stdcall;
function  ADsBuildEnumerator(const pADsContainer:IADsContainer;out ppEnumVariant:IEnumVARIANT):HRESULT; stdcall;
function  ADsFreeEnumerator(pEnumVariant:IEnumVARIANT):HRESULT; stdcall;
function  ADsEnumerateNext(pEnumVariant:IEnumVARIANT;cElements:ULONG;var pvar:OleVARIANT;var pcElementsFetched:ULONG):HRESULT; stdcall;
function  ADsBuildVarArrayStr(lppPathNames:PWideChar;dwPathNames:DWORD;var pVar:VARIANT):HRESULT; stdcall;
function  ADsBuildVarArrayInt(lpdwObjectTypes:LPDWORD;dwObjectTypes:DWORD;var pVar:VARIANT):HRESULT; stdcall;
function  ADsOpenObject(lpszPathName:WideString;lpszUserName:WideString;lpszPassword:WideString;dwReserved:DWORD;const riid:TGUID;out ppObject):HRESULT; stdcall;
function  ADsGetLastError(var pError:DWORD;lpErrorBuf:PWideChar;dwErrorBufLen:DWORD;lpNameBuf:PWideChar;dwNameBufLen:DWORD):HRESULT; stdcall;
procedure ADsSetLastError(dwErr:DWORD;pszErro,pszProvider:LPWSTR); stdcall;
procedure AllocADsMem(cb:DWORD);stdcall;
function  FreeADsMem(pMem:Pointer):BOOL;stdcall;
function  ReallocADsMem(pOldMem:Pointer;cbOld,cbNew:DWORD):Pointer;stdcall;
function  AllocADsStr(pStr:LPWSTR):LPWSTR;stdcall;
function  FreeADsStr(pStr:LPWSTR):BOOL;stdcall;
function  ReallocADsStr(ppStr:LPWSTR;pStr:LPWSTR):BOOL;stdcall;
function  ADsEncodeBinaryData (pbSrcData:PBYTE;dwSrcLen:DWORD;ppszDestData:LPWSTR):HRESULT;stdcall;
function  PropVariantToAdsType(pVariant:VARIANT;dwNumVariant:DWORD;ppAdsValues:Pointer;pdwNumValues:PDWORD):HRESULT;stdcall;
function  AdsTypeToPropVariant(pAdsValues:Pointer;dwNumValues:DWORD;pVariant:VARIANT):HRESULT;stdcall;
Procedure AdsFreeAdsValues(pAdsValues:Pointer;dwNumValues:DWORD);stdcall;

// -----------------------------------------------------------------------------
implementation
// -----------------------------------------------------------------------------
uses dialogs;
const ADSI = 'activeds.dll';
const ADSLDPC = 'adsldpc.dll';

function  ADsRechercheUtilisateur(strLogin : string ; strLDAPDomaine : string ; out strLDAPUser :string ):integer;
var
  search : IDirectorySearch;
  p   : array[0..0] of PWideChar;
  opt : array of ads_searchpref_info;
  hr  : HResult;
  dwErr : DWord;
  szErr : array[0..255] of WideCHar;
  szName : array[0..255] of WideChar;
  ptrResult : THandle;
  col : ads_search_column ;
begin
  result := -1;
  AdsGetObject(strLDAPDomaine, IDirectorySearch, search);
  p[0] := StringToOleStr('adspath');
  setlength(opt,1);
  opt[0].dwSearchPref := ADS_SEARCHPREF_SEARCH_SCOPE;
  opt[0].vValue.dwType := ADSTYPE_INTEGER;
  opt[0].vValue.Integer := ADS_SCOPE_SUBTREE;
  hr := search.SetSearchPreference(pointer(opt),1);
  if (hr <> 0) then
  begin
    ADsGetLastError(dwErr, @szErr[0], 254, @szName[0], 254);
    ShowMessage(WideCharToString(szErr));
    result := -1;
    Exit;
  end;
  search.ExecuteSearch('(samAccountName=' + strLogin +')',@p[0], 1, ptrResult);
  hr := search.GetNextRow(ptrResult);
  while (hr <> S_ADS_NOMORE_ROWS) do
  begin
    hr := search.GetColumn(ptrResult, p[0],col);
    if Succeeded(hr) then
    begin
      if col.pADsValues <> nil then
      begin
        strLDAPUser := (col.pAdsvalues^.CaseIgnoreString) ;
        result := 0 ;
      end;
      search.FreeColumn(col);
    end;
    Hr := search.GetNextRow(ptrResult);
  end;
end;

function  ADsRechercheUtilisateur1(strLogin : string ; strLDAPDomaine : string ; out strLDAPUser :string ):integer;
var
  search : IDirectorySearch;
  p   : array[0..0] of PWideChar;
  opt : array of ads_searchpref_info;
  hr  : HResult;
  dwErr : DWord;
  szErr : array[0..255] of WideCHar;
  szName : array[0..255] of WideChar;
  ptrResult : THandle;
  col : ads_search_column ;
begin
  result := -1;
  AdsGetObject(strLDAPDomaine, IDirectorySearch, search);
  p[0] := StringToOleStr('adspath');
  setlength(opt,1);
  opt[0].dwSearchPref := ADS_SEARCHPREF_SEARCH_SCOPE;
  opt[0].vValue.dwType := ADSTYPE_INTEGER;
  opt[0].vValue.Integer := ADS_SCOPE_SUBTREE;
  hr := search.SetSearchPreference(pointer(opt),1);
  if (hr <> 0) then
  begin
    ADsGetLastError(dwErr, @szErr[0], 254, @szName[0], 254);
    ShowMessage(WideCharToString(szErr));
    result := -1;
    Exit;
  end;
  search.ExecuteSearch('(cn=' + strLogin +')',@p[0], 1, ptrResult);
  hr := search.GetNextRow(ptrResult);
  while (hr <> S_ADS_NOMORE_ROWS) do
  begin
    hr := search.GetColumn(ptrResult, p[0],col);
    if Succeeded(hr) then
    begin
      if col.pADsValues <> nil then
      begin
        strLDAPUser := (col.pAdsvalues^.CaseIgnoreString) ;
        result := 0 ;
      end;
      search.FreeColumn(col);
    end;
    Hr := search.GetNextRow(ptrResult);
  end;
end;

// =============================================================================
//  Liste des fonctions externes
// =============================================================================

function ADsGetObject;        external ADSI;
function ADsBuildEnumerator;  external ADSI;
function ADsEnumerateNext;    external ADSI;
function ADsFreeEnumerator;   external ADSI;
function ADsBuildVarArrayStr; external ADSI;
function ADsBuildVarArrayInt; external ADSI;
function ADsOpenObject;       external ADSI;

function ADsGetLastError;     external ADSLDPC name 'ADsGetLastError';
procedure ADsSetLastError;    external ADSLDPC name 'ADsSetLastError';
procedure AllocADsMem;        external ADSLDPC name 'AllocADsMem';
function FreeADsMem;          external ADSLDPC name 'FreeADsMem';
function ReallocADsMem;       external ADSLDPC name 'ReallocADsMem';
function AllocADsStr;         external ADSLDPC name 'AllocADsStr';
function FreeADsStr;          external ADSLDPC name 'FreeADsStr';
function ReallocADsStr;       external ADSLDPC name 'ReallocADsStr';
function ADsEncodeBinaryData; external ADSLDPC name 'ADsEncodeBinaryData';

function  PropVariantToAdsType; external ADSI;
function  AdsTypeToPropVariant; external ADSI;
procedure AdsFreeAdsValues;     external ADSI;


end.
