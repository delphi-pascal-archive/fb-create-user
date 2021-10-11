unit ufonctionDiv;

// |===========================================================================|
// | unit ufonctionDiv                                                         |
// | 2010 F.BASSO                                                              |
// |___________________________________________________________________________|
// | unité permettant d'avoir les informations de version de copyrigth d'un    |
// | programme                                                                 |
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
  Function StrLectureVersion : String;
  Function StrLectureNomProduit : String;
  Function StrLectureCopyright : String;

implementation

uses forms,types,windows,sysutils;

function GetInformation (strinfo : string):string;
Var
  S         : String;
  Taille    : DWord;
  Buffer    : PChar;
  VersionPC : PChar;
  VersionL  : DWord;

Begin
  Result:='';
  // On demande la taille des informations sur l'application
  S := Application.ExeName;
  Taille := GetFileVersionInfoSize(PChar(S), Taille);
  If Taille>0
  Then Try
  // Réservation en mémoire d'une zone de la taille voulue
    Buffer := AllocMem(Taille);
  // Copie dans le buffer des informations
    GetFileVersionInfo(PChar(S), 0, Taille, Buffer);
  // Recherche de l'information de version
    If VerQueryValue(Buffer, PChar('\StringFileInfo\040C04E4\'+strinfo), Pointer(VersionPC), VersionL)
      Then Result:=VersionPC;
  Finally
    FreeMem(Buffer, Taille);
  End;
end;

Function StrLectureVersion : String;

//  ___________________________________________________________________________
// | function  StrLectureVersion                                               |
// | _________________________________________________________________________ |
// || Permet d'avoir le numero de version de l'appli                          ||
// ||_________________________________________________________________________||
// || Sorties | String                                                        ||
// ||         |   Version de l'appli                                          ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  result := GetInformation('FileVersion');
End;

Function StrLectureNomProduit : String;

//  ___________________________________________________________________________
// | function  StrLectureNomProduit                                            |
// | _________________________________________________________________________ |
// || Permet d'avoir le nom de l'appli                                        ||
// ||_________________________________________________________________________||
// || Sorties | String                                                        ||
// ||         |   Nom de l'appli                                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  result := GetInformation('InternalName');
End;

Function StrLectureCopyright : String;

//  ___________________________________________________________________________
// | function  StrLectureCopyright                                             |
// | _________________________________________________________________________ |
// || Permet d'avoir le Copyright de l'appli                                  ||
// ||_________________________________________________________________________||
// || Sorties | String                                                        ||
// ||         |   Copyright de l'appli                                        ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  result := GetInformation('LegalCopyright');
End;


end.
