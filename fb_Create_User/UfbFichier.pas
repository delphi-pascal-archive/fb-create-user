unit UfbFichier;

// |===========================================================================|
// | UNIT UfbFichier                                                           |
// | Copyright © 2009 F.BASSO                                                  |
// |___________________________________________________________________________|
// | Unité contenant les fonctions de manipulation de fichiers avec prise en   |
// | charge des nom de fichier > 255 caractères                                |
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
// |   1.1.0.0 Ajout FBFileSetShortName, FBFileGetShortName                    |
// |   1.0.0.0 Création de l'unité                                             |
// |===========================================================================|

interface
type
  TonBeforeDeleteFile = procedure (Max : integer)of object;
  TonprogressDeleteFile = procedure (position : integer)of object;

const
  fmOpenRead       = $0000;
  fmOpenWrite      = $0001;
  fmOpenReadWrite  = $0002;

  fmShareCompat    = $0000 ; // DOS compatibility mode is not portable
  fmShareExclusive = $0010;
  fmShareDenyWrite = $0020;
  fmShareDenyRead  = $0030 ; // write-only not supported on all platforms
  fmShareDenyNone  = $0040;

  faReadOnly  = $00000001 ;
  faHidden    = $00000002 ;
  faSysFile   = $00000004 ;
  faVolumeID  = $00000008  deprecated;  // not used in Win32
  faDirectory = $00000010;
  faArchive   = $00000020 ;
  faSymLink   = $00000040 ;
  faAnyFile   = $0000003F;

var
  FBonprogressdelete : TonprogressDeleteFile ;
  FBonbeforedelete : TonBeforeDeleteFile ;
  procedure assignB( p : TonBeforeDeleteFile);
  function  FBFileSetAttr ( const FileName: string; Attr: Integer ): Integer;
  function  FBFileGetAttr ( const FileName: string ): Integer;
  function  FBFileDelete ( const FileName : string ): boolean; overload;
  function  FBFileDelete ( const FileName : string ; booBroyage : boolean ): boolean;overload;
  function  FBFileOpen ( const FileName: string; Mode: LongWord ): Integer;
  function  FBFileCreate ( const FileName: string ): Integer;
  procedure FBFileClose ( Handle: Integer );
  function  FBGetFileCreationTime ( const FileName: string ): TDateTime;
  function  FBGetFileLastWriteTime ( const FileName: string ): TDateTime;
  function  FBGetFileLastAccessTime ( const FileName: string ): TDateTime;
  function  FBFileRename ( const OldName, NewName: string ): Boolean;
  function  FBFileWrite ( Handle: Integer; const Buffer; Count: LongWord ): Integer;
  function  FBFileRead ( Handle: Integer; var Buffer; Count: LongWord ): Integer;
  function  FBFileSeek ( Handle, Offset, Origin: Integer ): Integer;overload;
  function  FBFileSeek ( Handle: Integer; const Offset: Int64; Origin: Integer ): Int64;overload;
  function  FBGetFileSize ( const FileName: string ): Int64;
  function  FBFileExists ( const FileName: string ): boolean;
  function  FBFileSetShortName ( Const FileName : string ; Const shortfilename : string ) : boolean;
  function  FBFileGetShortName ( Const FileName : string ) : string;
  function  FBFileGetLongName (const Filename : string) : string;
  function  strCheminToUnicode ( Const Chemin:string ):pwidechar;

  function  FBDirectoryExists ( const DirName: string ): boolean;
  function  FBDirectoryRemove ( const FileName : string ): boolean;overload;
  function  FBDirectoryRemove ( const FileName : string ;booBroyage : boolean ): boolean;overload;
  function  FBDirectoryRemoveR ( const FileName : string ): integer;
  Function  FBDirectoryCreate ( const DirName: string ) : boolean;
  function  FBForceDirectories ( Dir: string ): Boolean;

  function  FBChangeFileExt ( const FileName, Extension: string ): string;
  function  FBExtractFilePath ( const FileName: string ): string;
  function  FBExtractFileDir ( const FileName: string ): string;
  function  FBExtractFileDrive ( const FileName: string ): string;
  function  FBExtractFileName ( const FileName: string ): string;
  function  FBExtractFileExt ( const FileName: string ): string;

implementation

uses windows,dialogs,sysutils;

type
  PDayTable = ^TDayTable;
  TDayTable = array[1..12] of Word;
  Int64Rec = packed record
    case Integer of
      0: (Lo, Hi: Cardinal);
      1: (Cardinals: array [0..1] of Cardinal);
      2: (Words: array [0..3] of Word);
      3: (Bytes: array [0..7] of Byte);
  end;

const
  MonthDays: array [Boolean] of TDayTable =
    ((31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31),
     (31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31));

function SetFileShortNameW(hFile: THandle;lpFileName: PWideChar): longbool ; stdcall; external 'Kernel32.dll' name 'SetFileShortNameW';

function EncodeTime(Hour, Min, Sec, MSec: Word): TDateTime;
begin
  if (Hour < 24) and (Min < 60) and (Sec < 60) and (MSec < 1000) then
    result := (Hour * 3600000 + Min * 60000 + Sec * 60 + MSec) / 86400000
  else
    result := -1;
end;

function IsLeapYear(Year: Word): Boolean;
begin
  Result := (Year mod 4 = 0) and ((Year mod 100 <> 0) or (Year mod 400 = 0));
end;

function EncodeDate(Year, Month, Day: Word): TDateTime;
var
  I: Integer;
  DayTable: PDayTable;
begin
  DayTable := @MonthDays[IsLeapYear(Year)];
  if (Year >= 1) and (Year <= 9999) and (Month >= 1) and (Month <= 12) and
    (Day >= 1) and (Day <= DayTable^[Month]) then
  begin
    for I := 1 to Month - 1 do Inc(Day, DayTable^[I]);
    I := Year - 1;
    result := I * 365 + I div 4 - I div 100 + I div 400 + Day - 693594;
  end else
    result := -1
end;

function SystemTimeToDateTime(const SystemTime: TSystemTime): TDateTime;
begin
  with SystemTime do
  begin
    Result := EncodeDate(wYear, wMonth, wDay);
    if Result >= 0 then
      Result := Result + EncodeTime(wHour, wMinute, wSecond, wMilliSeconds)
    else
      Result := Result - EncodeTime(wHour, wMinute, wSecond, wMilliSeconds);
  end;
end;

function ExcludeTrailingPathDelimiter(const S: string): string;
begin
  Result := S;
  if Result[Length(Result)]='\' then
    SetLength(Result, Length(Result)-1);
end;

procedure assignb( p : TonBeforeDeleteFile );
begin
  FBonbeforedelete := p;

end;

function SetPrivilege(Privilege: PChar; EnablePrivilege: boolean; out PreviousState: boolean): DWORD;
var
  Token: THandle;
  NewState: TTokenPrivileges;
  Luid: TLargeInteger;
  PrevState: TTokenPrivileges;
  Return: DWORD;
begin
  PreviousState := True;
  if (GetVersion() > $80000000) then
    Result := ERROR_SUCCESS
  else
  begin
    if not OpenProcessToken(GetCurrentProcess(), MAXIMUM_ALLOWED, Token) then
      Result := GetLastError()
    else
    try
      if not LookupPrivilegeValue(nil, Privilege, Luid) then
        Result := GetLastError()
      else
      begin
        NewState.PrivilegeCount := 1;
        NewState.Privileges[0].Luid := Luid;
        if EnablePrivilege then
          NewState.Privileges[0].Attributes := SE_PRIVILEGE_ENABLED
        else
          NewState.Privileges[0].Attributes := 0;
        if not AdjustTokenPrivileges(Token, False, NewState,
          SizeOf(TTokenPrivileges), PrevState, Return) then
          Result := GetLastError()
        else
        begin
          Result := ERROR_SUCCESS;
          PreviousState :=
            (PrevState.Privileges[0].Attributes and SE_PRIVILEGE_ENABLED <> 0);
        end;
      end;
    finally
      CloseHandle(Token);
    end;
  end;
end;

function FileTimetoDatetime (filetime: tfiletime):tdatetime;

//  ___________________________________________________________________________
// | function FileTimetoDatetime                                               |
// | _________________________________________________________________________ |
// || Permet la convertion des dates fichier en dates tdatetime               ||
// ||_________________________________________________________________________||
// || Entrées | filetime: tfiletime                                           ||
// ||         |   date à convertir                                            ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : tdatetime                                            ||
// ||         |   date convertie                                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  SysTime : TSystemTime;
begin
  fileTimeTosystemtime (filetime,SysTime);
  result := SystemTimeToDateTime(SysTime);
end;

function strCheminToUnicode(Const Chemin:string):pwidechar;

//  ___________________________________________________________________________
// | function strCheminToUnicode                                               |
// | _________________________________________________________________________ |
// || Permet de transformer un chemin de fichier en chemin compatible avec    ||
// || Les fonction unicode windows                                            ||
// ||_________________________________________________________________________||
// || Entrées | Const Chemin:string                                           ||
// ||         |   Chemin à convertir                                          ||
// ||_________|_______________________________________________________________||
// || Sorties | result : pwidechar                                            ||
// ||         |   Chemin converti                                             ||
// ||         |   "\\?\c:\rep"  ou "\\?\UNC\serveur\partage"                  ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
//  GetMem(result, sizeof(WideChar) * Succ(Length(Chemin)+8));

  if pos('\\?\',chemin)= 1 then
    Result := PWideChar(WideString(chemin))
  else
    if pos('\\',Chemin)=1 then
      result :=PWideChar(WideString('\\?\UNC\'+ copy(chemin,3,length(chemin)-2)))
    else
      result :=PWideChar(WideString('\\?\'+Chemin));
end;

function FBFileSetAttr(const FileName: string; Attr: Integer): Integer;

//  ___________________________________________________________________________
// | function FBFileSetAttr                                                    |
// | _________________________________________________________________________ |
// || Permet de modifier les attributs d'un fichier                           ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||         | Attr: Integer                                                 ||
// ||         |   Attributs à affecter au fichier                             ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : integer                                              ||
// ||         |    zéro si l'exécution de la fonction réussit                 ||
// ||         |    Sinon, la valeur renvoyée est un code d'erreur             ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Result := 0;
  if not SetFileAttributesW(strCheminToUnicode(FileName), Attr) then
    Result := GetLastError;
end;

function FBFileGetAttr(const FileName: string): Integer;

//  ___________________________________________________________________________
// | function FBFileGetAttr                                                    |
// | _________________________________________________________________________ |
// || Permet de récupérer les attribut d'un fichier                           ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : integer                                              ||
// ||         |   valeur des attributs du fichier sous forme de masque        ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Result := GetFileAttributesW(strCheminToUnicode(FileName));
end;

function FBFileDelete(const FileName : string): boolean;

//  ___________________________________________________________________________
// | function FBFileDelete                                                     |
// | _________________________________________________________________________ |
// || Permet de supprimer un fichier                                          ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName : string                                       ||
// ||         |   Nom du fichier à supprimer                                  ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |   Vrai en cas de réussite, faux le cas contraire voir         ||
// ||         |   GetLastError pour connaitre la raison                       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  result := DeleteFileW(strCheminToUnicode(FileName));
end;

function FBFileDelete(const FileName : string ; booBroyage : boolean): boolean;

//  ___________________________________________________________________________
// | function FBFileDelete                                                     |
// | _________________________________________________________________________ |
// || Permet de supprimer un fichier et de rendre ça restauration dificile    ||
// || via une double écriture de 1010 et de 0101                              ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName : string                                       ||
// ||         |   Nom du fichier à supprimer                                  ||
// ||         | BooBroyage : boolean                                          ||
// ||         |   Vrai pour appliquer un écrasement se sécurité               ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |   Vrai en cas de réussite, faux le cas contraire voir         ||
// ||         |   GetLastError pour connaitre la raison                       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

const
  tbuf =4096;
var
  fh : integer;
  i, taille : int64;
  buf : array [0..tbuf-1] of byte;
  strNomfichier : string;
  boo : boolean;
begin
  SetPrivilege('SeRestorePrivilege',true,boo);
  FBFileSetAttr(FileName,faArchive);
  strNomfichier := filename;
  if booBroyage then
  begin

// ****************************************************************************
// * Tentative de déplacer le fichier à la racine et de le renommer pour      *
// * brouiller les pistes                                                     *
// ****************************************************************************

    if FBFileExists(FBExtractFileDrive(FileName )+'\fb000.0000') then
    begin
      FBFileSetAttr(FBExtractFileDrive(FileName )+'\fb000.0000',faArchive);
      FBFileDelete(FBExtractFileDrive(FileName )+'\fb000.0000');
    end;
    if FBfileRename(FileName, FBExtractFileDrive(FileName )+'\fb000.0000') then
      strNomfichier := FBExtractFileDrive(FileName )+'\fb000.0000' else strnomfichier := FileName ;
    fh := FBFileOpen(strnomfichier,fmOpenReadWrite	 );
    if fh <> -1 then
    begin
      taille:=FBGetFileSize (strnomfichier);
      taille :=( taille div tbuf) +1;
      if assigned( FBonbeforedelete) then FBonbeforedelete(taille *2);
      FillChar ( buf,tbuf,$AA);
      FBFileSeek(fh,0,0);
      i := 0 ;

// ****************************************************************************
// * Ecriture d'une suite de 10101010                                         *
// ****************************************************************************

      repeat
        FBFileWrite(fh,buf[0],tbuf);
        i := i+1;
        if assigned(FBonprogressdelete ) then FBonprogressdelete(i);
      until i= taille;
      FlushFileBuffers(fh);
      FillChar ( buf,tbuf,$55);
      FBFileSeek(fh,0,0);
      i := 0 ;

// ****************************************************************************
// * Ecriture d'une suite de 01010101                                         *
// ****************************************************************************

      repeat
        FBFileWrite(fh,buf[0],tbuf);
        i := i+1;
        if assigned(FBonprogressdelete ) then FBonprogressdelete(taille + i);
      until i= taille;
      FlushFileBuffers(fh);
      FBfileclose(fh);
    end;
  end;
  result := FBFileDelete(strNomFichier );
  if not boo then SetPrivilege('SeRestorePrivilege',boo,boo);
end;


function FBFileOpen(const FileName: string; Mode: LongWord): Integer;

//  ___________________________________________________________________________
// | function FBFileOpen                                                       |
// | _________________________________________________________________________ |
// || Permet d'ouvrir un fichier et obtenir un handle de fichier              ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier à ouvrir                                     ||
// ||         | string; Mode: LongWord                                        ||
// ||         |   mode d'accès au fichier                                     ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : integer                                              ||
// ||         |   >-1 handle de fichier , =-1 le fichier n'a pu être ouvert   ||                                                            ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

const
  AccessMode: array[0..2] of LongWord = (
    GENERIC_READ,
    GENERIC_WRITE,
    GENERIC_READ or GENERIC_WRITE);
  ShareMode: array[0..4] of LongWord = (
    0,
    0,
    FILE_SHARE_READ,
    FILE_SHARE_WRITE,
    FILE_SHARE_READ or FILE_SHARE_WRITE);
begin
  Result := -1;
  if ((Mode and 3) <= fmOpenReadWrite) and
    ((Mode and $F0) <= fmShareDenyNone) then
    Result := Integer(CreateFilew(strCheminToUnicode(FileName), AccessMode[Mode and 3],
      ShareMode[(Mode and $F0) shr 4], nil, OPEN_EXISTING,
      FILE_ATTRIBUTE_NORMAL, 0));
end;

function FBFileCreate(const FileName: string): Integer;

//  ___________________________________________________________________________
// | function FBFileCreate                                                     |
// | _________________________________________________________________________ |
// || Permet de créer un nouveau fichier et d'obtenir un handle de fichier    ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier à créer                                      ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : integer                                              ||
// ||         |   >-1 handle de fichier , =-1 le fichier n'a pu être créé     ||                                                            ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Result := Integer(CreateFileW(strCheminToUnicode(FileName), GENERIC_READ or GENERIC_WRITE,
    0, nil, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0));
end;

function FBGetFileSize(const FileName: string): Int64;

//  ___________________________________________________________________________
// | function FBGetFileSize                                                    |
// | _________________________________________________________________________ |
// || Pemret de retourner la taille d'un fichier                              ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | Retsult : int64                                               ||
// ||         |   Taille du fichier en octet                                  ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  if FH = INVALID_HANDLE_VALUE then Result := -1
  else
    try
      Result := FD.nFileSizeHigh shl 32 + FD.nFileSizeLow;
    finally
      windows.findclose(FH);
    end;
end;

function FBGetFileCreationTime(const FileName: string): TDateTime;

//  ___________________________________________________________________________
// | function FBGetFileCreationTime                                            |
// | _________________________________________________________________________ |
// || Permet de connaître la date de création du fichier                      ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : tdatetime                                            ||
// ||         |   Date de création du fichier                                 ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  Result := -1;
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  if FH = INVALID_HANDLE_VALUE then Result := -1
  else
    try
      Result := FileTimetoDatetime (fd.ftCreationTime);
    finally
      windows.findclose(FH);
    end;
end;

function FBGetFileLastWriteTime(const FileName: string): TDateTime;

//  ___________________________________________________________________________
// | function FBGetFileLastWriteTime                                           |
// | _________________________________________________________________________ |
// || Permet de connaître la date du dernier enregistrement du fichier        ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : tdatetime                                            ||
// ||         |   Date du dernier enregistrement                              ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  Result := -1;
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  if FH = INVALID_HANDLE_VALUE then Result := -1
  else
    try
      Result := FileTimetoDatetime (fd.ftLastWriteTime);
    finally
      windows.findclose(FH);
    end;
end;

function FBGetFileLastAccessTime(const FileName: string): TDateTime;

//  ___________________________________________________________________________
// | function FBGetFileLastAccessTime                                          |
// | _________________________________________________________________________ |
// || Permet de connaître la date du dernier accés au fichier                 ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : tdatetime                                            ||
// ||         |   Date du dernier accés                                       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  Result := -1;
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  if FH = INVALID_HANDLE_VALUE then Result := -1
  else
    try
      Result := FileTimetoDatetime (fd.ftLastAccessTime);
    finally
      windows.findclose(FH);
    end;
end;

//  function FBFileExists(const FileName: string): boolean;
//
//  //  ___________________________________________________________________________
//  // | function FBFileExists                                                     |
//  // | _________________________________________________________________________ |
//  // || Permet de savoir si un fichioer existe                                  ||
//  // || cette fonction gére les noms de fichier > 255 Caractères                ||
//  // ||_________________________________________________________________________||
//  // || Entrées | const FileName: string                                        ||
//  // ||         |   Nom du fichier                                              ||
//  // ||_________|_______________________________________________________________||
//  // || Sorties | result : boolean                                              ||
//  // ||         |   Vrai si le fichier existe si non Faux                       ||
//  // ||_________|_______________________________________________________________||
//  // |___________________________________________________________________________|
//
//  var
//    Code: Integer;
//  begin
//    Code := GetFileAttributesW(strCheminToUnicode(FileName));
//    Result := (Code <> -1) and (FILE_ATTRIBUTE_DIRECTORY and Code = 0);
//
//    //result := FBGetFileLastWriteTime(FileName)> -1;
//  end;
function FBFileExists(const FileName: string): boolean;

//  ___________________________________________________________________________
// | function FBFileExists                                                     |
// | _________________________________________________________________________ |
// || Permet de savoir si un fichioer existe                                  ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | result : boolean                                              ||
// ||         |   Vrai si le fichier existe si non Faux                       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  result := (FH <> INVALID_HANDLE_VALUE) and ( fd.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY =0) ;
  windows.findclose(FH);
end;

function  FBFileSetShortName ( Const FileName : string ; Const shortfilename : string ) : boolean;

//  ___________________________________________________________________________
// | function FBFileSetShortName                                               |
// | _________________________________________________________________________ |
// || Permet de modifier le nom court d'un fichier                            ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||         | const shortfilename: string                                   ||
// ||         |   Nom court du fichier                                        ||
// ||_________|_______________________________________________________________||
// || Sorties | result : boolean                                              ||
// ||         |   Vrai si le modification OK                                  ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  fh : THandle ;
  boo : boolean;


begin
  SetPrivilege('SeRestorePrivilege',true,boo);
  fh := Integer(CreateFilew(strCheminToUnicode(FileName),
                GENERIC_ALL,
                FILE_SHARE_READ, nil,
                OPEN_EXISTING,
                FILE_FLAG_BACKUP_SEMANTICS, 0)); //FBFileOpen(filename,GENERIC_ALL or FILE_FLAG_BACKUP_SEMANTICS );
  result := SetFileShortNamew(fh,PWideChar(WideString(shortfilename)));
  FBFileClose(fh);
  if not boo then SetPrivilege('SeRestorePrivilege',boo,boo);
end;

function  FBFileGetShortName ( Const FileName : string ) : string;

//  ___________________________________________________________________________
// | function FBFileGetShortName                                               |
// | _________________________________________________________________________ |
// || Permet de trouver le nom court d'un fichier                             ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom du fichier                                              ||
// ||_________|_______________________________________________________________||
// || Sorties | result : string                                               ||
// ||         |   Nom court du fichier                                        ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  result := '';
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  if (FH <> INVALID_HANDLE_VALUE)then
    result := fd.cAlternateFileName;
  windows.findclose(FH);
end;

function FBFileGetLongName (const Filename : string) : string;
var
  FD: TWin32FindDataW;
  FH: THandle;
begin
  result := '';
  FH := 	FindFirstFileW(strCheminToUnicode(FileName ), FD);
  if (FH <> INVALID_HANDLE_VALUE)then
    result := fd.cFileName;
  windows.findclose(FH);
end;


function FBFileRename(const OldName, NewName: string): Boolean;

//  ___________________________________________________________________________
// | function FBFileRename                                                     |
// | _________________________________________________________________________ |
// || Permet de renommer un fichier                                           ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const OldName: string                                         ||
// ||         |   Nom du fichier d'origine                                    ||
// ||         | const NewName: string                                         ||
// ||         |   Nouveau nom du fichier                                      ||
// ||_________|_______________________________________________________________||
// || Sorties | result : boolean                                              ||
// ||         |   Vrai si le renommage du fichier est fait si non Faux        ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  s : widestring;
begin
  s := strCheminToUnicode(OldName);
  Result := MoveFilew(PWideChar(s),strCheminToUnicode(NewName));
end;

procedure FBFileClose(Handle: Integer);

//  ___________________________________________________________________________
// | procedure FBFileClose                                                     |
// | _________________________________________________________________________ |
// || Permet de fermer un fichier                                             ||
// ||_________________________________________________________________________||
// || Entrées | Handle: Integer                                               ||
// ||         |   Handle du fichier à fermer                                  ||
// ||_________|_______________________________________________________________||
// || Sorties |                                                               ||
// ||         |                                                               ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  CloseHandle(THandle(Handle));
End;

function FBFileRead(Handle: Integer; var Buffer; Count: LongWord): Integer;

//  ___________________________________________________________________________
// | function FBFileRead                                                       |
// | _________________________________________________________________________ |
// || Permet de lire des données dans un fichier                              ||
// ||_________________________________________________________________________||
// || Entrées | Handle: Integer                                               ||
// ||         |   Handle du fichier dans le quel lire les données             ||
// ||         | var Buffer : LongWord                                         ||
// ||         |   Tampon qui va recevoir les données lues                     ||
// ||         | Count: LongWord                                               ||
// ||         |   Nombre d'octets à lire                                      ||
// ||_________|_______________________________________________________________||
// || Sorties | var Buffer : LongWord                                         ||
// ||         |   Tampon rempli avec les données                              ||
// ||         | result : integer                                              ||
// ||         |   Nombre d'octets effectivement lues                          ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if not ReadFile(THandle(Handle), Buffer, Count, LongWord(Result), nil) then
    Result := -1;
end;

function FBFileWrite(Handle: Integer; const Buffer; Count: LongWord): Integer;

//  ___________________________________________________________________________
// | function FBFileWrite                                                      |
// | _________________________________________________________________________ |
// || Permet d'écrire des données dans un fichier                             ||
// ||_________________________________________________________________________||
// || Entrées | Handle: Integer                                               ||
// ||         |   Handle du fichier dans le quel écrire les données           ||
// ||         | var Buffer : LongWord                                         ||
// ||         |   Tampon qui contient les données à écrire                    ||
// ||         | Count: LongWord                                               ||
// ||         |   Nombre d'octets à écrire                                    ||
// ||_________|_______________________________________________________________||
// || Sorties | result : integer                                              ||
// ||         |   Nombre d'octets écrit, -1 en cas d'erreur                   ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  if not WriteFile(THandle(Handle), Buffer, Count, LongWord(Result), nil) then
    Result := -1;
end;

function FBFileSeek(Handle, Offset, Origin: Integer): Integer;

//  ___________________________________________________________________________
// | function FBFileSeek                                                       |
// | _________________________________________________________________________ |
// || Pemret de déplacer le pointeur d'écriture/ lecture d'un fichier         ||
// ||_________________________________________________________________________||
// || Entrées | Handle : Integer                                              ||
// ||         |   Handle du fichier dans le quel déplacer le pointeur         ||
// ||         | Offset : Integer                                              ||
// ||         |   décalage par rapport à Origin du pointeur de fichier        ||
// ||         | Origin : Integer                                              ||
// ||         |   code indiquant le début du fichier 0, la fin du fichier 2,  ||
// ||         |   ou encore la position en cours du pointeur 1                ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : integer                                              ||
// ||         |   Nouvel valeur du pointeur ou -1 en cas d'erreur             ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Result := SetFilePointer(THandle(Handle), Offset, nil, Origin);
end;

function FBFileSeek(Handle: Integer; const Offset: Int64; Origin: Integer): Int64;

//  ___________________________________________________________________________
// | function FBFileSeek                                                       |
// | _________________________________________________________________________ |
// || Pemret de déplacer le pointeur d'écriture/ lecture d'un fichier         ||
// ||_________________________________________________________________________||
// || Entrées | Handle : Integer                                              ||
// ||         |   Handle du fichier dans le quel déplacer le pointeur         ||
// ||         | Offset : Int64                                                ||
// ||         |   décalage par rapport à Origin du pointeur de fichier        ||
// ||         | Origin : Integer                                              ||
// ||         |   code indiquant le début du fichier 0, la fin du fichier 2,  ||
// ||         |   ou encore la position en cours du pointeur 1                ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : int64                                                ||
// ||         |   Nouvel valeur du pointeur ou -1 en cas d'erreur             ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Result := Offset;
  Int64Rec(Result).Lo := SetFilePointer(THandle(Handle), Int64Rec(Result).Lo,
    @Int64Rec(Result).Hi, Origin);
end;

function FBDirectoryExists(const DirName: string): boolean;

//  ___________________________________________________________________________
// | function FBDirectoryExists                                                |
// | _________________________________________________________________________ |
// || Permet de savoir si un répertoire existe                                ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const DirName: string                                         ||
// ||         |   Nom du répertoire                                           ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |   Vrai si le répertoire existe si non Faux                    ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  Code: Integer;
begin
  Code := GetFileAttributesW(strCheminToUnicode(DirName));
  Result := (Code <> -1) and (FILE_ATTRIBUTE_DIRECTORY and Code <> 0);
end;

Function FBDirectoryCreate(const DirName: string) : boolean;

//  ___________________________________________________________________________
// | Function FBDirectoryCreate                                                |
// | _________________________________________________________________________ |
// || Permet de créer un repertoire                                           ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const DirName: string                                         ||
// ||         |   Nom du répertoire à créer                                   ||
// ||_________|_______________________________________________________________||
// || Sorties | result : boolean                                              ||
// ||         |   Vrai si le repertoire est bien créé si non faux             ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  result := CreateDirectoryW(strCheminToUnicode(DirName),nil);
end;

function FBForceDirectories(Dir: string): Boolean;

//  ___________________________________________________________________________
// | function FBForceDirectories                                               |
// | _________________________________________________________________________ |
// || Permet de créer un nouveau répertoire et si nécessaire les répertoires  ||
// || parents                                                                 ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const DirName: string                                         ||
// ||         |   Nom du répertoire à créer                                   ||
// ||_________|_______________________________________________________________||
// || Sorties | result : boolean                                              ||
// ||         |   Vrai si le repertoire est bien créé si non faux             ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

begin
  Result := True;
  if Dir = '' then exit;
  Dir := ExcludeTrailingPathDelimiter(Dir);
  if (Length(Dir) < 3) or FBDirectoryExists(Dir) or (FBExtractFilePath(Dir) = Dir) then Exit;
  Result := FBForceDirectories(FBExtractFilePath(Dir)) and FBDirectoryCreate(Dir);
end;

function FBDirectoryRemove(const FileName : string): boolean;

//  ___________________________________________________________________________
// | function FBDirectoryRemove                                                |
// | _________________________________________________________________________ |
// || Permet de supprimer un répertoire, le répertoire doit être vide         ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName : string                                       ||
// ||         |   Nom du répertoire à supprimer                               ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |   Vrai en cas de réussite, faux le cas contraire voir         ||
// ||         |   GetLastError pour connaitre la raison                       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  boo : boolean;
begin
  SetPrivilege('SeRestorePrivilege',true,boo);
  result :=  RemoveDirectoryW(strCheminToUnicode(FileName));
  if not boo then SetPrivilege('SeRestorePrivilege',boo,boo);
end;

function FBDirectoryRemove(const FileName : string;booBroyage : boolean ): boolean;

//  ___________________________________________________________________________
// | function FBDirectoryRemove                                                |
// | _________________________________________________________________________ |
// || Permet de supprimer un répertoire, le répertoire doit être vide         ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName : string                                       ||
// ||         |   Nom du répertoire à supprimer                               ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : boolean                                              ||
// ||         |   Vrai en cas de réussite, faux le cas contraire voir         ||
// ||         |   GetLastError pour connaitre la raison                       ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  strtemp : string;
begin
  strtemp := FileName ;
  if booBroyage then
  begin
    if FBDirectoryExists(FBExtractFileDrive(FileName )+'\fb0000000') then
    begin
      FBDirectoryRemove (FBExtractFileDrive(FileName )+'\fb0000000');
    end;
    if FBFileRename(FileName ,FBExtractFileDrive(FileName )+'\fb0000000') then
    strtemp := FBExtractFileDrive(FileName )+'\fb0000000';
  end;
  result :=  FBDirectoryRemove(strtemp);
end;

function FBDirectoryRemoveR(const FileName : string): integer;

//  ___________________________________________________________________________
// | function FBDirectoryRemoveR                                               |
// | _________________________________________________________________________ |
// || Permet de supprimer un répertoire et tous ces sous répertoires si ils ne||
// || contiennent pas de fichiers                                             ||
// || cette fonction gére les noms de fichier > 255 Caractères                ||
// ||_________________________________________________________________________||
// || Entrées | const FileName : string                                       ||
// ||         |   Nom du répertoire à supprimer                               ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : integer                                              ||
// ||         |   1 en cas de réussite, -1 le cas contraire                   ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  FD: TWin32FindDataW;
  FH: THandle;
  boofind : boolean;
  strReporg : string;
  boook : integer;
begin
  boofind := true;
  strReporg := FileName ;
  boook := 1;
  if strReporg[length(strReporg)] <> '\' then strReporg := strReporg + '\';
  FH := FindFirstFilew ( strCheminToUnicode(strReporg + '*.*'), FD);
  if  FH <> INVALID_HANDLE_VALUE then
  while boofind do
  begin
    if ((FD.dwFileAttributes and fadirectory )= faDirectory ) and
       (string(FD.cFileName) <> '.') and (string(FD.cFileName) <>'..') then
    begin
      boook := FBDirectoryRemoveR(strReporg + fd.cfilename+ '\' );
    end;
    boofind := windows.FindNextFileW (FH,FD);
  end;
  windows.findclose(FH);
  if not RemoveDirectoryW(strCheminToUnicode(strReporg)) then boook := -1 ;
  result := boook;
end;

function FBChangeFileExt(const FileName, Extension: string): string;

//  ___________________________________________________________________________
// | function FBChangeFileExt                                                  |
// | _________________________________________________________________________ |
// || Permet de changer l'extension d'un fichier                              ||
// ||_________________________________________________________________________||
// || Entrées | const FileName : string                                       ||
// ||         |   Nom de fichier à modifier                                   ||
// ||         | Extension : string                                            ||
// ||         |   Nouvelle extention du fichier                               ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   Nouveau nom de fichier                                      ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  I: Integer;
begin
  i := length(FileName );
  while (i > 0)and (FileName[i]<>'.')and (FileName[i]<>'\')and (FileName[i]<>':') do dec(i);
  if (I = 0) or (FileName[I] <> '.') then I := MaxInt;
  Result := Copy(FileName, 1, I - 1) + Extension;
end;

function FBExtractFilePath(const FileName: string): string;

//  ___________________________________________________________________________
// | function FBExtractFilePath                                                |
// | _________________________________________________________________________ |
// || Permet d'extraire le lecteur et le répertoire d'un nom de fichier       ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom complet                                                 ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   Chemin du fichier incluant le '\' finale                    ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  I: Integer;
begin
  i := length(FileName );
  while (i > 0)and (FileName[i]<>'\')and (FileName[i]<>':')do dec(i);
  Result := Copy(FileName, 1, I);
end;

function FBExtractFileDir(const FileName: string): string;

//  ___________________________________________________________________________
// | function FBExtractFileDir                                                 |
// | _________________________________________________________________________ |
// || Permet d'extraire le lecteur et le répertoire d'un nom de fichier       ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom complet                                                 ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   Chemin du fichier sans le '\' finale                        ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  I: Integer;
begin
  i := length(FileName );
  while (i > 0)and (FileName[i]<>'\')and (FileName[i]<>':')do dec(i);
  if (I > 1) and (FileName[I] = '\') and (FileName[I-1] <> ':')and(FileName[I-1] <> '\') then Dec(I);
  Result := Copy(FileName, 1, I);
end;

function FBExtractFileDrive(const FileName: string): string;

//  ___________________________________________________________________________
// | function FBExtractFileDrive                                               |
// | _________________________________________________________________________ |
// || Permet d'extraire la partie lecteur d'un nom de fichier                 ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom complet                                                 ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   Partie lecteur 'c:' ou '\\Serveur\partage'                  ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  I, J: Integer;
begin
  if (Length(FileName) >= 2) and (FileName[2] = ':') then
    Result := Copy(FileName, 1, 2)
  else if (Length(FileName) >= 2) and (FileName[1] = '\') and
    (FileName[2] = '\') then
  begin
    J := 0;
    I := 3;
    While (I < Length(FileName)) and (J < 2) do
    begin
      if FileName[I] = '\' then Inc(J);
      if J < 2 then Inc(I);
    end;
    if FileName[I] = '\' then Dec(I);
    Result := Copy(FileName, 1, I);
  end else Result := '';
end;

function FBExtractFileName(const FileName: string): string;

//  ___________________________________________________________________________
// | function FBExtractFileName                                                |
// | _________________________________________________________________________ |
// || Permet d'extraire la partie nom et extention d'un nom de  fichier       ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom complet                                                 ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   nom et extention du fichier                                 ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  I: Integer;
begin
  i := length(FileName );
  while (i > 0)and (FileName[i]<>'\')and (FileName[i]<>':')do dec(i);
  Result := Copy(FileName, I + 1, MaxInt);
end;

function FBExtractFileExt(const FileName: string): string;

//  ___________________________________________________________________________
// | function FBExtractFileExt                                                 |
// | _________________________________________________________________________ |
// || Permet d'extraire la partie extentino d'un nom de fichier               ||
// ||_________________________________________________________________________||
// || Entrées | const FileName: string                                        ||
// ||         |   Nom complet                                                 ||
// ||_________|_______________________________________________________________||
// || Sorties | Result : string                                               ||
// ||         |   extention du fichier                                        ||
// ||_________|_______________________________________________________________||
// |___________________________________________________________________________|

var
  I: Integer;
begin
  i := length(FileName );
  while (i > 0)and (FileName[i]<>'.')and (FileName[i]<>'\')and (FileName[i]<>':')do dec(i);
  if (I > 0) and (FileName[I] = '.') then
    Result := Copy(FileName, I, MaxInt) else
    Result := '';
end;


end.
