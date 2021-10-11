unit uGestDroit;

interface

Uses Windows, AclApi, AccCtrl;

Const

  SECURITY_NULL_SID_AUTHORITY : _SID_IDENTIFIER_AUTHORITY = ( Value : (0,0,0,0,0,0));
  SECURITY_WORLD_SID_AUTHORITY : _SID_IDENTIFIER_AUTHORITY = ( Value : (0,0,0,0,0,1));
  SECURITY_LOCAL_SID_AUTHORITY : _SID_IDENTIFIER_AUTHORITY = ( Value : (0,0,0,0,0,2));
  SECURITY_CREATOR_SID_AUTHORITY : _SID_IDENTIFIER_AUTHORITY = ( Value : (0,0,0,0,0,3));
  SECURITY_NON_UNIQUE_AUTHORITY : _SID_IDENTIFIER_AUTHORITY = ( Value : (0,0,0,0,0,4));
  SECURITY_NT_AUTHORITY : _SID_IDENTIFIER_AUTHORITY = ( Value : (0,0,0,0,0,5));


  SECURITY_WORLD_RID :CARDINAL =$00000000;
  SECURITY_BUILTIN_DOMAIN_RID : CARDINAL = $00000020;
  DOMAIN_ALIAS_RID_ADMINS : CARDINAL = $00000220;
  DOMAIN_ALIAS_RID_USERS  : CARDINAL = $00000221;
  DOMAIN_ALIAS_RID_GUESTS  : CARDINAL = $00000222;


  STANDARD_RIGHTS_ALL : CARDINAL =  $001F0000;
  ACL_REVISION : CARDINAL = 2; // current revision;


const
  ACCESS_MIN_MS_ACE_TYPE    = ($0);
  {$EXTERNALSYM ACCESS_MIN_MS_ACE_TYPE}
  ACCESS_ALLOWED_ACE_TYPE   = ($0);
  {$EXTERNALSYM ACCESS_ALLOWED_ACE_TYPE}
  ACCESS_DENIED_ACE_TYPE    = ($1);
  {$EXTERNALSYM ACCESS_DENIED_ACE_TYPE}
  SYSTEM_AUDIT_ACE_TYPE     = ($2);
  {$EXTERNALSYM SYSTEM_AUDIT_ACE_TYPE}
  SYSTEM_ALARM_ACE_TYPE     = ($3);
  {$EXTERNALSYM SYSTEM_ALARM_ACE_TYPE}
  ACCESS_MAX_MS_V2_ACE_TYPE = ($3);
  {$EXTERNALSYM ACCESS_MAX_MS_V2_ACE_TYPE}

  ACCESS_ALLOWED_COMPOUND_ACE_TYPE = ($4);
  {$EXTERNALSYM ACCESS_ALLOWED_COMPOUND_ACE_TYPE}
  ACCESS_MAX_MS_V3_ACE_TYPE        = ($4);
  {$EXTERNALSYM ACCESS_MAX_MS_V3_ACE_TYPE}

  ACCESS_MIN_MS_OBJECT_ACE_TYPE  = ($5);
  {$EXTERNALSYM ACCESS_MIN_MS_OBJECT_ACE_TYPE}
  ACCESS_ALLOWED_OBJECT_ACE_TYPE = ($5);
  {$EXTERNALSYM ACCESS_ALLOWED_OBJECT_ACE_TYPE}
  ACCESS_DENIED_OBJECT_ACE_TYPE  = ($6);
  {$EXTERNALSYM ACCESS_DENIED_OBJECT_ACE_TYPE}
  SYSTEM_AUDIT_OBJECT_ACE_TYPE   = ($7);
  {$EXTERNALSYM SYSTEM_AUDIT_OBJECT_ACE_TYPE}
  SYSTEM_ALARM_OBJECT_ACE_TYPE   = ($8);
  {$EXTERNALSYM SYSTEM_ALARM_OBJECT_ACE_TYPE}
  ACCESS_MAX_MS_OBJECT_ACE_TYPE  = ($8);
  {$EXTERNALSYM ACCESS_MAX_MS_OBJECT_ACE_TYPE}

  ACCESS_MAX_MS_V4_ACE_TYPE = ($8);
  {$EXTERNALSYM ACCESS_MAX_MS_V4_ACE_TYPE}
  ACCESS_MAX_MS_ACE_TYPE    = ($8);
  {$EXTERNALSYM ACCESS_MAX_MS_ACE_TYPE}

  ACCESS_ALLOWED_CALLBACK_ACE_TYPE        = $9;
  {$EXTERNALSYM ACCESS_ALLOWED_CALLBACK_ACE_TYPE}
  ACCESS_DENIED_CALLBACK_ACE_TYPE         = $A;
  {$EXTERNALSYM ACCESS_DENIED_CALLBACK_ACE_TYPE}
  ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE = $B;
  {$EXTERNALSYM ACCESS_ALLOWED_CALLBACK_OBJECT_ACE_TYPE}
  ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE  = $C;
  {$EXTERNALSYM ACCESS_DENIED_CALLBACK_OBJECT_ACE_TYPE}
  SYSTEM_AUDIT_CALLBACK_ACE_TYPE          = $D;
  {$EXTERNALSYM SYSTEM_AUDIT_CALLBACK_ACE_TYPE}
  SYSTEM_ALARM_CALLBACK_ACE_TYPE          = $E;
  {$EXTERNALSYM SYSTEM_ALARM_CALLBACK_ACE_TYPE}
  SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE   = $F;
  {$EXTERNALSYM SYSTEM_AUDIT_CALLBACK_OBJECT_ACE_TYPE}
  SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE   = $10;
  {$EXTERNALSYM SYSTEM_ALARM_CALLBACK_OBJECT_ACE_TYPE}

  ACCESS_MAX_MS_V5_ACE_TYPE               = $10;
  {$EXTERNALSYM ACCESS_MAX_MS_V5_ACE_TYPE}

  SUCCESSFUL_ACCESS_ACE_FLAG = ($40);
  {$EXTERNALSYM SUCCESSFUL_ACCESS_ACE_FLAG}
  FAILED_ACCESS_ACE_FLAG     = ($80);
  {$EXTERNALSYM FAILED_ACCESS_ACE_FLAG}



Type

  PACE_HEADER = ^ACE_HEADER;
  _ACE_HEADER = record
    AceType: Byte;
    AceFlags: Byte;
    AceSize: Word;
  end;

  ACE_HEADER = _ACE_HEADER;

  TAceHeader = ACE_HEADER;
  PAceHeader = PACE_HEADER;


  //Access Allowed ACE
PAccessAllowedAce = ^TAccessAllowedAce;
_ACCESS_ALLOWED_ACE = record
    Header : ACE_HEADER;
    Mask : DWORD;
    SidStart : DWORD;
end;
TAccessAllowedAce = _ACCESS_ALLOWED_ACE;


Type
//=== ACL (Access Control List)==============================
  //Size information
  PACL_SIZE_INFORMATION = ^ACL_SIZE_INFORMATION;
  _ACL_SIZE_INFORMATION = record
       AceCount,
       AclBytesInUse,
       AclBytesFree : DWORD;
  end;
  ACL_SIZE_INFORMATION =_ACL_SIZE_INFORMATION;
  TAclSizeInformation = ACL_SIZE_INFORMATION;
  PAclSizeInformation = PACL_SIZE_INFORMATION;

  //Revision Information
  PACL_REVISION_INFORMATION = ^ACL_REVISION_INFORMATION;
  _ACL_REVISION_INFORMATION  = record
       AclRevision : DWORD ;
  end;
  ACL_REVISION_INFORMATION = _ACL_REVISION_INFORMATION;
  TAclRevisionInformation = ACL_REVISION_INFORMATION;
  PAclRevisionInformation = PACL_REVISION_INFORMATION;


function IsAdmin : Boolean; stdcall; //is logged user is member of admins or
                                     //domain admins
function IsNT : Boolean; stdcall;    //is system NT based
function IsNT4 : Boolean; stdcall;   //is system NT 4
function GetEveryOneSid : Pointer; stdcall; //Security identifier of well known
                                            // group Everyone
function GetAccountSID( anAccountName : String) : Pointer; stdcall; overload ;
function GetAccountSid( AccountName: String; var SID: PSID): boolean; stdcall; overload;

function SetFileObjectAccessRights(aFileObject: String;
  aSID: Pointer; anAccess: CARDINAL; isInheritedAccess : BOOLEAN): BOOLEAN; stdcall;

function SetFileObjectAndSubobjectsAccessRights(aFileObject: String;
  aSID: Pointer; anAccess: CARDINAL): BOOLEAN; stdcall;

function SetEveryoneRWEDAccessToFileOrFolder( aFileOrFolder : String) : BOOLEAN; stdcall;
function SetEveryoneRWEDAccessToFileOrFolderAndSubobjects( aFileOrFolder : String) : BOOLEAN; stdcall;

function VolumeSupportsPersistentACLs( aPath : String) : Boolean; stdcall;

implementation
Uses
  SysUtils, ComObj;

function VolumeSupportsPersistentACLs( aPath : String) : Boolean;
Var
  maxClen,
  driveFlags : Cardinal;
  i : Integer;
  VolName, FSysName : Array[0..MAX_PATH] of Char;
begin
  aPath := ExtractFileDrive( aPath) +'\';
  Result := FALSE;
  if GetVolumeInformation(
     PChar( aPath),
     VolName,
     SizeOf( VolName),
     nil,
     maxClen,
     driveFlags,
     FsysName,
     SizeOf( FsysName)
  ) then
    Result := driveFlags and FS_PERSISTENT_ACLS = FS_PERSISTENT_ACLS ;
end;


function CheckCARDINALRslt( aCardinal : DWORD) : Boolean;
begin
  Result := aCardinal = ERROR_SUCCESS;
  if not Result then SetLastError( aCardinal);
end;


function IsAdmin : Boolean;
Var
  ntauth : SID_IDENTIFIER_AUTHORITY;
  psidAdmin : Pointer;
  bIsAdmin : Boolean;
  htok : THandle;
  cb : DWORD;
  ptg : ^TOKEN_GROUPS;
    i : Integer;
  grp : PSIDAndAttributes;
begin
  Result := FALSE;
  if not IsNT then
    Result := TRUE
  else
    begin
      bIsAdmin := FALSE;
      ntauth := SECURITY_NT_AUTHORITY;
      psidAdmin := nil;
      AllocateAndInitializeSid(
        ntauth, 2,
        SECURITY_BUILTIN_DOMAIN_RID,
        DOMAIN_ALIAS_RID_ADMINS,
        0, 0, 0, 0, 0, 0, psidAdmin
      );

      htok := 0;
      OpenProcessToken( GetCurrentProcess(),  TOKEN_QUERY, htok );
      GetTokenInformation( htok, TokenGroups, nil, 0, cb );
      GetMem( ptg, cb);
      GetTokenInformation( htok, TokenGroups,  ptg, cb, cb );

      grp := @(ptg.Groups[0]);


      for i := 0 to ptg.GroupCount-1 do
      begin
        if  EqualSid( psidAdmin, grp.Sid ) then
          begin
            bIsAdmin := TRUE;
            Break;
          end;
        Inc( grp) ; //, SizeOf( TSIDAndAttributes));
      end;
      freemem( ptg);
      CloseHandle(htok);
      FreeSid( psidAdmin);
      Result := bIsAdmin;
    end; // else of : if not IsNT
end;

function IsNT : Boolean;
Var
  ovi : TOSVersionInfo;
begin
  FillChar( ovi, SizeOf( Ovi), 0);
  ovi.dwOSVersionInfoSize := SizeOf( Ovi);
  GetVersionEx( ovi);
  Result := ovi.dwPlatformId = VER_PLATFORM_WIN32_NT;
end;

function IsNT4 : Boolean;
Var
  ovi : TOSVersionInfo;
begin
  FillChar( ovi, SizeOf( Ovi), 0);
  ovi.dwOSVersionInfoSize := SizeOf( Ovi);
  GetVersionEx( ovi);
  Result := (ovi.dwPlatformId = VER_PLATFORM_WIN32_NT) and (ovi.dwMajorVersion = 4);
end;

function GetEveryOneSid : Pointer;
begin
  AllocateAndInitializeSid(
       SECURITY_WORLD_SID_AUTHORITY,
       1,
       SECURITY_WORLD_RID,
       0,
       0, 0, 0, 0, 0, 0,
       Result
     );
end;



function GetAccountSID( anAccountName : String) : Pointer;
Var
  cb : CARDINAL;
  refDomainName : Array[0..1024] of Char;
  cbRefDomainName : Cardinal;
  peUse : Cardinal;
  SD : Pointer;
begin
  SD := NIL;
  try
    cbRefDomainName := SizeOf(refDomainName);
    FillChar( refDomainName, cbRefDomainName, 0);
    cb := 0;
    LookupAccountName(nil, PChar( anAccountName), nil, cb, refDomainName, cbRefDomainName, peUse);
    if cb > 0 then
      begin
        GetMem( SD, cb);
        FillChar( SD^, cb, 0);
        if not LookupAccountName(nil, PChar( anAccountName), SD, cb, refDomainName, cbRefDomainName, peUse) then
        begin
          FreeMem( SD, cb);
          SD := NIL;
        end;
      end
    else
      begin
        SD := NIL;
      end;
   finally
     Result := SD;
   end;
end;

function GetAccountSid(AccountName: String; var SID: PSID): boolean; stdcall;
var
  ReferencedDomain: pointer;
  cbSid: dword;
  cchReferencedDomain: dword; // initial allocation size
  peUse: SID_NAME_USE;
  bSuccess: BOOL; // assume this function will fail
  waccountname : Pwidechar;
begin
Result:=FALSE;
bSuccess:=FALSE;
SID:=nil;
ReferencedDomain:=nil;
cbSid:=128;
cchReferencedDomain:=16;
waccountname := PWideChar(WideString(AccountName ));
try
//. initial memory allocations
SID:=HeapAlloc(GetProcessHeap(),0,cbSid);
if SID = nil then Exit; //. ->
ReferencedDomain:=HeapAlloc(GetProcessHeap(),0,cchReferencedDomain*2);
if ReferencedDomain = nil then Exit; //. ->
//. Obtain the SID of the specified account on the specified system.
while NOT LookupAccountNameW(
        nil, // machine to lookup account on
        waccountname,        // account to lookup
        Sid,                // SID of interest
        cbSid,              // size of SID
        ReferencedDomain,   // domain account was found on
        cchReferencedDomain,
        peUse
        ) do begin

  if GetLastError() = ERROR_INSUFFICIENT_BUFFER
   then begin
    //. reallocate memory
    Sid:=HeapReAlloc(GetProcessHeap(),0,Sid,cbSid);
    if SID = nil then Exit; //. ->;
    ReferencedDomain:=HeapReAlloc(GetProcessHeap(),0,ReferencedDomain,cchReferencedDomain*2);
    if ReferencedDomain = nil then Exit; //. ->
    end
   else
    Exit; //. ->
  end;
//. Indicate success.
bSuccess:=TRUE;
finally
//. Cleanup and indicate failure, if appropriate.
HeapFree(GetProcessHeap(), 0, ReferencedDomain);
end;
if NOT bSuccess AND (SID <> nil)
 then begin
  HeapFree(GetProcessHeap(), 0, Sid);
  SID:=nil;
  end;
Result:=bSuccess;
end;



function SetFileObjectAndSubobjectsAccessRights(aFileObject: String;
  aSID: Pointer; anAccess: CARDINAL): BOOLEAN;
  function RecursiveSet( aPath : String) : Boolean;
  Var
    F : TSearchRec;
    i : Integer;
  begin
    Result := SetFileObjectAccessRights(aPath, aSID, anAccess, TRUE);
    i := FindFirst(aPath + '\*.*', faAnyFile, F);
    try
      While i = 0 do
      begin
        if (F.Name <> '') and (F.Name[1] <> '.') then
        begin
          if F.Attr and faDirectory = faDirectory then
             Result := Result and RecursiveSet( aPath + '\'+F.Name)
          else
             Result := Result and SetFileObjectAccessRights(aPath + '\' + F.Name, aSID, anAccess, TRUE);
          if not Result then Exit;
        end;
        i := FindNext( F);
      end;
    finally
      FindClose(F);
    end;
  end;

begin
  Result := FALSE;
  aFileObject := TRIM( aFileObject);
  if aFileObject <> '' then
  begin
    if DirectoryExists(aFileObject) then
      begin
        if aFileObject[Length( aFileObject)] = '\' then Delete( aFileObject, Length( aFileObject), 1);
        Result := RecursiveSet(aFileObject);
        Result := Result and SetFileObjectAccessRights(aFileObject, aSID, anAccess, FALSE);
      end
    else
      Result := SetFileObjectAccessRights(aFileObject, aSID, anAccess, FALSE);
  end;
end;

function SetEveryoneRWEDAccessToFileOrFolder( aFileOrFolder : String) : BOOLEAN;
Var
  SID : Pointer;
begin
  Result := FALSE;
  AllocateAndInitializeSid(
       SECURITY_WORLD_SID_AUTHORITY,
       1,
       SECURITY_WORLD_RID,
       0,
       0, 0, 0, 0, 0, 0,
       SID
     );
  if IsValidSid(SID) then
  try
    Result := SetFileObjectAccessRights(aFileOrFolder,
                                        SID,
                                        GENERIC_READ + GENERIC_WRITE + GENERIC_EXECUTE + _DELETE,
                                        FALSE
                                        );
  finally
    FreeSid(SID);
  end;
end;

function SetEveryoneRWEDAccessToFileOrFolderAndSubobjects( aFileOrFolder : String) : BOOLEAN;
Var
  SID : Pointer;
begin
  SID := GetEveryOneSid;
  try
    Result := SetFileObjectAndSubobjectsAccessRights(aFileOrFolder,
                                                     SID,
                                                     GENERIC_READ + GENERIC_WRITE + GENERIC_EXECUTE + _DELETE
                                                     );
  finally
    FreeSid(SID);
  end;
end;

function SetFileObjectAccessRights(aFileObject: String;
  aSID: Pointer; anAccess: CARDINAL; isInheritedAccess : BOOLEAN ): BOOLEAN;
var
  PPACL, PPACL2: PACL;
  newDacl: PACL;
  SecDescPtr, SD2: PSECURITY_DESCRIPTOR;
  needed : Cardinal;
  SD_Control : WORD;
  SD_Revision : Cardinal;
  aTrustee : TRUSTEE;
  expAccess : PExplicit_Access;
  isFile : Boolean;
  CurACEBr, CurACEInd : CARDINAL;
  OldAclSI : TAclSizeInformation;
  OldAclRI : TAclRevisionInformation;
  anACE : PAccessAllowedAce;
  i : Integer;
  oldACLSize, newACLSize, newACESize : Cardinal;
  bPresent, bDefaulted : LongBool;
begin
  Result := false;
  if not IsValidSid( aSID) then Exit;

  isFile := FileExists(aFileObject);

  PPACL := nil;
  if not IsNT4 then
    begin
      if not CheckCardinalRslt(
         GetNamedSecurityInfo(PChar(aFileObject), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION,nil, nil, PACL(@PPACL), nil, SecDescPtr)
         ) then Exit;
    end
  else
    begin
      GetFileSecurity(PChar(aFileObject), DACL_SECURITY_INFORMATION, nil, 0, needed);
      GetMem( SecDescPtr, needed);
      FillChar( SecDescPtr^, needed, 0);
      if not GetFileSecurity(PChar(aFileObject), DACL_SECURITY_INFORMATION, SecDescPtr, needed, needed) then  Exit;
      if not GetSecurityDescriptorDacl(SecDescPtr, bPresent, PPACL, bDefaulted) then Exit;
    end;
  try
    if not Assigned( PPACL) then Exit;
    if not GetSecurityDescriptorControl(SecDescPtr, SD_Control, SD_Revision) then Exit;
    if SD_Control and SE_DACL_PRESENT <> SE_DACL_PRESENT then Exit;

    if not GetAclInformation(PPACL^, @oldAclSI, SizeOF( TAclSizeInformation), AclSizeInformation) then  Exit;
    if not GetAclInformation(PPACL^,@oldAclRI, SizeOf( TAclRevisionInformation), AclRevisionInformation) then  Exit;


    //Delete previous ACE, for a given aSID
    CurACEBr := oldAclSI.AceCount;
    For i := oldAclSI.AceCount-1 downto 0 do
    begin
      if GetAce(PPACL^, i, Pointer(anAce)) then
      begin
        if  EqualSID( @(anACE.SidStart), aSID) then
        begin
          DeleteAce(PPACL^, i);
          CurACEBr := CurACEBr -1;
        end;
      end;
    end;

    if not GetAclInformation(PPACL^, @oldAclSI, SizeOF( TAclSizeInformation), AclSizeInformation) then  Exit;
    if not GetAclInformation(PPACL^,@oldAclRI, SizeOf( TAclRevisionInformation), AclRevisionInformation) then  Exit;

    NewACESize := SizeOf(TAccessAllowedACE) + GetLengthSid(aSID) - SizeOf( DWORD);
    OldACLSize := oldAclSI.AclBytesInUse + oldAclSI.AclBytesFree;
    NewACLSize := oldAclSI.AclBytesInUse + NewAceSize * 2 - oldAclSI.AclBytesFree;

    if NewAclSize < OldAclSize then NewAclSize := OldAclSize;

    GetMem( PPACL2, NewACLSize);
    try
      FillChar( PPACL2^, NewACLSize, 0);

      Move( PPACL^, PPACL2^, oldACLSize);
      PPACL2.AclSize := newACLSize;

      if not GetAclInformation(PPACL2^, @oldAclSI, SizeOF( TAclSizeInformation), AclSizeInformation) then  Exit;

      CurACEInd := 0;

      if not IsNT4 then
        begin
          //Construct Our Ace
          GetMem( anACE, newACESize);
          try
            FillChar( anACE^, newACESize, 0);
            anACE.Header.AceType := ACCESS_ALLOWED_ACE_TYPE;

            if not isFile then //demek e folder
            begin
              anACE.Header.AceFlags := SUB_CONTAINERS_ONLY_INHERIT + SUB_OBJECTS_ONLY_INHERIT;
            end;
            if isInheritedAccess then
            begin
              if not IsNt4 then
                anACE.Header.AceFlags := anACE.Header.AceFlags + INHERITED_ACCESS_ENTRY;
            end;


            anACE.Header.AceSize := newACESize;
            anAce.Mask := anAccess;
            Move( aSID^, anAce.SidStart, GetLengthSid(aSID));

            if not AddAce( PPACL2^, OldAclRI.AclRevision, CurACEInd, anACE, newACESize) then Exit;
          finally
            FreeMem( anACE, newACESize);
          end;
        end
      else
        begin
          CurACEInd := 0;
          if not isFile then
          begin
            GetMem( anACE, newACESize);
            try
              FillChar( anACE^, newACESize, 0);
              anACE.Header.AceType := ACCESS_ALLOWED_ACE_TYPE;

              anACE.Header.AceFlags := SUB_CONTAINERS_ONLY_INHERIT + SUB_OBJECTS_ONLY_INHERIT + INHERIT_ONLY;

              anACE.Header.AceSize := newACESize;
              anAce.Mask := anAccess;
              Move( aSID^, anAce.SidStart, GetLengthSid(aSID));

              if not AddAce( PPACL2^, OldAclRI.AclRevision, CurACEInd, anACE, newACESize) then Exit;
            finally
              FreeMem( anACE, newACESize);
            end;
          end;

          //Add ACE for Files
          GetMem( anACE, newACESize);
          try
            FillChar( anACE^, newACESize, 0);
            anACE.Header.AceType := ACCESS_ALLOWED_ACE_TYPE;

            anACE.Header.AceFlags := 0; // Empty flags, but ACE
            anACE.Header.AceSize := newACESize;
            anAce.Mask := anAccess;
            Move( aSID^, anAce.SidStart, GetLengthSid(aSID));

            if not AddAce( PPACL2^, OldAclRI.AclRevision, CurACEInd, anACE, newACESize) then Exit;
          finally
            FreeMem( anACE, newACESize);
          end;
        end;

      if not GetAclInformation(PPACL2^, @oldAclSI, SizeOF( TAclSizeInformation), AclSizeInformation) then  Exit;

      if not IsNT4 then
        begin
          Result := CheckCARDINALRslt(
            SetNamedSecurityInfo(PChar(aFileObject), SE_FILE_OBJECT, DACL_SECURITY_INFORMATION, nil, nil, PPACL2, nil)
          );
        end
      else
        begin
          GetMem( SD2, SizeOf( TSecurityDescriptor));
          try
            if not InitializeSecurityDescriptor(SD2, SECURITY_DESCRIPTOR_REVISION) then Exit;
            if not SetSecurityDescriptorDacl(SD2, bPresent, PPACL2, bDefaulted) then Exit;
            Result := SetFileSecurity(PChar(aFileObject), DACL_SECURITY_INFORMATION, SD2);
          finally
            FreeMem( SD2, SizeOf( TSecurityDescriptor));
          end;
        end;
    finally
       FreeMem( PPACL2, NewACLSize);
    end;
  finally
    LocalFree(HLOCAL(SecDescPtr));
  end;
end;

end.
