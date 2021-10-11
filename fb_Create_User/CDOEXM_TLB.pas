unit CDOEXM_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 01/03/2010 11:18:10 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Program Files\Exchsrvr\bin\cdoexm.dll (1)
// LIBID: {25150F00-5734-11D2-A593-00C04F990D8A}
// LCID: 0
// Helpfile: 
// HelpString: Microsoft CDO for Exchange Management Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
//   (2) v1.0 CDO, (C:\WINDOWS\system32\cdosys.dll)
//   (3) v2.5 ADODB, (C:\Program Files\Common Files\System\ado\msado25.tlb)
// Errors:
//   Hint: Parameter 'Interface' of IExchangeServer.GetInterface changed to 'Interface_'
//   Hint: Parameter 'Interface' of IFolderTree.GetInterface changed to 'Interface_'
//   Hint: Parameter 'Interface' of IPublicStoreDB.GetInterface changed to 'Interface_'
//   Hint: Parameter 'Interface' of IMailboxStoreDB.GetInterface changed to 'Interface_'
//   Hint: Parameter 'Interface' of IStorageGroup.GetInterface changed to 'Interface_'
//   Error creating palette bitmap of (TExchangeServer) : Indice de liste hors limites (12)
//   Error creating palette bitmap of (TFolderTree) : Indice de liste hors limites (16)
//   Error creating palette bitmap of (TPublicStoreDB) : Indice de liste hors limites (19)
//   Error creating palette bitmap of (TMailboxStoreDB) : Indice de liste hors limites (23)
//   Error creating palette bitmap of (TStorageGroup) : Indice de liste hors limites (26)
// ************************************************************************ //
// *************************************************************************//
// NOTE:                                                                      
// Items guarded by $IFDEF_LIVE_SERVER_AT_DESIGN_TIME are used by properties  
// which return objects that may need to be explicitly created via a function 
// call prior to any access via the property. These items have been disabled  
// in order to prevent accidental use from within the object inspector. You   
// may enable them by defining LIVE_SERVER_AT_DESIGN_TIME or by selectively   
// removing them from the $IFDEF blocks. However, such items must still be    
// programmatically created via a method of the appropriate CoClass before    
// they can be used.                                                          
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, ADODB_TLB, CDO_TLB, Classes, Graphics, OleServer, StdVCL, 
Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  CDOEXMMajorVersion = 6;
  CDOEXMMinorVersion = 5;

  LIBID_CDOEXM: TGUID = '{25150F00-5734-11D2-A593-00C04F990D8A}';

  IID_IDistributionList: TGUID = '{25150F3F-5734-11D2-A593-00C04F990D8A}';
  IID_IMailRecipient: TGUID = '{25150F40-5734-11D2-A593-00C04F990D8A}';
  IID_IMailboxStore: TGUID = '{25150F41-5734-11D2-A593-00C04F990D8A}';
  IID_IExchangeMailbox: TGUID = '{25150F51-5734-11D2-A593-00C04F990D8A}';
  IID_ISupportErrorInfo: TGUID = '{DF0B3D60-548F-101B-8E65-08002B2BD119}';
  CLASS_MailGroup: TGUID = '{25150F1F-5734-11D2-A593-00C04F990D8A}';
  CLASS_MailRecipient: TGUID = '{25150F20-5734-11D2-A593-00C04F990D8A}';
  CLASS_Mailbox: TGUID = '{25150F21-5734-11D2-A593-00C04F990D8A}';
  CLASS_FolderAdmin: TGUID = '{25150F22-5734-11D2-A593-00C04F990D8A}';
  IID_IExchangeServer: TGUID = '{25150F47-5734-11D2-A593-00C04F990D8A}';
  IID_IDataSource2: TGUID = '{25150F48-5734-11D2-A593-00C04F990D8A}';
  CLASS_ExchangeServer: TGUID = '{25150F27-5734-11D2-A593-00C04F990D8A}';
  IID_IFolderTree: TGUID = '{25150F43-5734-11D2-A593-00C04F990D8A}';
  CLASS_FolderTree: TGUID = '{25150F23-5734-11D2-A593-00C04F990D8A}';
  IID_IPublicStoreDB: TGUID = '{25150F44-5734-11D2-A593-00C04F990D8A}';
  IID_IPublicStoreDB2: TGUID = '{25150F49-5734-11D2-A593-00C04F990D8A}';
  CLASS_PublicStoreDB: TGUID = '{25150F24-5734-11D2-A593-00C04F990D8A}';
  IID_IMailboxStoreDB: TGUID = '{25150F45-5734-11D2-A593-00C04F990D8A}';
  IID_IMailboxStoreDB2: TGUID = '{25150F4A-5734-11D2-A593-00C04F990D8A}';
  CLASS_MailboxStoreDB: TGUID = '{25150F25-5734-11D2-A593-00C04F990D8A}';
  IID_IStorageGroup: TGUID = '{25150F46-5734-11D2-A593-00C04F990D8A}';
  CLASS_StorageGroup: TGUID = '{25150F26-5734-11D2-A593-00C04F990D8A}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum CDOEXMErrorCodes
type
  CDOEXMErrorCodes = TOleEnum;
const
  cdoexmRecoveryObjectsNotSupported = $C1032221;

// Constants for enum CDOEXMRestrictedAddressType
type
  CDOEXMRestrictedAddressType = TOleEnum;
const
  cdoexmAccept = $00000000;
  cdoexmReject = $00000001;

// Constants for enum CDOEXMDeliverAndRedirect
type
  CDOEXMDeliverAndRedirect = TOleEnum;
const
  cdoexmRecipientOrForward = $00000000;
  cdoexmDeliverToBoth = $00000001;

// Constants for enum CDOEXMServerType
type
  CDOEXMServerType = TOleEnum;
const
  cdoexmBackEnd = $00000000;
  cdoexmFrontEnd = $00000001;

// Constants for enum CDOEXMFolderTreeType
type
  CDOEXMFolderTreeType = TOleEnum;
const
  cdoexmGeneralPurpose = $00000000;
  cdoexmMAPI = $00000001;
  cdoexmNNTPOnly = $00000002;

// Constants for enum CDOEXMStoreDBStatus
type
  CDOEXMStoreDBStatus = TOleEnum;
const
  cdoexmOnline = $00000000;
  cdoexmOffline = $00000001;
  cdoexmMounting = $00000002;
  cdoexmDismounting = $00000003;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IDistributionList = interface;
  IDistributionListDisp = dispinterface;
  IMailRecipient = interface;
  IMailRecipientDisp = dispinterface;
  IMailboxStore = interface;
  IMailboxStoreDisp = dispinterface;
  IExchangeMailbox = interface;
  IExchangeMailboxDisp = dispinterface;
  ISupportErrorInfo = interface;
  IExchangeServer = interface;
  IExchangeServerDisp = dispinterface;
  IDataSource2 = interface;
  IDataSource2Disp = dispinterface;
  IFolderTree = interface;
  IFolderTreeDisp = dispinterface;
  IPublicStoreDB = interface;
  IPublicStoreDBDisp = dispinterface;
  IPublicStoreDB2 = interface;
  IPublicStoreDB2Disp = dispinterface;
  IMailboxStoreDB = interface;
  IMailboxStoreDBDisp = dispinterface;
  IMailboxStoreDB2 = interface;
  IMailboxStoreDB2Disp = dispinterface;
  IStorageGroup = interface;
  IStorageGroupDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  MailGroup = IMailRecipient;
  MailRecipient = IMailRecipient;
  Mailbox = IMailRecipient;
  FolderAdmin = IMailRecipient;
  ExchangeServer = IExchangeServer;
  FolderTree = IFolderTree;
  PublicStoreDB = IPublicStoreDB2;
  MailboxStoreDB = IMailboxStoreDB2;
  StorageGroup = IStorageGroup;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PUserType1 = ^TGUID; {*}


// *********************************************************************//
// Interface: IDistributionList
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F3F-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IDistributionList = interface(IDispatch)
    ['{25150F3F-5734-11D2-A593-00C04F990D8A}']
    function Get_HideDLMembership: WordBool; safecall;
    procedure Set_HideDLMembership(pHideDLMembership: WordBool); safecall;
    property HideDLMembership: WordBool read Get_HideDLMembership write Set_HideDLMembership;
  end;

// *********************************************************************//
// DispIntf:  IDistributionListDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F3F-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IDistributionListDisp = dispinterface
    ['{25150F3F-5734-11D2-A593-00C04F990D8A}']
    property HideDLMembership: WordBool dispid 49;
  end;

// *********************************************************************//
// Interface: IMailRecipient
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F40-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailRecipient = interface(IDispatch)
    ['{25150F40-5734-11D2-A593-00C04F990D8A}']
    function Get_IncomingLimit: Integer; safecall;
    procedure Set_IncomingLimit(pIncomingLimit: Integer); safecall;
    function Get_OutgoingLimit: Integer; safecall;
    procedure Set_OutgoingLimit(pOutgoingLimit: Integer); safecall;
    function Get_RestrictedAddressList: OleVariant; safecall;
    procedure Set_RestrictedAddressList(pRestrictedAddressList: OleVariant); safecall;
    function Get_RestrictedAddresses: CDOEXMRestrictedAddressType; safecall;
    procedure Set_RestrictedAddresses(pRestrictedAddresses: CDOEXMRestrictedAddressType); safecall;
    function Get_ForwardTo: WideString; safecall;
    procedure Set_ForwardTo(const pForwardTo: WideString); safecall;
    function Get_ForwardingStyle: CDOEXMDeliverAndRedirect; safecall;
    procedure Set_ForwardingStyle(pForwardingStyle: CDOEXMDeliverAndRedirect); safecall;
    function Get_HideFromAddressBook: WordBool; safecall;
    procedure Set_HideFromAddressBook(pHideFromAddressBook: WordBool); safecall;
    function Get_X400Email: WideString; safecall;
    procedure Set_X400Email(const pX400Email: WideString); safecall;
    function Get_SMTPEmail: WideString; safecall;
    procedure Set_SMTPEmail(const pSMTPEmail: WideString); safecall;
    function Get_ProxyAddresses: OleVariant; safecall;
    procedure Set_ProxyAddresses(pProxyAddresses: OleVariant); safecall;
    function Get_AutoGenerateEmailAddresses: WordBool; safecall;
    procedure Set_AutoGenerateEmailAddresses(pAutoGenerateEmailAddresses: WordBool); safecall;
    function Get_Alias: WideString; safecall;
    procedure Set_Alias(const pAlias: WideString); safecall;
    function Get_TargetAddress: WideString; safecall;
    procedure MailEnable(const TargetMailAddress: WideString); safecall;
    procedure MailDisable; safecall;
    property IncomingLimit: Integer read Get_IncomingLimit write Set_IncomingLimit;
    property OutgoingLimit: Integer read Get_OutgoingLimit write Set_OutgoingLimit;
    property RestrictedAddressList: OleVariant read Get_RestrictedAddressList write Set_RestrictedAddressList;
    property RestrictedAddresses: CDOEXMRestrictedAddressType read Get_RestrictedAddresses write Set_RestrictedAddresses;
    property ForwardTo: WideString read Get_ForwardTo write Set_ForwardTo;
    property ForwardingStyle: CDOEXMDeliverAndRedirect read Get_ForwardingStyle write Set_ForwardingStyle;
    property HideFromAddressBook: WordBool read Get_HideFromAddressBook write Set_HideFromAddressBook;
    property X400Email: WideString read Get_X400Email write Set_X400Email;
    property SMTPEmail: WideString read Get_SMTPEmail write Set_SMTPEmail;
    property ProxyAddresses: OleVariant read Get_ProxyAddresses write Set_ProxyAddresses;
    property AutoGenerateEmailAddresses: WordBool read Get_AutoGenerateEmailAddresses write Set_AutoGenerateEmailAddresses;
    property Alias: WideString read Get_Alias write Set_Alias;
    property TargetAddress: WideString read Get_TargetAddress;
  end;

// *********************************************************************//
// DispIntf:  IMailRecipientDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F40-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailRecipientDisp = dispinterface
    ['{25150F40-5734-11D2-A593-00C04F990D8A}']
    property IncomingLimit: Integer dispid 16;
    property OutgoingLimit: Integer dispid 17;
    property RestrictedAddressList: OleVariant dispid 18;
    property RestrictedAddresses: CDOEXMRestrictedAddressType dispid 19;
    property ForwardTo: WideString dispid 20;
    property ForwardingStyle: CDOEXMDeliverAndRedirect dispid 21;
    property HideFromAddressBook: WordBool dispid 22;
    property X400Email: WideString dispid 23;
    property SMTPEmail: WideString dispid 24;
    property ProxyAddresses: OleVariant dispid 25;
    property AutoGenerateEmailAddresses: WordBool dispid 26;
    property Alias: WideString dispid 27;
    property TargetAddress: WideString readonly dispid 28;
    procedure MailEnable(const TargetMailAddress: WideString); dispid 29;
    procedure MailDisable; dispid 30;
  end;

// *********************************************************************//
// Interface: IMailboxStore
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F41-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailboxStore = interface(IDispatch)
    ['{25150F41-5734-11D2-A593-00C04F990D8A}']
    function Get_EnableStoreDefaults: OleVariant; safecall;
    procedure Set_EnableStoreDefaults(pEnableStoreDefaults: OleVariant); safecall;
    function Get_StoreQuota: Integer; safecall;
    procedure Set_StoreQuota(pStoreQuota: Integer); safecall;
    function Get_OverQuotaLimit: Integer; safecall;
    procedure Set_OverQuotaLimit(pOverQuotaLimit: Integer); safecall;
    function Get_HardLimit: Integer; safecall;
    procedure Set_HardLimit(pHardLimit: Integer); safecall;
    function Get_OverrideStoreGarbageCollection: WordBool; safecall;
    procedure Set_OverrideStoreGarbageCollection(pOverrideStoreGarbageCollection: WordBool); safecall;
    function Get_DaysBeforeGarbageCollection: Integer; safecall;
    procedure Set_DaysBeforeGarbageCollection(pDaysBeforeGarbageCollection: Integer); safecall;
    function Get_GarbageCollectOnlyAfterBackup: WordBool; safecall;
    procedure Set_GarbageCollectOnlyAfterBackup(pGarbageCollectOnlyAfterBackup: WordBool); safecall;
    function Get_Delegates: OleVariant; safecall;
    procedure Set_Delegates(pDelegates: OleVariant); safecall;
    function Get_HomeMDB: WideString; safecall;
    function Get_RecipientLimit: Integer; safecall;
    procedure Set_RecipientLimit(pRecipientLimit: Integer); safecall;
    procedure CreateMailbox(const HomeMDBURL: WideString); safecall;
    procedure DeleteMailbox; safecall;
    procedure MoveMailbox(const HomeMDBURL: WideString); safecall;
    property EnableStoreDefaults: OleVariant read Get_EnableStoreDefaults write Set_EnableStoreDefaults;
    property StoreQuota: Integer read Get_StoreQuota write Set_StoreQuota;
    property OverQuotaLimit: Integer read Get_OverQuotaLimit write Set_OverQuotaLimit;
    property HardLimit: Integer read Get_HardLimit write Set_HardLimit;
    property OverrideStoreGarbageCollection: WordBool read Get_OverrideStoreGarbageCollection write Set_OverrideStoreGarbageCollection;
    property DaysBeforeGarbageCollection: Integer read Get_DaysBeforeGarbageCollection write Set_DaysBeforeGarbageCollection;
    property GarbageCollectOnlyAfterBackup: WordBool read Get_GarbageCollectOnlyAfterBackup write Set_GarbageCollectOnlyAfterBackup;
    property Delegates: OleVariant read Get_Delegates write Set_Delegates;
    property HomeMDB: WideString read Get_HomeMDB;
    property RecipientLimit: Integer read Get_RecipientLimit write Set_RecipientLimit;
  end;

// *********************************************************************//
// DispIntf:  IMailboxStoreDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F41-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailboxStoreDisp = dispinterface
    ['{25150F41-5734-11D2-A593-00C04F990D8A}']
    property EnableStoreDefaults: OleVariant dispid 50;
    property StoreQuota: Integer dispid 51;
    property OverQuotaLimit: Integer dispid 52;
    property HardLimit: Integer dispid 53;
    property OverrideStoreGarbageCollection: WordBool dispid 54;
    property DaysBeforeGarbageCollection: Integer dispid 55;
    property GarbageCollectOnlyAfterBackup: WordBool dispid 56;
    property Delegates: OleVariant dispid 57;
    property HomeMDB: WideString readonly dispid 58;
    property RecipientLimit: Integer dispid 59;
    procedure CreateMailbox(const HomeMDBURL: WideString); dispid 60;
    procedure DeleteMailbox; dispid 61;
    procedure MoveMailbox(const HomeMDBURL: WideString); dispid 62;
  end;

// *********************************************************************//
// Interface: IExchangeMailbox
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F51-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IExchangeMailbox = interface(IMailboxStore)
    ['{25150F51-5734-11D2-A593-00C04F990D8A}']
    function Get_MailboxRights: OleVariant; safecall;
    procedure Set_MailboxRights(pMailboxRights: OleVariant); safecall;
    property MailboxRights: OleVariant read Get_MailboxRights write Set_MailboxRights;
  end;

// *********************************************************************//
// DispIntf:  IExchangeMailboxDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F51-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IExchangeMailboxDisp = dispinterface
    ['{25150F51-5734-11D2-A593-00C04F990D8A}']
    property MailboxRights: OleVariant dispid 63;
    property EnableStoreDefaults: OleVariant dispid 50;
    property StoreQuota: Integer dispid 51;
    property OverQuotaLimit: Integer dispid 52;
    property HardLimit: Integer dispid 53;
    property OverrideStoreGarbageCollection: WordBool dispid 54;
    property DaysBeforeGarbageCollection: Integer dispid 55;
    property GarbageCollectOnlyAfterBackup: WordBool dispid 56;
    property Delegates: OleVariant dispid 57;
    property HomeMDB: WideString readonly dispid 58;
    property RecipientLimit: Integer dispid 59;
    procedure CreateMailbox(const HomeMDBURL: WideString); dispid 60;
    procedure DeleteMailbox; dispid 61;
    procedure MoveMailbox(const HomeMDBURL: WideString); dispid 62;
  end;

// *********************************************************************//
// Interface: ISupportErrorInfo
// Flags:     (0)
// GUID:      {DF0B3D60-548F-101B-8E65-08002B2BD119}
// *********************************************************************//
  ISupportErrorInfo = interface(IUnknown)
    ['{DF0B3D60-548F-101B-8E65-08002B2BD119}']
    function InterfaceSupportsErrorInfo(var riid: TGUID): HResult; stdcall;
  end;

// *********************************************************************//
// Interface: IExchangeServer
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F47-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IExchangeServer = interface(IDispatch)
    ['{25150F47-5734-11D2-A593-00C04F990D8A}']
    function Get_DataSource: IDataSource2; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Name: WideString; safecall;
    function Get_ExchangeVersion: WideString; safecall;
    function Get_StorageGroups: OleVariant; safecall;
    function Get_SubjectLoggingEnabled: WordBool; safecall;
    procedure Set_SubjectLoggingEnabled(pSubjectLoggingEnabled: WordBool); safecall;
    function Get_MessageTrackingEnabled: WordBool; safecall;
    procedure Set_MessageTrackingEnabled(pMessageTrackingEnabled: WordBool); safecall;
    function Get_DaysBeforeLogFileRemoval: Integer; safecall;
    procedure Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval: Integer); safecall;
    function Get_ServerType: CDOEXMServerType; safecall;
    procedure Set_ServerType(pServerType: CDOEXMServerType); safecall;
    function Get_DirectoryServer: WideString; safecall;
    function GetInterface(const Interface_: WideString): IDispatch; safecall;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property Name: WideString read Get_Name;
    property ExchangeVersion: WideString read Get_ExchangeVersion;
    property StorageGroups: OleVariant read Get_StorageGroups;
    property SubjectLoggingEnabled: WordBool read Get_SubjectLoggingEnabled write Set_SubjectLoggingEnabled;
    property MessageTrackingEnabled: WordBool read Get_MessageTrackingEnabled write Set_MessageTrackingEnabled;
    property DaysBeforeLogFileRemoval: Integer read Get_DaysBeforeLogFileRemoval write Set_DaysBeforeLogFileRemoval;
    property ServerType: CDOEXMServerType read Get_ServerType write Set_ServerType;
    property DirectoryServer: WideString read Get_DirectoryServer;
  end;

// *********************************************************************//
// DispIntf:  IExchangeServerDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F47-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IExchangeServerDisp = dispinterface
    ['{25150F47-5734-11D2-A593-00C04F990D8A}']
    property DataSource: IDataSource2 readonly dispid 75;
    property Fields: Fields readonly dispid 76;
    property Name: WideString readonly dispid 77;
    property ExchangeVersion: WideString readonly dispid 78;
    property StorageGroups: OleVariant readonly dispid 79;
    property SubjectLoggingEnabled: WordBool dispid 80;
    property MessageTrackingEnabled: WordBool dispid 81;
    property DaysBeforeLogFileRemoval: Integer dispid 82;
    property ServerType: CDOEXMServerType dispid 83;
    property DirectoryServer: WideString readonly dispid 84;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 85;
  end;

// *********************************************************************//
// Interface: IDataSource2
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F48-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IDataSource2 = interface(IDataSource)
    ['{25150F48-5734-11D2-A593-00C04F990D8A}']
    procedure Delete; safecall;
    procedure MoveToContainer(const ContainerURL: WideString); safecall;
  end;

// *********************************************************************//
// DispIntf:  IDataSource2Disp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F48-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IDataSource2Disp = dispinterface
    ['{25150F48-5734-11D2-A593-00C04F990D8A}']
    procedure Delete; dispid 257;
    procedure MoveToContainer(const ContainerURL: WideString); dispid 258;
    property SourceClass: WideString readonly dispid 207;
    property Source: IUnknown readonly dispid 208;
    property IsDirty: WordBool dispid 209;
    property SourceURL: WideString readonly dispid 210;
    property ActiveConnection: _Connection readonly dispid 211;
    procedure SaveToObject(const Source: IUnknown; const InterfaceName: WideString); dispid 251;
    procedure OpenObject(const Source: IUnknown; const InterfaceName: WideString); dispid 252;
    procedure SaveTo(const SourceURL: WideString; const ActiveConnection: IDispatch; 
                     Mode: ConnectModeEnum; CreateOptions: RecordCreateOptionsEnum; 
                     Options: RecordOpenOptionsEnum; const UserName: WideString; 
                     const Password: WideString); dispid 253;
    procedure Open(const SourceURL: WideString; const ActiveConnection: IDispatch; 
                   Mode: ConnectModeEnum; CreateOptions: RecordCreateOptionsEnum; 
                   Options: RecordOpenOptionsEnum; const UserName: WideString; 
                   const Password: WideString); dispid 254;
    procedure Save; dispid 255;
    procedure SaveToContainer(const ContainerURL: WideString; const ActiveConnection: IDispatch; 
                              Mode: ConnectModeEnum; CreateOptions: RecordCreateOptionsEnum; 
                              Options: RecordOpenOptionsEnum; const UserName: WideString; 
                              const Password: WideString); dispid 256;
  end;

// *********************************************************************//
// Interface: IFolderTree
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F43-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IFolderTree = interface(IDispatch)
    ['{25150F43-5734-11D2-A593-00C04F990D8A}']
    function Get_DataSource: IDataSource2; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pName: WideString); safecall;
    function Get_StoreDBs: OleVariant; safecall;
    function Get_TreeType: CDOEXMFolderTreeType; safecall;
    function Get_RootFolderURL: WideString; safecall;
    function GetInterface(const Interface_: WideString): IDispatch; safecall;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property Name: WideString read Get_Name write Set_Name;
    property StoreDBs: OleVariant read Get_StoreDBs;
    property TreeType: CDOEXMFolderTreeType read Get_TreeType;
    property RootFolderURL: WideString read Get_RootFolderURL;
  end;

// *********************************************************************//
// DispIntf:  IFolderTreeDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F43-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IFolderTreeDisp = dispinterface
    ['{25150F43-5734-11D2-A593-00C04F990D8A}']
    property DataSource: IDataSource2 readonly dispid 175;
    property Fields: Fields readonly dispid 176;
    property Name: WideString dispid 177;
    property StoreDBs: OleVariant readonly dispid 178;
    property TreeType: CDOEXMFolderTreeType readonly dispid 180;
    property RootFolderURL: WideString readonly dispid 181;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 182;
  end;

// *********************************************************************//
// Interface: IPublicStoreDB
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F44-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IPublicStoreDB = interface(IDispatch)
    ['{25150F44-5734-11D2-A593-00C04F990D8A}']
    function Get_DataSource: IDataSource2; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pName: WideString); safecall;
    function Get_FolderTree: WideString; safecall;
    procedure Set_FolderTree(const pFolderTree: WideString); safecall;
    function Get_DBPath: WideString; safecall;
    function Get_SLVPath: WideString; safecall;
    function Get_Status: CDOEXMStoreDBStatus; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(pEnabled: WordBool); safecall;
    function Get_StoreQuota: Integer; safecall;
    procedure Set_StoreQuota(pStoreQuota: Integer); safecall;
    function Get_HardLimit: Integer; safecall;
    procedure Set_HardLimit(pHardLimit: Integer); safecall;
    function Get_ItemSizeLimit: Integer; safecall;
    procedure Set_ItemSizeLimit(pItemSizeLimit: Integer); safecall;
    function Get_DaysBeforeItemExpiration: Integer; safecall;
    procedure Set_DaysBeforeItemExpiration(pDaysBeforeItemExpiration: Integer); safecall;
    function Get_DaysBeforeGarbageCollection: Integer; safecall;
    procedure Set_DaysBeforeGarbageCollection(pDaysBeforeGarbageCollection: Integer); safecall;
    function Get_GarbageCollectOnlyAfterBackup: WordBool; safecall;
    procedure Set_GarbageCollectOnlyAfterBackup(pGarbageCollectOnlyAfterBackup: WordBool); safecall;
    function GetInterface(const Interface_: WideString): IDispatch; safecall;
    procedure MoveDataFiles(const DBPath: WideString; const SLVPath: WideString; Flags: Integer); safecall;
    procedure Mount(Timeout: Integer); safecall;
    procedure Dismount(Timeout: Integer); safecall;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property Name: WideString read Get_Name write Set_Name;
    property FolderTree: WideString read Get_FolderTree write Set_FolderTree;
    property DBPath: WideString read Get_DBPath;
    property SLVPath: WideString read Get_SLVPath;
    property Status: CDOEXMStoreDBStatus read Get_Status;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property StoreQuota: Integer read Get_StoreQuota write Set_StoreQuota;
    property HardLimit: Integer read Get_HardLimit write Set_HardLimit;
    property ItemSizeLimit: Integer read Get_ItemSizeLimit write Set_ItemSizeLimit;
    property DaysBeforeItemExpiration: Integer read Get_DaysBeforeItemExpiration write Set_DaysBeforeItemExpiration;
    property DaysBeforeGarbageCollection: Integer read Get_DaysBeforeGarbageCollection write Set_DaysBeforeGarbageCollection;
    property GarbageCollectOnlyAfterBackup: WordBool read Get_GarbageCollectOnlyAfterBackup write Set_GarbageCollectOnlyAfterBackup;
  end;

// *********************************************************************//
// DispIntf:  IPublicStoreDBDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F44-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IPublicStoreDBDisp = dispinterface
    ['{25150F44-5734-11D2-A593-00C04F990D8A}']
    property DataSource: IDataSource2 readonly dispid 125;
    property Fields: Fields readonly dispid 126;
    property Name: WideString dispid 127;
    property FolderTree: WideString dispid 128;
    property DBPath: WideString readonly dispid 129;
    property SLVPath: WideString readonly dispid 130;
    property Status: CDOEXMStoreDBStatus readonly dispid 131;
    property Enabled: WordBool dispid 132;
    property StoreQuota: Integer dispid 133;
    property HardLimit: Integer dispid 134;
    property ItemSizeLimit: Integer dispid 135;
    property DaysBeforeItemExpiration: Integer dispid 136;
    property DaysBeforeGarbageCollection: Integer dispid 137;
    property GarbageCollectOnlyAfterBackup: WordBool dispid 138;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 139;
    procedure MoveDataFiles(const DBPath: WideString; const SLVPath: WideString; Flags: Integer); dispid 140;
    procedure Mount(Timeout: Integer); dispid 141;
    procedure Dismount(Timeout: Integer); dispid 142;
  end;

// *********************************************************************//
// Interface: IPublicStoreDB2
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F49-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IPublicStoreDB2 = interface(IPublicStoreDB)
    ['{25150F49-5734-11D2-A593-00C04F990D8A}']
    function Get_LastFullBackupTime: OleVariant; safecall;
    function Get_LastIncrementalBackupTime: OleVariant; safecall;
    property LastFullBackupTime: OleVariant read Get_LastFullBackupTime;
    property LastIncrementalBackupTime: OleVariant read Get_LastIncrementalBackupTime;
  end;

// *********************************************************************//
// DispIntf:  IPublicStoreDB2Disp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F49-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IPublicStoreDB2Disp = dispinterface
    ['{25150F49-5734-11D2-A593-00C04F990D8A}']
    property LastFullBackupTime: OleVariant readonly dispid 143;
    property LastIncrementalBackupTime: OleVariant readonly dispid 144;
    property DataSource: IDataSource2 readonly dispid 125;
    property Fields: Fields readonly dispid 126;
    property Name: WideString dispid 127;
    property FolderTree: WideString dispid 128;
    property DBPath: WideString readonly dispid 129;
    property SLVPath: WideString readonly dispid 130;
    property Status: CDOEXMStoreDBStatus readonly dispid 131;
    property Enabled: WordBool dispid 132;
    property StoreQuota: Integer dispid 133;
    property HardLimit: Integer dispid 134;
    property ItemSizeLimit: Integer dispid 135;
    property DaysBeforeItemExpiration: Integer dispid 136;
    property DaysBeforeGarbageCollection: Integer dispid 137;
    property GarbageCollectOnlyAfterBackup: WordBool dispid 138;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 139;
    procedure MoveDataFiles(const DBPath: WideString; const SLVPath: WideString; Flags: Integer); dispid 140;
    procedure Mount(Timeout: Integer); dispid 141;
    procedure Dismount(Timeout: Integer); dispid 142;
  end;

// *********************************************************************//
// Interface: IMailboxStoreDB
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F45-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailboxStoreDB = interface(IDispatch)
    ['{25150F45-5734-11D2-A593-00C04F990D8A}']
    function Get_DataSource: IDataSource2; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pName: WideString); safecall;
    function Get_PublicStoreDB: WideString; safecall;
    procedure Set_PublicStoreDB(const pPublicStoreDB: WideString); safecall;
    function Get_OfflineAddressList: WideString; safecall;
    procedure Set_OfflineAddressList(const pOfflineAddressList: WideString); safecall;
    function Get_DBPath: WideString; safecall;
    function Get_SLVPath: WideString; safecall;
    function Get_Status: CDOEXMStoreDBStatus; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(pEnabled: WordBool); safecall;
    function Get_StoreQuota: Integer; safecall;
    procedure Set_StoreQuota(pStoreQuota: Integer); safecall;
    function Get_OverQuotaLimit: Integer; safecall;
    procedure Set_OverQuotaLimit(pOverQuotaLimit: Integer); safecall;
    function Get_HardLimit: Integer; safecall;
    procedure Set_HardLimit(pHardLimit: Integer); safecall;
    function Get_DaysBeforeGarbageCollection: Integer; safecall;
    procedure Set_DaysBeforeGarbageCollection(pDaysBeforeGarbageCollection: Integer); safecall;
    function Get_DaysBeforeDeletedMailboxCleanup: Integer; safecall;
    procedure Set_DaysBeforeDeletedMailboxCleanup(pDaysBeforeDeletedMailboxCleanup: Integer); safecall;
    function Get_GarbageCollectOnlyAfterBackup: WordBool; safecall;
    procedure Set_GarbageCollectOnlyAfterBackup(pGarbageCollectOnlyAfterBackup: WordBool); safecall;
    function GetInterface(const Interface_: WideString): IDispatch; safecall;
    procedure MoveDataFiles(const DBPath: WideString; const SLVPath: WideString; Flags: Integer); safecall;
    procedure Mount(Timeout: Integer); safecall;
    procedure Dismount(Timeout: Integer); safecall;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property Name: WideString read Get_Name write Set_Name;
    property PublicStoreDB: WideString read Get_PublicStoreDB write Set_PublicStoreDB;
    property OfflineAddressList: WideString read Get_OfflineAddressList write Set_OfflineAddressList;
    property DBPath: WideString read Get_DBPath;
    property SLVPath: WideString read Get_SLVPath;
    property Status: CDOEXMStoreDBStatus read Get_Status;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property StoreQuota: Integer read Get_StoreQuota write Set_StoreQuota;
    property OverQuotaLimit: Integer read Get_OverQuotaLimit write Set_OverQuotaLimit;
    property HardLimit: Integer read Get_HardLimit write Set_HardLimit;
    property DaysBeforeGarbageCollection: Integer read Get_DaysBeforeGarbageCollection write Set_DaysBeforeGarbageCollection;
    property DaysBeforeDeletedMailboxCleanup: Integer read Get_DaysBeforeDeletedMailboxCleanup write Set_DaysBeforeDeletedMailboxCleanup;
    property GarbageCollectOnlyAfterBackup: WordBool read Get_GarbageCollectOnlyAfterBackup write Set_GarbageCollectOnlyAfterBackup;
  end;

// *********************************************************************//
// DispIntf:  IMailboxStoreDBDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F45-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailboxStoreDBDisp = dispinterface
    ['{25150F45-5734-11D2-A593-00C04F990D8A}']
    property DataSource: IDataSource2 readonly dispid 150;
    property Fields: Fields readonly dispid 151;
    property Name: WideString dispid 152;
    property PublicStoreDB: WideString dispid 153;
    property OfflineAddressList: WideString dispid 154;
    property DBPath: WideString readonly dispid 155;
    property SLVPath: WideString readonly dispid 156;
    property Status: CDOEXMStoreDBStatus readonly dispid 157;
    property Enabled: WordBool dispid 158;
    property StoreQuota: Integer dispid 159;
    property OverQuotaLimit: Integer dispid 160;
    property HardLimit: Integer dispid 161;
    property DaysBeforeGarbageCollection: Integer dispid 162;
    property DaysBeforeDeletedMailboxCleanup: Integer dispid 163;
    property GarbageCollectOnlyAfterBackup: WordBool dispid 164;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 165;
    procedure MoveDataFiles(const DBPath: WideString; const SLVPath: WideString; Flags: Integer); dispid 166;
    procedure Mount(Timeout: Integer); dispid 167;
    procedure Dismount(Timeout: Integer); dispid 168;
  end;

// *********************************************************************//
// Interface: IMailboxStoreDB2
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F4A-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailboxStoreDB2 = interface(IMailboxStoreDB)
    ['{25150F4A-5734-11D2-A593-00C04F990D8A}']
    function Get_LastFullBackupTime: OleVariant; safecall;
    function Get_LastIncrementalBackupTime: OleVariant; safecall;
    property LastFullBackupTime: OleVariant read Get_LastFullBackupTime;
    property LastIncrementalBackupTime: OleVariant read Get_LastIncrementalBackupTime;
  end;

// *********************************************************************//
// DispIntf:  IMailboxStoreDB2Disp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F4A-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IMailboxStoreDB2Disp = dispinterface
    ['{25150F4A-5734-11D2-A593-00C04F990D8A}']
    property LastFullBackupTime: OleVariant readonly dispid 169;
    property LastIncrementalBackupTime: OleVariant readonly dispid 170;
    property DataSource: IDataSource2 readonly dispid 150;
    property Fields: Fields readonly dispid 151;
    property Name: WideString dispid 152;
    property PublicStoreDB: WideString dispid 153;
    property OfflineAddressList: WideString dispid 154;
    property DBPath: WideString readonly dispid 155;
    property SLVPath: WideString readonly dispid 156;
    property Status: CDOEXMStoreDBStatus readonly dispid 157;
    property Enabled: WordBool dispid 158;
    property StoreQuota: Integer dispid 159;
    property OverQuotaLimit: Integer dispid 160;
    property HardLimit: Integer dispid 161;
    property DaysBeforeGarbageCollection: Integer dispid 162;
    property DaysBeforeDeletedMailboxCleanup: Integer dispid 163;
    property GarbageCollectOnlyAfterBackup: WordBool dispid 164;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 165;
    procedure MoveDataFiles(const DBPath: WideString; const SLVPath: WideString; Flags: Integer); dispid 166;
    procedure Mount(Timeout: Integer); dispid 167;
    procedure Dismount(Timeout: Integer); dispid 168;
  end;

// *********************************************************************//
// Interface: IStorageGroup
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F46-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IStorageGroup = interface(IDispatch)
    ['{25150F46-5734-11D2-A593-00C04F990D8A}']
    function Get_DataSource: IDataSource2; safecall;
    function Get_Fields: Fields; safecall;
    function Get_Name: WideString; safecall;
    procedure Set_Name(const pName: WideString); safecall;
    function Get_PublicStoreDBs: OleVariant; safecall;
    function Get_MailboxStoreDBs: OleVariant; safecall;
    function Get_LogFilePath: WideString; safecall;
    function Get_SystemFilePath: WideString; safecall;
    function Get_CircularLogging: WordBool; safecall;
    procedure Set_CircularLogging(pCircularLogging: WordBool); safecall;
    function Get_ZeroDatabase: WordBool; safecall;
    procedure Set_ZeroDatabase(pZeroDatabase: WordBool); safecall;
    function GetInterface(const Interface_: WideString): IDispatch; safecall;
    procedure MoveLogFiles(const LogFilePath: WideString; Flags: Integer); safecall;
    procedure MoveSystemFiles(const SystemFilePath: WideString; Flags: Integer); safecall;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property Name: WideString read Get_Name write Set_Name;
    property PublicStoreDBs: OleVariant read Get_PublicStoreDBs;
    property MailboxStoreDBs: OleVariant read Get_MailboxStoreDBs;
    property LogFilePath: WideString read Get_LogFilePath;
    property SystemFilePath: WideString read Get_SystemFilePath;
    property CircularLogging: WordBool read Get_CircularLogging write Set_CircularLogging;
    property ZeroDatabase: WordBool read Get_ZeroDatabase write Set_ZeroDatabase;
  end;

// *********************************************************************//
// DispIntf:  IStorageGroupDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {25150F46-5734-11D2-A593-00C04F990D8A}
// *********************************************************************//
  IStorageGroupDisp = dispinterface
    ['{25150F46-5734-11D2-A593-00C04F990D8A}']
    property DataSource: IDataSource2 readonly dispid 100;
    property Fields: Fields readonly dispid 101;
    property Name: WideString dispid 102;
    property PublicStoreDBs: OleVariant readonly dispid 103;
    property MailboxStoreDBs: OleVariant readonly dispid 104;
    property LogFilePath: WideString readonly dispid 105;
    property SystemFilePath: WideString readonly dispid 106;
    property CircularLogging: WordBool dispid 107;
    property ZeroDatabase: WordBool dispid 108;
    function GetInterface(const Interface_: WideString): IDispatch; dispid 109;
    procedure MoveLogFiles(const LogFilePath: WideString; Flags: Integer); dispid 110;
    procedure MoveSystemFiles(const SystemFilePath: WideString; Flags: Integer); dispid 111;
  end;

// *********************************************************************//
// The Class CoMailGroup provides a Create and CreateRemote method to          
// create instances of the default interface IMailRecipient exposed by              
// the CoClass MailGroup. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMailGroup = class
    class function Create: IMailRecipient;
    class function CreateRemote(const MachineName: string): IMailRecipient;
  end;

// *********************************************************************//
// The Class CoMailRecipient provides a Create and CreateRemote method to          
// create instances of the default interface IMailRecipient exposed by              
// the CoClass MailRecipient. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMailRecipient = class
    class function Create: IMailRecipient;
    class function CreateRemote(const MachineName: string): IMailRecipient;
  end;

// *********************************************************************//
// The Class CoMailbox provides a Create and CreateRemote method to          
// create instances of the default interface IMailRecipient exposed by              
// the CoClass Mailbox. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMailbox = class
    class function Create: IMailRecipient;
    class function CreateRemote(const MachineName: string): IMailRecipient;
  end;

// *********************************************************************//
// The Class CoFolderAdmin provides a Create and CreateRemote method to          
// create instances of the default interface IMailRecipient exposed by              
// the CoClass FolderAdmin. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFolderAdmin = class
    class function Create: IMailRecipient;
    class function CreateRemote(const MachineName: string): IMailRecipient;
  end;

// *********************************************************************//
// The Class CoExchangeServer provides a Create and CreateRemote method to          
// create instances of the default interface IExchangeServer exposed by              
// the CoClass ExchangeServer. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoExchangeServer = class
    class function Create: IExchangeServer;
    class function CreateRemote(const MachineName: string): IExchangeServer;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TExchangeServer
// Help String      : Microsoft Exchange Server Object
// Default Interface: IExchangeServer
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TExchangeServerProperties= class;
{$ENDIF}
  TExchangeServer = class(TOleServer)
  private
    FIntf:        IExchangeServer;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TExchangeServerProperties;
    function      GetServerProperties: TExchangeServerProperties;
{$ENDIF}
    function      GetDefaultInterface: IExchangeServer;
  protected
    procedure InitServerData; override;
    function Get_DataSource: IDataSource2;
    function Get_Fields: Fields;
    function Get_Name: WideString;
    function Get_ExchangeVersion: WideString;
    function Get_StorageGroups: OleVariant;
    function Get_SubjectLoggingEnabled: WordBool;
    procedure Set_SubjectLoggingEnabled(pSubjectLoggingEnabled: WordBool);
    function Get_MessageTrackingEnabled: WordBool;
    procedure Set_MessageTrackingEnabled(pMessageTrackingEnabled: WordBool);
    function Get_DaysBeforeLogFileRemoval: Integer;
    procedure Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval: Integer);
    function Get_ServerType: CDOEXMServerType;
    procedure Set_ServerType(pServerType: CDOEXMServerType);
    function Get_DirectoryServer: WideString;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IExchangeServer);
    procedure Disconnect; override;
    function GetInterface(const Interface_: WideString): IDispatch;
    property DefaultInterface: IExchangeServer read GetDefaultInterface;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property Name: WideString read Get_Name;
    property ExchangeVersion: WideString read Get_ExchangeVersion;
    property StorageGroups: OleVariant read Get_StorageGroups;
    property DirectoryServer: WideString read Get_DirectoryServer;
    property SubjectLoggingEnabled: WordBool read Get_SubjectLoggingEnabled write Set_SubjectLoggingEnabled;
    property MessageTrackingEnabled: WordBool read Get_MessageTrackingEnabled write Set_MessageTrackingEnabled;
    property DaysBeforeLogFileRemoval: Integer read Get_DaysBeforeLogFileRemoval write Set_DaysBeforeLogFileRemoval;
    property ServerType: CDOEXMServerType read Get_ServerType write Set_ServerType;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TExchangeServerProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TExchangeServer
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TExchangeServerProperties = class(TPersistent)
  private
    FServer:    TExchangeServer;
    function    GetDefaultInterface: IExchangeServer;
    constructor Create(AServer: TExchangeServer);
  protected
    function Get_DataSource: IDataSource2;
    function Get_Fields: Fields;
    function Get_Name: WideString;
    function Get_ExchangeVersion: WideString;
    function Get_StorageGroups: OleVariant;
    function Get_SubjectLoggingEnabled: WordBool;
    procedure Set_SubjectLoggingEnabled(pSubjectLoggingEnabled: WordBool);
    function Get_MessageTrackingEnabled: WordBool;
    procedure Set_MessageTrackingEnabled(pMessageTrackingEnabled: WordBool);
    function Get_DaysBeforeLogFileRemoval: Integer;
    procedure Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval: Integer);
    function Get_ServerType: CDOEXMServerType;
    procedure Set_ServerType(pServerType: CDOEXMServerType);
    function Get_DirectoryServer: WideString;
  public
    property DefaultInterface: IExchangeServer read GetDefaultInterface;
  published
    property SubjectLoggingEnabled: WordBool read Get_SubjectLoggingEnabled write Set_SubjectLoggingEnabled;
    property MessageTrackingEnabled: WordBool read Get_MessageTrackingEnabled write Set_MessageTrackingEnabled;
    property DaysBeforeLogFileRemoval: Integer read Get_DaysBeforeLogFileRemoval write Set_DaysBeforeLogFileRemoval;
    property ServerType: CDOEXMServerType read Get_ServerType write Set_ServerType;
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoFolderTree provides a Create and CreateRemote method to          
// create instances of the default interface IFolderTree exposed by              
// the CoClass FolderTree. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoFolderTree = class
    class function Create: IFolderTree;
    class function CreateRemote(const MachineName: string): IFolderTree;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TFolderTree
// Help String      : Microsoft Exchange Public Folder Tree Object
// Default Interface: IFolderTree
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TFolderTreeProperties= class;
{$ENDIF}
  TFolderTree = class(TOleServer)
  private
    FIntf:        IFolderTree;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TFolderTreeProperties;
    function      GetServerProperties: TFolderTreeProperties;
{$ENDIF}
    function      GetDefaultInterface: IFolderTree;
  protected
    procedure InitServerData; override;
    function Get_DataSource: IDataSource2;
    function Get_Fields: Fields;
    function Get_Name: WideString;
    procedure Set_Name(const pName: WideString);
    function Get_StoreDBs: OleVariant;
    function Get_TreeType: CDOEXMFolderTreeType;
    function Get_RootFolderURL: WideString;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IFolderTree);
    procedure Disconnect; override;
    function GetInterface(const Interface_: WideString): IDispatch;
    property DefaultInterface: IFolderTree read GetDefaultInterface;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property StoreDBs: OleVariant read Get_StoreDBs;
    property TreeType: CDOEXMFolderTreeType read Get_TreeType;
    property RootFolderURL: WideString read Get_RootFolderURL;
    property Name: WideString read Get_Name write Set_Name;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TFolderTreeProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TFolderTree
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TFolderTreeProperties = class(TPersistent)
  private
    FServer:    TFolderTree;
    function    GetDefaultInterface: IFolderTree;
    constructor Create(AServer: TFolderTree);
  protected
    function Get_DataSource: IDataSource2;
    function Get_Fields: Fields;
    function Get_Name: WideString;
    procedure Set_Name(const pName: WideString);
    function Get_StoreDBs: OleVariant;
    function Get_TreeType: CDOEXMFolderTreeType;
    function Get_RootFolderURL: WideString;
  public
    property DefaultInterface: IFolderTree read GetDefaultInterface;
  published
    property Name: WideString read Get_Name write Set_Name;
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoPublicStoreDB provides a Create and CreateRemote method to          
// create instances of the default interface IPublicStoreDB2 exposed by              
// the CoClass PublicStoreDB. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoPublicStoreDB = class
    class function Create: IPublicStoreDB2;
    class function CreateRemote(const MachineName: string): IPublicStoreDB2;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TPublicStoreDB
// Help String      : Microsoft Exchange Public Store Database Object
// Default Interface: IPublicStoreDB2
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TPublicStoreDBProperties= class;
{$ENDIF}
  TPublicStoreDB = class(TOleServer)
  private
    FIntf:        IPublicStoreDB2;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TPublicStoreDBProperties;
    function      GetServerProperties: TPublicStoreDBProperties;
{$ENDIF}
    function      GetDefaultInterface: IPublicStoreDB2;
  protected
    procedure InitServerData; override;
    function Get_LastFullBackupTime: OleVariant;
    function Get_LastIncrementalBackupTime: OleVariant;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IPublicStoreDB2);
    procedure Disconnect; override;
    property DefaultInterface: IPublicStoreDB2 read GetDefaultInterface;
    property LastFullBackupTime: OleVariant read Get_LastFullBackupTime;
    property LastIncrementalBackupTime: OleVariant read Get_LastIncrementalBackupTime;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TPublicStoreDBProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TPublicStoreDB
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TPublicStoreDBProperties = class(TPersistent)
  private
    FServer:    TPublicStoreDB;
    function    GetDefaultInterface: IPublicStoreDB2;
    constructor Create(AServer: TPublicStoreDB);
  protected
    function Get_LastFullBackupTime: OleVariant;
    function Get_LastIncrementalBackupTime: OleVariant;
  public
    property DefaultInterface: IPublicStoreDB2 read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoMailboxStoreDB provides a Create and CreateRemote method to          
// create instances of the default interface IMailboxStoreDB2 exposed by              
// the CoClass MailboxStoreDB. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoMailboxStoreDB = class
    class function Create: IMailboxStoreDB2;
    class function CreateRemote(const MachineName: string): IMailboxStoreDB2;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TMailboxStoreDB
// Help String      : Microsoft Exchange Mailbox Store Database Object
// Default Interface: IMailboxStoreDB2
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TMailboxStoreDBProperties= class;
{$ENDIF}
  TMailboxStoreDB = class(TOleServer)
  private
    FIntf:        IMailboxStoreDB2;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TMailboxStoreDBProperties;
    function      GetServerProperties: TMailboxStoreDBProperties;
{$ENDIF}
    function      GetDefaultInterface: IMailboxStoreDB2;
  protected
    procedure InitServerData; override;
    function Get_LastFullBackupTime: OleVariant;
    function Get_LastIncrementalBackupTime: OleVariant;
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IMailboxStoreDB2);
    procedure Disconnect; override;
    property DefaultInterface: IMailboxStoreDB2 read GetDefaultInterface;
    property LastFullBackupTime: OleVariant read Get_LastFullBackupTime;
    property LastIncrementalBackupTime: OleVariant read Get_LastIncrementalBackupTime;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TMailboxStoreDBProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TMailboxStoreDB
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TMailboxStoreDBProperties = class(TPersistent)
  private
    FServer:    TMailboxStoreDB;
    function    GetDefaultInterface: IMailboxStoreDB2;
    constructor Create(AServer: TMailboxStoreDB);
  protected
    function Get_LastFullBackupTime: OleVariant;
    function Get_LastIncrementalBackupTime: OleVariant;
  public
    property DefaultInterface: IMailboxStoreDB2 read GetDefaultInterface;
  published
  end;
{$ENDIF}


// *********************************************************************//
// The Class CoStorageGroup provides a Create and CreateRemote method to          
// create instances of the default interface IStorageGroup exposed by              
// the CoClass StorageGroup. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoStorageGroup = class
    class function Create: IStorageGroup;
    class function CreateRemote(const MachineName: string): IStorageGroup;
  end;


// *********************************************************************//
// OLE Server Proxy class declaration
// Server Object    : TStorageGroup
// Help String      : Microsoft Exchange Storage Group Object
// Default Interface: IStorageGroup
// Def. Intf. DISP? : No
// Event   Interface: 
// TypeFlags        : (2) CanCreate
// *********************************************************************//
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  TStorageGroupProperties= class;
{$ENDIF}
  TStorageGroup = class(TOleServer)
  private
    FIntf:        IStorageGroup;
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    FProps:       TStorageGroupProperties;
    function      GetServerProperties: TStorageGroupProperties;
{$ENDIF}
    function      GetDefaultInterface: IStorageGroup;
  protected
    procedure InitServerData; override;
    function Get_DataSource: IDataSource2;
    function Get_Fields: Fields;
    function Get_Name: WideString;
    procedure Set_Name(const pName: WideString);
    function Get_PublicStoreDBs: OleVariant;
    function Get_MailboxStoreDBs: OleVariant;
    function Get_LogFilePath: WideString;
    function Get_SystemFilePath: WideString;
    function Get_CircularLogging: WordBool;
    procedure Set_CircularLogging(pCircularLogging: WordBool);
    function Get_ZeroDatabase: WordBool;
    procedure Set_ZeroDatabase(pZeroDatabase: WordBool);
  public
    constructor Create(AOwner: TComponent); override;
    destructor  Destroy; override;
    procedure Connect; override;
    procedure ConnectTo(svrIntf: IStorageGroup);
    procedure Disconnect; override;
    function GetInterface(const Interface_: WideString): IDispatch;
    procedure MoveLogFiles(const LogFilePath: WideString; Flags: Integer);
    procedure MoveSystemFiles(const SystemFilePath: WideString; Flags: Integer);
    property DefaultInterface: IStorageGroup read GetDefaultInterface;
    property DataSource: IDataSource2 read Get_DataSource;
    property Fields: Fields read Get_Fields;
    property PublicStoreDBs: OleVariant read Get_PublicStoreDBs;
    property MailboxStoreDBs: OleVariant read Get_MailboxStoreDBs;
    property LogFilePath: WideString read Get_LogFilePath;
    property SystemFilePath: WideString read Get_SystemFilePath;
    property Name: WideString read Get_Name write Set_Name;
    property CircularLogging: WordBool read Get_CircularLogging write Set_CircularLogging;
    property ZeroDatabase: WordBool read Get_ZeroDatabase write Set_ZeroDatabase;
  published
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
    property Server: TStorageGroupProperties read GetServerProperties;
{$ENDIF}
  end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
// *********************************************************************//
// OLE Server Properties Proxy Class
// Server Object    : TStorageGroup
// (This object is used by the IDE's Property Inspector to allow editing
//  of the properties of this server)
// *********************************************************************//
 TStorageGroupProperties = class(TPersistent)
  private
    FServer:    TStorageGroup;
    function    GetDefaultInterface: IStorageGroup;
    constructor Create(AServer: TStorageGroup);
  protected
    function Get_DataSource: IDataSource2;
    function Get_Fields: Fields;
    function Get_Name: WideString;
    procedure Set_Name(const pName: WideString);
    function Get_PublicStoreDBs: OleVariant;
    function Get_MailboxStoreDBs: OleVariant;
    function Get_LogFilePath: WideString;
    function Get_SystemFilePath: WideString;
    function Get_CircularLogging: WordBool;
    procedure Set_CircularLogging(pCircularLogging: WordBool);
    function Get_ZeroDatabase: WordBool;
    procedure Set_ZeroDatabase(pZeroDatabase: WordBool);
  public
    property DefaultInterface: IStorageGroup read GetDefaultInterface;
  published
    property Name: WideString read Get_Name write Set_Name;
    property CircularLogging: WordBool read Get_CircularLogging write Set_CircularLogging;
    property ZeroDatabase: WordBool read Get_ZeroDatabase write Set_ZeroDatabase;
  end;
{$ENDIF}


procedure Register;

resourcestring
  dtlServerPage = '(none)';

  dtlOcxPage = '(none)';

implementation

uses ComObj;

class function CoMailGroup.Create: IMailRecipient;
begin
  Result := CreateComObject(CLASS_MailGroup) as IMailRecipient;
end;

class function CoMailGroup.CreateRemote(const MachineName: string): IMailRecipient;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_MailGroup) as IMailRecipient;
end;

class function CoMailRecipient.Create: IMailRecipient;
begin
  Result := CreateComObject(CLASS_MailRecipient) as IMailRecipient;
end;

class function CoMailRecipient.CreateRemote(const MachineName: string): IMailRecipient;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_MailRecipient) as IMailRecipient;
end;

class function CoMailbox.Create: IMailRecipient;
begin
  Result := CreateComObject(CLASS_Mailbox) as IMailRecipient;
end;

class function CoMailbox.CreateRemote(const MachineName: string): IMailRecipient;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_Mailbox) as IMailRecipient;
end;

class function CoFolderAdmin.Create: IMailRecipient;
begin
  Result := CreateComObject(CLASS_FolderAdmin) as IMailRecipient;
end;

class function CoFolderAdmin.CreateRemote(const MachineName: string): IMailRecipient;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FolderAdmin) as IMailRecipient;
end;

class function CoExchangeServer.Create: IExchangeServer;
begin
  Result := CreateComObject(CLASS_ExchangeServer) as IExchangeServer;
end;

class function CoExchangeServer.CreateRemote(const MachineName: string): IExchangeServer;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ExchangeServer) as IExchangeServer;
end;

procedure TExchangeServer.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{25150F27-5734-11D2-A593-00C04F990D8A}';
    IntfIID:   '{25150F47-5734-11D2-A593-00C04F990D8A}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TExchangeServer.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IExchangeServer;
  end;
end;

procedure TExchangeServer.ConnectTo(svrIntf: IExchangeServer);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TExchangeServer.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TExchangeServer.GetDefaultInterface: IExchangeServer;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TExchangeServer.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TExchangeServerProperties.Create(Self);
{$ENDIF}
end;

destructor TExchangeServer.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TExchangeServer.GetServerProperties: TExchangeServerProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TExchangeServer.Get_DataSource: IDataSource2;
begin
    Result := DefaultInterface.DataSource;
end;

function TExchangeServer.Get_Fields: Fields;
begin
    Result := DefaultInterface.Fields;
end;

function TExchangeServer.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

function TExchangeServer.Get_ExchangeVersion: WideString;
begin
    Result := DefaultInterface.ExchangeVersion;
end;

function TExchangeServer.Get_StorageGroups: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.StorageGroups;
end;

function TExchangeServer.Get_SubjectLoggingEnabled: WordBool;
begin
    Result := DefaultInterface.SubjectLoggingEnabled;
end;

procedure TExchangeServer.Set_SubjectLoggingEnabled(pSubjectLoggingEnabled: WordBool);
begin
  DefaultInterface.Set_SubjectLoggingEnabled(pSubjectLoggingEnabled);
end;

function TExchangeServer.Get_MessageTrackingEnabled: WordBool;
begin
    Result := DefaultInterface.MessageTrackingEnabled;
end;

procedure TExchangeServer.Set_MessageTrackingEnabled(pMessageTrackingEnabled: WordBool);
begin
  DefaultInterface.Set_MessageTrackingEnabled(pMessageTrackingEnabled);
end;

function TExchangeServer.Get_DaysBeforeLogFileRemoval: Integer;
begin
    Result := DefaultInterface.DaysBeforeLogFileRemoval;
end;

procedure TExchangeServer.Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval: Integer);
begin
  DefaultInterface.Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval);
end;

function TExchangeServer.Get_ServerType: CDOEXMServerType;
begin
    Result := DefaultInterface.ServerType;
end;

procedure TExchangeServer.Set_ServerType(pServerType: CDOEXMServerType);
begin
  DefaultInterface.Set_ServerType(pServerType);
end;

function TExchangeServer.Get_DirectoryServer: WideString;
begin
    Result := DefaultInterface.DirectoryServer;
end;

function TExchangeServer.GetInterface(const Interface_: WideString): IDispatch;
begin
  Result := DefaultInterface.GetInterface(Interface_);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TExchangeServerProperties.Create(AServer: TExchangeServer);
begin
  inherited Create;
  FServer := AServer;
end;

function TExchangeServerProperties.GetDefaultInterface: IExchangeServer;
begin
  Result := FServer.DefaultInterface;
end;

function TExchangeServerProperties.Get_DataSource: IDataSource2;
begin
    Result := DefaultInterface.DataSource;
end;

function TExchangeServerProperties.Get_Fields: Fields;
begin
    Result := DefaultInterface.Fields;
end;

function TExchangeServerProperties.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

function TExchangeServerProperties.Get_ExchangeVersion: WideString;
begin
    Result := DefaultInterface.ExchangeVersion;
end;

function TExchangeServerProperties.Get_StorageGroups: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.StorageGroups;
end;

function TExchangeServerProperties.Get_SubjectLoggingEnabled: WordBool;
begin
    Result := DefaultInterface.SubjectLoggingEnabled;
end;

procedure TExchangeServerProperties.Set_SubjectLoggingEnabled(pSubjectLoggingEnabled: WordBool);
begin
  DefaultInterface.Set_SubjectLoggingEnabled(pSubjectLoggingEnabled);
end;

function TExchangeServerProperties.Get_MessageTrackingEnabled: WordBool;
begin
    Result := DefaultInterface.MessageTrackingEnabled;
end;

procedure TExchangeServerProperties.Set_MessageTrackingEnabled(pMessageTrackingEnabled: WordBool);
begin
  DefaultInterface.Set_MessageTrackingEnabled(pMessageTrackingEnabled);
end;

function TExchangeServerProperties.Get_DaysBeforeLogFileRemoval: Integer;
begin
    Result := DefaultInterface.DaysBeforeLogFileRemoval;
end;

procedure TExchangeServerProperties.Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval: Integer);
begin
  DefaultInterface.Set_DaysBeforeLogFileRemoval(pDaysBeforeLogFileRemoval);
end;

function TExchangeServerProperties.Get_ServerType: CDOEXMServerType;
begin
    Result := DefaultInterface.ServerType;
end;

procedure TExchangeServerProperties.Set_ServerType(pServerType: CDOEXMServerType);
begin
  DefaultInterface.Set_ServerType(pServerType);
end;

function TExchangeServerProperties.Get_DirectoryServer: WideString;
begin
    Result := DefaultInterface.DirectoryServer;
end;

{$ENDIF}

class function CoFolderTree.Create: IFolderTree;
begin
  Result := CreateComObject(CLASS_FolderTree) as IFolderTree;
end;

class function CoFolderTree.CreateRemote(const MachineName: string): IFolderTree;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_FolderTree) as IFolderTree;
end;

procedure TFolderTree.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{25150F23-5734-11D2-A593-00C04F990D8A}';
    IntfIID:   '{25150F43-5734-11D2-A593-00C04F990D8A}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TFolderTree.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IFolderTree;
  end;
end;

procedure TFolderTree.ConnectTo(svrIntf: IFolderTree);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TFolderTree.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TFolderTree.GetDefaultInterface: IFolderTree;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TFolderTree.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TFolderTreeProperties.Create(Self);
{$ENDIF}
end;

destructor TFolderTree.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TFolderTree.GetServerProperties: TFolderTreeProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TFolderTree.Get_DataSource: IDataSource2;
begin
    Result := DefaultInterface.DataSource;
end;

function TFolderTree.Get_Fields: Fields;
begin
    Result := DefaultInterface.Fields;
end;

function TFolderTree.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

procedure TFolderTree.Set_Name(const pName: WideString);
  { Warning: The property Name has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Name := pName;
end;

function TFolderTree.Get_StoreDBs: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.StoreDBs;
end;

function TFolderTree.Get_TreeType: CDOEXMFolderTreeType;
begin
    Result := DefaultInterface.TreeType;
end;

function TFolderTree.Get_RootFolderURL: WideString;
begin
    Result := DefaultInterface.RootFolderURL;
end;

function TFolderTree.GetInterface(const Interface_: WideString): IDispatch;
begin
  Result := DefaultInterface.GetInterface(Interface_);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TFolderTreeProperties.Create(AServer: TFolderTree);
begin
  inherited Create;
  FServer := AServer;
end;

function TFolderTreeProperties.GetDefaultInterface: IFolderTree;
begin
  Result := FServer.DefaultInterface;
end;

function TFolderTreeProperties.Get_DataSource: IDataSource2;
begin
    Result := DefaultInterface.DataSource;
end;

function TFolderTreeProperties.Get_Fields: Fields;
begin
    Result := DefaultInterface.Fields;
end;

function TFolderTreeProperties.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

procedure TFolderTreeProperties.Set_Name(const pName: WideString);
  { Warning: The property Name has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Name := pName;
end;

function TFolderTreeProperties.Get_StoreDBs: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.StoreDBs;
end;

function TFolderTreeProperties.Get_TreeType: CDOEXMFolderTreeType;
begin
    Result := DefaultInterface.TreeType;
end;

function TFolderTreeProperties.Get_RootFolderURL: WideString;
begin
    Result := DefaultInterface.RootFolderURL;
end;

{$ENDIF}

class function CoPublicStoreDB.Create: IPublicStoreDB2;
begin
  Result := CreateComObject(CLASS_PublicStoreDB) as IPublicStoreDB2;
end;

class function CoPublicStoreDB.CreateRemote(const MachineName: string): IPublicStoreDB2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_PublicStoreDB) as IPublicStoreDB2;
end;

procedure TPublicStoreDB.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{25150F24-5734-11D2-A593-00C04F990D8A}';
    IntfIID:   '{25150F49-5734-11D2-A593-00C04F990D8A}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TPublicStoreDB.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IPublicStoreDB2;
  end;
end;

procedure TPublicStoreDB.ConnectTo(svrIntf: IPublicStoreDB2);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TPublicStoreDB.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TPublicStoreDB.GetDefaultInterface: IPublicStoreDB2;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TPublicStoreDB.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TPublicStoreDBProperties.Create(Self);
{$ENDIF}
end;

destructor TPublicStoreDB.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TPublicStoreDB.GetServerProperties: TPublicStoreDBProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TPublicStoreDB.Get_LastFullBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastFullBackupTime;
end;

function TPublicStoreDB.Get_LastIncrementalBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastIncrementalBackupTime;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TPublicStoreDBProperties.Create(AServer: TPublicStoreDB);
begin
  inherited Create;
  FServer := AServer;
end;

function TPublicStoreDBProperties.GetDefaultInterface: IPublicStoreDB2;
begin
  Result := FServer.DefaultInterface;
end;

function TPublicStoreDBProperties.Get_LastFullBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastFullBackupTime;
end;

function TPublicStoreDBProperties.Get_LastIncrementalBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastIncrementalBackupTime;
end;

{$ENDIF}

class function CoMailboxStoreDB.Create: IMailboxStoreDB2;
begin
  Result := CreateComObject(CLASS_MailboxStoreDB) as IMailboxStoreDB2;
end;

class function CoMailboxStoreDB.CreateRemote(const MachineName: string): IMailboxStoreDB2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_MailboxStoreDB) as IMailboxStoreDB2;
end;

procedure TMailboxStoreDB.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{25150F25-5734-11D2-A593-00C04F990D8A}';
    IntfIID:   '{25150F4A-5734-11D2-A593-00C04F990D8A}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TMailboxStoreDB.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IMailboxStoreDB2;
  end;
end;

procedure TMailboxStoreDB.ConnectTo(svrIntf: IMailboxStoreDB2);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TMailboxStoreDB.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TMailboxStoreDB.GetDefaultInterface: IMailboxStoreDB2;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TMailboxStoreDB.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TMailboxStoreDBProperties.Create(Self);
{$ENDIF}
end;

destructor TMailboxStoreDB.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TMailboxStoreDB.GetServerProperties: TMailboxStoreDBProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TMailboxStoreDB.Get_LastFullBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastFullBackupTime;
end;

function TMailboxStoreDB.Get_LastIncrementalBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastIncrementalBackupTime;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TMailboxStoreDBProperties.Create(AServer: TMailboxStoreDB);
begin
  inherited Create;
  FServer := AServer;
end;

function TMailboxStoreDBProperties.GetDefaultInterface: IMailboxStoreDB2;
begin
  Result := FServer.DefaultInterface;
end;

function TMailboxStoreDBProperties.Get_LastFullBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastFullBackupTime;
end;

function TMailboxStoreDBProperties.Get_LastIncrementalBackupTime: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.LastIncrementalBackupTime;
end;

{$ENDIF}

class function CoStorageGroup.Create: IStorageGroup;
begin
  Result := CreateComObject(CLASS_StorageGroup) as IStorageGroup;
end;

class function CoStorageGroup.CreateRemote(const MachineName: string): IStorageGroup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_StorageGroup) as IStorageGroup;
end;

procedure TStorageGroup.InitServerData;
const
  CServerData: TServerData = (
    ClassID:   '{25150F26-5734-11D2-A593-00C04F990D8A}';
    IntfIID:   '{25150F46-5734-11D2-A593-00C04F990D8A}';
    EventIID:  '';
    LicenseKey: nil;
    Version: 500);
begin
  ServerData := @CServerData;
end;

procedure TStorageGroup.Connect;
var
  punk: IUnknown;
begin
  if FIntf = nil then
  begin
    punk := GetServer;
    Fintf:= punk as IStorageGroup;
  end;
end;

procedure TStorageGroup.ConnectTo(svrIntf: IStorageGroup);
begin
  Disconnect;
  FIntf := svrIntf;
end;

procedure TStorageGroup.DisConnect;
begin
  if Fintf <> nil then
  begin
    FIntf := nil;
  end;
end;

function TStorageGroup.GetDefaultInterface: IStorageGroup;
begin
  if FIntf = nil then
    Connect;
  Assert(FIntf <> nil, 'DefaultInterface is NULL. Component is not connected to Server. You must call ''Connect'' or ''ConnectTo'' before this operation');
  Result := FIntf;
end;

constructor TStorageGroup.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps := TStorageGroupProperties.Create(Self);
{$ENDIF}
end;

destructor TStorageGroup.Destroy;
begin
{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
  FProps.Free;
{$ENDIF}
  inherited Destroy;
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
function TStorageGroup.GetServerProperties: TStorageGroupProperties;
begin
  Result := FProps;
end;
{$ENDIF}

function TStorageGroup.Get_DataSource: IDataSource2;
begin
    Result := DefaultInterface.DataSource;
end;

function TStorageGroup.Get_Fields: Fields;
begin
    Result := DefaultInterface.Fields;
end;

function TStorageGroup.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

procedure TStorageGroup.Set_Name(const pName: WideString);
  { Warning: The property Name has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Name := pName;
end;

function TStorageGroup.Get_PublicStoreDBs: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.PublicStoreDBs;
end;

function TStorageGroup.Get_MailboxStoreDBs: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.MailboxStoreDBs;
end;

function TStorageGroup.Get_LogFilePath: WideString;
begin
    Result := DefaultInterface.LogFilePath;
end;

function TStorageGroup.Get_SystemFilePath: WideString;
begin
    Result := DefaultInterface.SystemFilePath;
end;

function TStorageGroup.Get_CircularLogging: WordBool;
begin
    Result := DefaultInterface.CircularLogging;
end;

procedure TStorageGroup.Set_CircularLogging(pCircularLogging: WordBool);
begin
  DefaultInterface.Set_CircularLogging(pCircularLogging);
end;

function TStorageGroup.Get_ZeroDatabase: WordBool;
begin
    Result := DefaultInterface.ZeroDatabase;
end;

procedure TStorageGroup.Set_ZeroDatabase(pZeroDatabase: WordBool);
begin
  DefaultInterface.Set_ZeroDatabase(pZeroDatabase);
end;

function TStorageGroup.GetInterface(const Interface_: WideString): IDispatch;
begin
  Result := DefaultInterface.GetInterface(Interface_);
end;

procedure TStorageGroup.MoveLogFiles(const LogFilePath: WideString; Flags: Integer);
begin
  DefaultInterface.MoveLogFiles(LogFilePath, Flags);
end;

procedure TStorageGroup.MoveSystemFiles(const SystemFilePath: WideString; Flags: Integer);
begin
  DefaultInterface.MoveSystemFiles(SystemFilePath, Flags);
end;

{$IFDEF LIVE_SERVER_AT_DESIGN_TIME}
constructor TStorageGroupProperties.Create(AServer: TStorageGroup);
begin
  inherited Create;
  FServer := AServer;
end;

function TStorageGroupProperties.GetDefaultInterface: IStorageGroup;
begin
  Result := FServer.DefaultInterface;
end;

function TStorageGroupProperties.Get_DataSource: IDataSource2;
begin
    Result := DefaultInterface.DataSource;
end;

function TStorageGroupProperties.Get_Fields: Fields;
begin
    Result := DefaultInterface.Fields;
end;

function TStorageGroupProperties.Get_Name: WideString;
begin
    Result := DefaultInterface.Name;
end;

procedure TStorageGroupProperties.Set_Name(const pName: WideString);
  { Warning: The property Name has a setter and a getter whose
    types do not match. Delphi was unable to generate a property of
    this sort and so is using a Variant as a passthrough. }
var
  InterfaceVariant: OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  InterfaceVariant.Name := pName;
end;

function TStorageGroupProperties.Get_PublicStoreDBs: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.PublicStoreDBs;
end;

function TStorageGroupProperties.Get_MailboxStoreDBs: OleVariant;
var
  InterfaceVariant : OleVariant;
begin
  InterfaceVariant := DefaultInterface;
  Result := InterfaceVariant.MailboxStoreDBs;
end;

function TStorageGroupProperties.Get_LogFilePath: WideString;
begin
    Result := DefaultInterface.LogFilePath;
end;

function TStorageGroupProperties.Get_SystemFilePath: WideString;
begin
    Result := DefaultInterface.SystemFilePath;
end;

function TStorageGroupProperties.Get_CircularLogging: WordBool;
begin
    Result := DefaultInterface.CircularLogging;
end;

procedure TStorageGroupProperties.Set_CircularLogging(pCircularLogging: WordBool);
begin
  DefaultInterface.Set_CircularLogging(pCircularLogging);
end;

function TStorageGroupProperties.Get_ZeroDatabase: WordBool;
begin
    Result := DefaultInterface.ZeroDatabase;
end;

procedure TStorageGroupProperties.Set_ZeroDatabase(pZeroDatabase: WordBool);
begin
  DefaultInterface.Set_ZeroDatabase(pZeroDatabase);
end;

{$ENDIF}

procedure Register;
begin
  RegisterComponents(dtlServerPage, [TExchangeServer, TFolderTree, TPublicStoreDB, TMailboxStoreDB, 
    TStorageGroup]);
end;

end.
