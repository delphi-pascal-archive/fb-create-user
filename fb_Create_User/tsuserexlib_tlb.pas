unit TSUSEREXLib_TLB;

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
// File generated on 23/10/2008 11:41:12 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\WINDOWS\system32\tsuserex.dll (1)
// LIBID: {45413F04-DF86-11D1-AE27-00C04FA35813}
// LCID: 0
// Helpfile: 
// HelpString: tsexusrm 1.0 Type Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\Stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  TSUSEREXLibMajorVersion = 1;
  TSUSEREXLibMinorVersion = 0;

  LIBID_TSUSEREXLib: TGUID = '{45413F04-DF86-11D1-AE27-00C04FA35813}';

  CLASS_TSUserExInterfaces: TGUID = '{0910DD01-DF8C-11D1-AE27-00C04FA35813}';
  IID_IADsTSUserEx: TGUID = '{C4930E79-2989-4462-8A60-2FCF2F2955EF}';
  CLASS_ADsTSUserEx: TGUID = '{E2E9CAE6-1E7B-4B8E-BABD-E9BF6292AC29}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IADsTSUserEx = interface;
  IADsTSUserExDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  TSUserExInterfaces = IUnknown;
  ADsTSUserEx = IADsTSUserEx;


// *********************************************************************//
// Interface: IADsTSUserEx
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C4930E79-2989-4462-8A60-2FCF2F2955EF}
// *********************************************************************//
  IADsTSUserEx = interface(IDispatch)
    ['{C4930E79-2989-4462-8A60-2FCF2F2955EF}']
    function Get_TerminalServicesProfilePath: WideString; safecall;
    procedure Set_TerminalServicesProfilePath(const pVal: WideString); safecall;
    function Get_TerminalServicesHomeDirectory: WideString; safecall;
    procedure Set_TerminalServicesHomeDirectory(const pVal: WideString); safecall;
    function Get_TerminalServicesHomeDrive: WideString; safecall;
    procedure Set_TerminalServicesHomeDrive(const pVal: WideString); safecall;
    function Get_AllowLogon: Integer; safecall;
    procedure Set_AllowLogon(pVal: Integer); safecall;
    function Get_EnableRemoteControl: Integer; safecall;
    procedure Set_EnableRemoteControl(pVal: Integer); safecall;
    function Get_MaxDisconnectionTime: Integer; safecall;
    procedure Set_MaxDisconnectionTime(pVal: Integer); safecall;
    function Get_MaxConnectionTime: Integer; safecall;
    procedure Set_MaxConnectionTime(pVal: Integer); safecall;
    function Get_MaxIdleTime: Integer; safecall;
    procedure Set_MaxIdleTime(pVal: Integer); safecall;
    function Get_ReconnectionAction: Integer; safecall;
    procedure Set_ReconnectionAction(pNewVal: Integer); safecall;
    function Get_BrokenConnectionAction: Integer; safecall;
    procedure Set_BrokenConnectionAction(pNewVal: Integer); safecall;
    function Get_ConnectClientDrivesAtLogon: Integer; safecall;
    procedure Set_ConnectClientDrivesAtLogon(pNewVal: Integer); safecall;
    function Get_ConnectClientPrintersAtLogon: Integer; safecall;
    procedure Set_ConnectClientPrintersAtLogon(pVal: Integer); safecall;
    function Get_DefaultToMainPrinter: Integer; safecall;
    procedure Set_DefaultToMainPrinter(pVal: Integer); safecall;
    function Get_TerminalServicesWorkDirectory: WideString; safecall;
    procedure Set_TerminalServicesWorkDirectory(const pVal: WideString); safecall;
    function Get_TerminalServicesInitialProgram: WideString; safecall;
    procedure Set_TerminalServicesInitialProgram(const pVal: WideString); safecall;
    property TerminalServicesProfilePath: WideString read Get_TerminalServicesProfilePath write Set_TerminalServicesProfilePath;
    property TerminalServicesHomeDirectory: WideString read Get_TerminalServicesHomeDirectory write Set_TerminalServicesHomeDirectory;
    property TerminalServicesHomeDrive: WideString read Get_TerminalServicesHomeDrive write Set_TerminalServicesHomeDrive;
    property AllowLogon: Integer read Get_AllowLogon write Set_AllowLogon;
    property EnableRemoteControl: Integer read Get_EnableRemoteControl write Set_EnableRemoteControl;
    property MaxDisconnectionTime: Integer read Get_MaxDisconnectionTime write Set_MaxDisconnectionTime;
    property MaxConnectionTime: Integer read Get_MaxConnectionTime write Set_MaxConnectionTime;
    property MaxIdleTime: Integer read Get_MaxIdleTime write Set_MaxIdleTime;
    property ReconnectionAction: Integer read Get_ReconnectionAction write Set_ReconnectionAction;
    property BrokenConnectionAction: Integer read Get_BrokenConnectionAction write Set_BrokenConnectionAction;
    property ConnectClientDrivesAtLogon: Integer read Get_ConnectClientDrivesAtLogon write Set_ConnectClientDrivesAtLogon;
    property ConnectClientPrintersAtLogon: Integer read Get_ConnectClientPrintersAtLogon write Set_ConnectClientPrintersAtLogon;
    property DefaultToMainPrinter: Integer read Get_DefaultToMainPrinter write Set_DefaultToMainPrinter;
    property TerminalServicesWorkDirectory: WideString read Get_TerminalServicesWorkDirectory write Set_TerminalServicesWorkDirectory;
    property TerminalServicesInitialProgram: WideString read Get_TerminalServicesInitialProgram write Set_TerminalServicesInitialProgram;
  end;

// *********************************************************************//
// DispIntf:  IADsTSUserExDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {C4930E79-2989-4462-8A60-2FCF2F2955EF}
// *********************************************************************//
  IADsTSUserExDisp = dispinterface
    ['{C4930E79-2989-4462-8A60-2FCF2F2955EF}']
    property TerminalServicesProfilePath: WideString dispid 1;
    property TerminalServicesHomeDirectory: WideString dispid 2;
    property TerminalServicesHomeDrive: WideString dispid 3;
    property AllowLogon: Integer dispid 4;
    property EnableRemoteControl: Integer dispid 5;
    property MaxDisconnectionTime: Integer dispid 6;
    property MaxConnectionTime: Integer dispid 7;
    property MaxIdleTime: Integer dispid 8;
    property ReconnectionAction: Integer dispid 9;
    property BrokenConnectionAction: Integer dispid 10;
    property ConnectClientDrivesAtLogon: Integer dispid 11;
    property ConnectClientPrintersAtLogon: Integer dispid 12;
    property DefaultToMainPrinter: Integer dispid 13;
    property TerminalServicesWorkDirectory: WideString dispid 14;
    property TerminalServicesInitialProgram: WideString dispid 15;
  end;

// *********************************************************************//
// The Class CoTSUserExInterfaces provides a Create and CreateRemote method to          
// create instances of the default interface IUnknown exposed by              
// the CoClass TSUserExInterfaces. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoTSUserExInterfaces = class
    class function Create: IUnknown;
    class function CreateRemote(const MachineName: string): IUnknown;
  end;

// *********************************************************************//
// The Class CoADsTSUserEx provides a Create and CreateRemote method to          
// create instances of the default interface IADsTSUserEx exposed by              
// the CoClass ADsTSUserEx. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoADsTSUserEx = class
    class function Create: IADsTSUserEx;
    class function CreateRemote(const MachineName: string): IADsTSUserEx;
  end;

implementation

uses ComObj;

class function CoTSUserExInterfaces.Create: IUnknown;
begin
  Result := CreateComObject(CLASS_TSUserExInterfaces) as IUnknown;
end;

class function CoTSUserExInterfaces.CreateRemote(const MachineName: string): IUnknown;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_TSUserExInterfaces) as IUnknown;
end;

class function CoADsTSUserEx.Create: IADsTSUserEx;
begin
  Result := CreateComObject(CLASS_ADsTSUserEx) as IADsTSUserEx;
end;

class function CoADsTSUserEx.CreateRemote(const MachineName: string): IADsTSUserEx;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_ADsTSUserEx) as IADsTSUserEx;
end;

end.
