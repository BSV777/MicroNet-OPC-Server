unit MicroNet_TLB;

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

// PASTLWTR : $Revision:   1.88.1.0.1.0  $
// File generated on 30.09.02 15:37:22 from Type Library described below.

// ************************************************************************ //
// Type Lib: C:\Arbeit\Коксохим\MicroNet OPC-Server\MicroNet.tlb (1)
// IID\LCID: {A42F19F1-608B-11D3-B98D-00403357BAA5}\0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\System32\StdOle2.tlb)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\STDVCL40.DLL)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls, StdVCL;

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  MicroNetMajorVersion = 2;
  MicroNetMinorVersion = 0;

  LIBID_MicroNet: TGUID = '{A42F19F1-608B-11D3-B98D-00403357BAA5}';

  IID_IDA2: TGUID = '{A42F19F2-608B-11D3-B98D-00403357BAA5}';
  CLASS_OPCServer: TGUID = '{A42F19F4-608B-11D3-B98D-00403357BAA5}';
  IID_DA2Unknown: TGUID = '{DEBAB081-6AE4-11D3-B995-00403357BAA5}';
  IID_IOPCGroup: TGUID = '{B1779B97-71B6-11D3-B996-00104B33F2C4}';
  CLASS_OPCGroup: TGUID = '{B1779B99-71B6-11D3-B996-00104B33F2C4}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IDA2 = interface;
  IDA2Disp = dispinterface;
  DA2Unknown = interface;
  DA2UnknownDisp = dispinterface;
  IOPCGroup = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  OPCServer = IDA2;
  OPCGroup = IOPCGroup;


// *********************************************************************//
// Interface: IDA2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A42F19F2-608B-11D3-B98D-00403357BAA5}
// *********************************************************************//
  IDA2 = interface(IDispatch)
    ['{A42F19F2-608B-11D3-B98D-00403357BAA5}']
  end;

// *********************************************************************//
// DispIntf:  IDA2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A42F19F2-608B-11D3-B98D-00403357BAA5}
// *********************************************************************//
  IDA2Disp = dispinterface
    ['{A42F19F2-608B-11D3-B98D-00403357BAA5}']
  end;

// *********************************************************************//
// Interface: DA2Unknown
// Flags:     (320) Dual OleAutomation
// GUID:      {DEBAB081-6AE4-11D3-B995-00403357BAA5}
// *********************************************************************//
  DA2Unknown = interface(IUnknown)
    ['{DEBAB081-6AE4-11D3-B995-00403357BAA5}']
  end;

// *********************************************************************//
// DispIntf:  DA2UnknownDisp
// Flags:     (320) Dual OleAutomation
// GUID:      {DEBAB081-6AE4-11D3-B995-00403357BAA5}
// *********************************************************************//
  DA2UnknownDisp = dispinterface
    ['{DEBAB081-6AE4-11D3-B995-00403357BAA5}']
  end;

// *********************************************************************//
// Interface: IOPCGroup
// Flags:     (0)
// GUID:      {B1779B97-71B6-11D3-B996-00104B33F2C4}
// *********************************************************************//
  IOPCGroup = interface(IUnknown)
    ['{B1779B97-71B6-11D3-B996-00104B33F2C4}']
  end;

// *********************************************************************//
// The Class CoOPCServer provides a Create and CreateRemote method to          
// create instances of the default interface IDA2 exposed by              
// the CoClass OPCServer. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOPCServer = class
    class function Create: IDA2;
    class function CreateRemote(const MachineName: string): IDA2;
  end;

// *********************************************************************//
// The Class CoOPCGroup provides a Create and CreateRemote method to          
// create instances of the default interface IOPCGroup exposed by              
// the CoClass OPCGroup. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOPCGroup = class
    class function Create: IOPCGroup;
    class function CreateRemote(const MachineName: string): IOPCGroup;
  end;

implementation

uses ComObj;

class function CoOPCServer.Create: IDA2;
begin
  Result := CreateComObject(CLASS_OPCServer) as IDA2;
end;

class function CoOPCServer.CreateRemote(const MachineName: string): IDA2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCServer) as IDA2;
end;

class function CoOPCGroup.Create: IOPCGroup;
begin
  Result := CreateComObject(CLASS_OPCGroup) as IOPCGroup;
end;

class function CoOPCGroup.CreateRemote(const MachineName: string): IOPCGroup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCGroup) as IOPCGroup;
end;

end.
unit MicroNet_TLB;

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

// PASTLWTR : $Revision:   1.88.1.0.1.0  $
// File generated on 30.09.02 15:37:22 from Type Library described below.

// ************************************************************************ //
// Type Lib: C:\Arbeit\Коксохим\MicroNet OPC-Server\MicroNet.tlb (1)
// IID\LCID: {A42F19F1-608B-11D3-B98D-00403357BAA5}\0
// Helpfile: 
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINNT\System32\StdOle2.tlb)
//   (2) v4.0 StdVCL, (C:\WINNT\System32\STDVCL40.DLL)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
interface

uses Windows, ActiveX, Classes, Graphics, OleServer, OleCtrls, StdVCL;

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  MicroNetMajorVersion = 2;
  MicroNetMinorVersion = 0;

  LIBID_MicroNet: TGUID = '{A42F19F1-608B-11D3-B98D-00403357BAA5}';

  IID_IDA2: TGUID = '{A42F19F2-608B-11D3-B98D-00403357BAA5}';
  CLASS_OPCServer: TGUID = '{A42F19F4-608B-11D3-B98D-00403357BAA5}';
  IID_DA2Unknown: TGUID = '{DEBAB081-6AE4-11D3-B995-00403357BAA5}';
  IID_IOPCGroup: TGUID = '{B1779B97-71B6-11D3-B996-00104B33F2C4}';
  CLASS_OPCGroup: TGUID = '{B1779B99-71B6-11D3-B996-00104B33F2C4}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IDA2 = interface;
  IDA2Disp = dispinterface;
  DA2Unknown = interface;
  DA2UnknownDisp = dispinterface;
  IOPCGroup = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  OPCServer = IDA2;
  OPCGroup = IOPCGroup;


// *********************************************************************//
// Interface: IDA2
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A42F19F2-608B-11D3-B98D-00403357BAA5}
// *********************************************************************//
  IDA2 = interface(IDispatch)
    ['{A42F19F2-608B-11D3-B98D-00403357BAA5}']
  end;

// *********************************************************************//
// DispIntf:  IDA2Disp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A42F19F2-608B-11D3-B98D-00403357BAA5}
// *********************************************************************//
  IDA2Disp = dispinterface
    ['{A42F19F2-608B-11D3-B98D-00403357BAA5}']
  end;

// *********************************************************************//
// Interface: DA2Unknown
// Flags:     (320) Dual OleAutomation
// GUID:      {DEBAB081-6AE4-11D3-B995-00403357BAA5}
// *********************************************************************//
  DA2Unknown = interface(IUnknown)
    ['{DEBAB081-6AE4-11D3-B995-00403357BAA5}']
  end;

// *********************************************************************//
// DispIntf:  DA2UnknownDisp
// Flags:     (320) Dual OleAutomation
// GUID:      {DEBAB081-6AE4-11D3-B995-00403357BAA5}
// *********************************************************************//
  DA2UnknownDisp = dispinterface
    ['{DEBAB081-6AE4-11D3-B995-00403357BAA5}']
  end;

// *********************************************************************//
// Interface: IOPCGroup
// Flags:     (0)
// GUID:      {B1779B97-71B6-11D3-B996-00104B33F2C4}
// *********************************************************************//
  IOPCGroup = interface(IUnknown)
    ['{B1779B97-71B6-11D3-B996-00104B33F2C4}']
  end;

// *********************************************************************//
// The Class CoOPCServer provides a Create and CreateRemote method to          
// create instances of the default interface IDA2 exposed by              
// the CoClass OPCServer. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOPCServer = class
    class function Create: IDA2;
    class function CreateRemote(const MachineName: string): IDA2;
  end;

// *********************************************************************//
// The Class CoOPCGroup provides a Create and CreateRemote method to          
// create instances of the default interface IOPCGroup exposed by              
// the CoClass OPCGroup. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoOPCGroup = class
    class function Create: IOPCGroup;
    class function CreateRemote(const MachineName: string): IOPCGroup;
  end;

implementation

uses ComObj;

class function CoOPCServer.Create: IDA2;
begin
  Result := CreateComObject(CLASS_OPCServer) as IDA2;
end;

class function CoOPCServer.CreateRemote(const MachineName: string): IDA2;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCServer) as IDA2;
end;

class function CoOPCGroup.Create: IOPCGroup;
begin
  Result := CreateComObject(CLASS_OPCGroup) as IOPCGroup;
end;

class function CoOPCGroup.CreateRemote(const MachineName: string): IOPCGroup;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_OPCGroup) as IOPCGroup;
end;

end.
