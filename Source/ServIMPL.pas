unit ServIMPL;

interface

uses Windows,ComObj,ActiveX,Axctrls,MicroNet_TLB,OPCDA,SysUtils,Dialogs,Classes,
     OPCCOMN,StdVCL,enumstring,Globals,OPCErrorStrings, EnumUnknown, OPCTypes;

type
  TOPCItemProp = class
  public
   function QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                          out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                          out ppvtDataTypes:PVarTypeList):HResult;stdcall;
   function GetItemProperties(szItemID:POleStr;
                              dwCount:DWORD;
                              pdwPropertyIDs:PDWORDARRAY;
                              out ppvData:POleVariantArray;
                              out ppErrors:PResultList):HResult;stdcall;
   function LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                           out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
end;


  TDA2 = class(TAutoObject,IDA2,IOPCServer,IOPCCommon,IOPCServerPublicGroups,
               IOPCBrowseServerAddressSpace,IPersist,IPersistFile,
               IConnectionPointContainer,IOPCItemProperties)
  private
   FIOPCItemProperties:TOPCItemProp;
   FIConnectionPoints:TConnectionPoints;
  protected
   property iFIConnectionPoints:TConnectionPoints read FIConnectionPoints
                          write FIConnectionPoints implements IConnectionPointContainer;
//IOPCServer begin
    function AddGroup(szName:POleStr;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                      hClientGroup:OPCHANDLE; pTimeBias:PLongint; pPercentDeadband:PSingle;
                      dwLCID:DWORD; out phServerGroup: OPCHANDLE;
                                    out pRevisedUpdateRate:DWORD;
                                    const riid: TIID;
                                    out ppUnk:IUnknown):HResult;stdcall;
    function GetErrorString(dwError:HResult; dwLocale:TLCID; out ppString:POleStr):HResult;overload; stdcall;
    function GetGroupByName(szName:POleStr; const riid: TIID; out ppUnk:IUnknown):HResult; stdcall;
    function GetStatus(out ppServerStatus:POPCSERVERSTATUS): HResult; stdcall;
    function RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
    function CreateGroupEnumerator(dwScope:OPCENUMSCOPE; const riid:TIID; out ppUnk:IUnknown):HResult; stdcall;
//IOPCServer end;

//IOPCCommon begin
    function SetLocaleID(dwLcid:TLCID):HResult;stdcall;
    function GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
    function QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
    function GetErrorString(dwError:HResult; out ppString:POleStr):HResult;overload;stdcall;
    function SetClientName(szName:POleStr):HResult;stdcall;
//IOPCCommon end

//IOPCServerPublicGroups begin
    function GetPublicGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult;stdcall;
    function RemovePublicGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
//IOPCServerPublicGroups end

//IOPCBrowseServerAddressSpace begin
   function QueryOrganization(out pNameSpaceType:OPCNAMESPACETYPE):HResult;stdcall;
   function ChangeBrowsePosition(dwBrowseDirection:OPCBROWSEDIRECTION;
                                 szString:POleStr):HResult;stdcall;
   function BrowseOPCItemIDs(dwBrowseFilterType:OPCBROWSETYPE; szFilterCriteria:POleStr;
                             vtDataTypeFilter:TVarType; dwAccessRightsFilter:DWORD;
                             out ppIEnumString:IEnumString):HResult;stdcall;
   function GetItemID(szItemDataID:POleStr; out szItemID:POleStr):HResult;stdcall;
   function BrowseAccessPaths(szItemID:POleStr; out ppIEnumString:IEnumString):HResult;stdcall;
//IOPCBrowseServerAddressSpace end

//IPersistFile begin
    function GetClassID(out classID: TCLSID):HResult;stdcall;
    function IsDirty:HResult;stdcall;
    function Load(pszFileName:POleStr; dwMode:Longint):HResult;stdcall;
    function Save(pszFileName:POleStr; fRemember:BOOL):HResult;stdcall;
    function SaveCompleted(pszFileName:POleStr):HResult;stdcall;
    function GetCurFile(out pszFileName:POleStr):HResult;stdcall;
//IPersistFile end
  public
   grps,pubGrps:TList;
   localID:longword;
   clientName,errString:string;
   srvStarted,lastClientUpdate:TDateTime;
   FOnSDConnect: TConnectEvent;
   ClientIUnknown:IUnknown;
   property iFIOPCItemProperties:TOPCItemProp read FIOPCItemProperties
                                              write FIOPCItemProperties
                                              implements IOPCItemProperties;

   procedure CreateGroups;
   procedure Initialize; override;
   procedure ShutdownOnConnect(const Sink: IUnknown; Connecting: Boolean);

   destructor Destroy;override;
   function GetNewGroupNumber:longword;
   function GetNewItemNumber:longword;
   function FindIndexViaGrpNumber(wGrp:TList;gNum:longword):integer;
   procedure GroupRemovingSelf(wGrp:TList;gNum:integer);
   function GetGroupCount(gList:TList):integer;
   function CreateGrpNameList(gList:TList):TStringList;
   function IsGroupNamePresent(gList:TList;theName:string):integer;
   function IsNameUsedInAnyGroup(theName:string):boolean;
   function IsThisGroupPublic(aList:TList):boolean;
   procedure TimeSlice(cTime:TDateTime);
   function CloneAGroup(szName:string;aGrp:TTypedComObject; out res:HResult):IUnknown;
  end;

var
 theServers:array [0..10] of TDA2;

implementation

uses ComServ,Main,GroupUnit;

function TOPCItemProp.QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                       out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                       out ppvtDataTypes:PVarTypeList):HResult;stdcall;
var
 memErr:boolean;
 propID:longword;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 pdwCount:=1;
 memErr:=false;
 ppPropertyIDs:=PDWORDARRAY(CoTaskMemAlloc(pdwCount*sizeof(DWORD)));
 if ppPropertyIDs = nil then
  memErr:=true;
 if not memErr then
  ppDescriptions:=POleStrList(CoTaskMemAlloc(pdwCount*sizeof(POleStr)));
 if ppDescriptions = nil then
  memErr:=true;
 if not memErr then
  ppvtDataTypes:=PVarTypeList(CoTaskMemAlloc(pdwCount*sizeof(TVarType)));
 if ppvtDataTypes = nil then
  memErr:=true;

 if memErr then
  begin
   if ppPropertyIDs <> nil then  CoTaskMemFree(ppPropertyIDs);
   if ppDescriptions <> nil then  CoTaskMemFree(ppDescriptions);
   if ppvtDataTypes <> nil then  CoTaskMemFree(ppvtDataTypes);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ppPropertyIDs[0]:=propID;
 ppDescriptions[0]:=StringToLPOLESTR(ReturnTagnameFromPropID(propID));
 ppvtDataTypes[0]:=ReturnDataTypeFromPropID(propID);
 result:=S_OK;
end;

function TOPCItemProp.GetItemProperties(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                      out ppvData:POleVariantArray; out ppErrors:PResultList):HResult;stdcall;
var
 data:variant;
 i:integer;
 memErr:boolean;
 propID:longword;
 ppArray:PDWORDARRAY;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 memErr:=false;
 ppvData:=POleVariantArray(CoTaskMemAlloc(dwCount * sizeof(OleVariant)));
 if ppvData = nil then
  memErr:=true;
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount * sizeof(HRESULT)));
 if ppErrors = nil then
  memErr:=true;

 if memErr then
  begin
   if ppvData <> nil then  CoTaskMemFree(ppvData);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ppArray:=@pdwPropertyIDs^;
 result:=S_OK;
 for i:= 0 to dwCount-1 do
  begin
   case ppArray[i] of
    1:  data:=ReturnDataTypeFromPropID(propID);   //ppArray[i]
    2:  data:=MicroNetOPCForm.ReturnItemValue(ppArray[i]);
    5:
     begin
      if CanPropIDBeWritten(propID) then
       data:=OPC_READABLE or OPC_WRITEABLE
      else
       data:=OPC_READABLE;
     end;
    //5000 - 5022
//    posItems[low(posItems)].PropID..posItems[high(posItems)].PropID:
//    5000..5022:
//     data:=MicroNetOPCForm.ReturnItemValue(ppArray[i] - posItems[low(posItems)].PropID);

    else
     begin
      ppErrors[i]:=OPC_E_INVALID_PID;
      result:=S_FALSE;
      Continue;
     end;
   end;
   ppvData[i]:=data;
   ppErrors[i]:=S_OK;
  end;

end;

function TOPCItemProp.LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                      out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
var
 i:integer;
 propID:longword;
 memErr:boolean;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 memErr:=false;
 ppszNewItemIDs:=POleStrList(CoTaskMemAlloc(dwCount*sizeof(POleStr)));
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  memErr:=true;

 if memErr then
  begin
   if ppszNewItemIDs <> nil then  CoTaskMemFree(ppszNewItemIDs);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 for i:= 0 to dwCount-1 do
  begin
   ppszNewItemIDs[i]:=StringToLPOLESTR(szItemID);
   ppErrors[i]:=S_OK;
  end;

 result:=S_OK;
end;

//{$INCLUDE IOPCServerIMPL}
function TDA2.AddGroup(szName:POleStr;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                  hClientGroup:OPCHANDLE; pTimeBias:PLongint; pPercentDeadband:PSingle;
                  dwLCID:DWORD; out phServerGroup: OPCHANDLE;
                                out pRevisedUpdateRate:DWORD;
                                const riid: TIID;
                                out ppUnk:IUnknown):HResult;stdcall;
var
 s1:string;
 i:longint;
 aGrp:TOPCGroup;
 perDeadband:single;
begin
 result:=S_OK;
 s1:=szName;
 i:=0;

 if (s1 = '') then
  repeat
   s1:=s1 + IntToStr(GetTickCount);
   i:=succ(i);
  until (not IsNameUsedInAnyGroup(s1)) or (i > 9);

 if i > 9 then
  begin
   result:=OPC_E_DUPLICATENAME;
   phServerGroup:=0;
   Exit;
  end;

 if IsNameUsedInAnyGroup(s1) then
  begin
   result:=OPC_E_DUPLICATENAME;
   phServerGroup:=0;
   Exit;
  end;

 aGrp:=TOPCGroup.Create(self,grps);
 if aGrp = nil then
  begin
   result:=E_OUTOFMEMORY;
   phServerGroup:=0;
   Exit;
  end;

 grps.Add(aGrp);
 phServerGroup:=GetNewGroupNumber;
 if assigned(pPercentDeadband) then
  perDeadband:=pPercentDeadband^
 else
  perDeadband:=0;

 aGrp.SetUp(s1,bActive,dwRequestedUpdateRate,hClientGroup,0,
            perDeadband,dwLCID,phServerGroup);

 aGrp.ValidateTimeBias(pTimeBias);

 if dwRequestedUpdateRate <> aGrp.requestedUpdateRate then
  result:=OPC_S_UNSUPPORTEDRATE;

 pRevisedUpdateRate:=aGrp.requestedUpdateRate;

 MicroNetOPCForm.UpdateGroupCount;
 ppUnk:=aGrp;
end;

function TDA2.GetErrorString(dwError:HResult; dwLocale:TLCID; out ppString:POleStr):HResult; stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TDA2.GetGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult; stdcall;
var
 i,gNum:integer;
begin
 gNum:=IsGroupNamePresent(grps,szName);     //returns the index to the groups list
 if (addr(ppUnk) = nil) or (gNum = -1)then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 i:=gNum;
// i:=FindIndexViaGrpNumber(grps,gNum);
 if (i = -1) or (i > (grps.count-1) )  then
  begin
   Result:=E_FAIL;              Exit;
  end;
 result:=IUnknown(TOPCGroup(grps[i])).QueryInterface(riid,ppUnk);
end;

function TDA2.GetStatus(out ppServerStatus:POPCSERVERSTATUS):HResult;stdcall;
var
 aFileTime:TFileTime;
begin
 if (addr(ppServerStatus) = nil) then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 result:=S_OK;
 ppServerStatus:=POPCSERVERSTATUS(CoTaskMemAlloc(sizeof(OPCSERVERSTATUS)));
 if ppServerStatus = nil then
  begin
   ppServerStatus:=nil;         result:=E_OUTOFMEMORY;   Exit;
  end;

 DataTimeToOPCTime(srvStarted,aFileTime);
 CoFiletimeNow(ppServerStatus.ftCurrentTime);
 DataTimeToOPCTime(lastClientUpdate,ppServerStatus.ftLastUpdateTime);

 ppServerStatus.dwServerState:=OPC_STATUS_RUNNING;
 ppServerStatus.dwGroupCount:=GetGroupCount(grps) + GetGroupCount(pubGrps);
 ppServerStatus.dwBandWidth:=100;
 ppServerStatus.wMajorVersion:=1;
 ppServerStatus.wMinorVersion:=2;
 ppServerStatus.wBuildNumber:=5;
 ppServerStatus.szVendorInfo:=StringToLPOLESTR('MRD');
end;

function TDA2.RemoveGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
var
 i:integer;
 aGrp:TOPCGroup;
begin
 result:=S_OK;
 if hServerGroup < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 i:=FindIndexViaGrpNumber(grps,hServerGroup);
 if i = -1 then            //the client as already freed the group
  Exit;                    //and we deleted it from the server list in the group destroy

 aGrp:=grps[i];
 if (aGrp.RefCount > 2) and not bForce then
  begin
   result:=OPC_S_INUSE;
   Exit;
  end;

 GroupRemovingSelf(grps,hServerGroup);
end;

function TDA2.CreateGroupEnumerator(dwScope:OPCENUMSCOPE; const riid:TIID;
                                         out ppUnk:IUnknown):HResult; stdcall;

 procedure EnumerateForStrings;
 var
  i:integer;
  pvList,pubList:TStringList;
 begin
  pvList:=nil;               pubList:=nil;
  try
   case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
     pvList:=CreateGrpNameList(grps);
    OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
     pubList:=CreateGrpNameList(pubGrps);
    OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
     begin
      pvList:=CreateGrpNameList(grps);
      pubList:=CreateGrpNameList(pubGrps);
     end
    else
     begin
      result:=E_INVALIDARG;
      Exit;
     end;
   end;

   if pvList = nil then
    pvList:=TStringList.Create;
   if pubList <> nil then
    if pubList.count > 0 then
     for i:= 0 to pubList.count-1 do
      pvList.Add(pubList[i]);

   result:=S_OK;
   ppUnk:=TOPCStringsEnumerator.Create(pvList);
  finally
   pvList.Free;
   pubList.Free;
  end;
 end;


 procedure EumerateForUnknown;
 var
  aList:TList;

 procedure AddAGroup(inList:TList);
 var
  i:integer;
  Obj:Pointer;
 begin
  for i:= 0 to inList.count -1 do
   begin
    Obj:=nil;
    IUnknown(TOPCGroup(inList[i])).QueryInterface(IUnknown,Obj);
    if Assigned(Obj) then
     aList.Add(Obj);
   end;
 end;

 begin
  aList:=nil;
  try
   aList:=TList.Create;
   case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
     if Assigned(grps) then
      AddAGroup(grps);
    OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
     if Assigned(pubGrps) then
      AddAGroup(pubGrps);
    OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
     begin
     if Assigned(grps) then
      AddAGroup(grps);
     if Assigned(pubGrps) then
      AddAGroup(pubGrps);
     end
    else
     begin
      result:=E_INVALIDARG;
      Exit;
     end;
   end;

   if Assigned(aList) and (aList.count > 0) then
    ppUnk:=TS3UnknownEnumerator.Create(aList);
  finally
   aList.Free;
  end;
  result:=S_OK;
 end;

begin
 if not (IsEqualIID(riid,IEnumUnknown) or IsEqualIID(riid,IEnumString)) then
  begin
   result:=E_NOINTERFACE;
   Exit;
  end;

 if IsEqualIID(riid,IEnumString) then
  EnumerateForStrings
 else if IsEqualIID(riid,IEnumUnknown) then
  EumerateForUnknown
 else
  result:=E_FAIL;

end;


//{$INCLUDE IOPCCommonIMPL}
function TDA2.SetLocaleID(dwLcid:TLCID):HResult;stdcall;
begin
 if (dwLcid = LOCALE_SYSTEM_DEFAULT) or (dwLcid = LOCALE_USER_DEFAULT) then
  begin
   localID:=dwLcid;
   result:=S_OK;
  end
 else
  result:=E_INVALIDARG;
end;

function TDA2.GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
begin
 pdwLcid:=localID;
 result:=S_OK;
end;

function TDA2.QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
begin
 pdwCount:=2;
 pdwLcid:=PLCIDARRAY(CoTaskMemAlloc(pdwCount*sizeof(LCID)));
 if (pdwLcid = nil) then
  begin
   if pdwLcid <> nil then  CoTaskMemFree(pdwLcid);
   result:=E_OUTOFMEMORY;
   Exit;
  end;
 pdwLcid[0]:=LOCALE_SYSTEM_DEFAULT;
 pdwLcid[1]:=LOCALE_USER_DEFAULT;
 result:=S_OK;
end;

function TDA2.GetErrorString(dwError:HResult; out ppString:POleStr):HResult;stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TDA2.SetClientName(szName:POleStr):HResult;stdcall;
begin
 if (addr(szName) = nil) then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 clientName:=szName;
 result:=S_OK;
end;

//{$INCLUDE IOPCServerPublicGroupsIMPL}
function TDA2.GetPublicGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult;stdcall;
begin
 if IsGroupNamePresent(pubGrps,szName) <> -1 then
  begin
   result:=S_OK;
   ppUnk:=self;
  end
 else
 result:=OPC_E_NOTFOUND;
end;

function TDA2.RemovePublicGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
begin
 //do something with the forse if needed
 GroupRemovingSelf(pubGrps,hServerGroup);
 result:=S_OK;
end;

//{$INCLUDE IOPCBrowseServerAddressSpaceIMPL}
function TDA2.QueryOrganization(out pNameSpaceType:OPCNAMESPACETYPE):HResult;stdcall;
begin
 pNameSpaceType:=OPC_NS_FLAT;
 result:=S_OK;
end;

function TDA2.ChangeBrowsePosition(dwBrowseDirection:OPCBROWSEDIRECTION;
                              szString:POleStr):HResult;stdcall;
begin
 result:=E_FAIL
end;

function TDA2.BrowseOPCItemIDs(dwBrowseFilterType:OPCBROWSETYPE; szFilterCriteria:POleStr;
                          vtDataTypeFilter:TVarType; dwAccessRightsFilter:DWORD;
                          out ppIEnumString:IEnumString):HResult;stdcall;
var
 i:integer;
 tList:TStringList;
begin
//add filter support
 result:=S_OK;
 tList:=nil;
 try
  tList:=TStringList.Create;
  if tList = nil then
   begin
    result:=E_OUTOFMEMORY;
    Exit;
   end;

  for i:= low(posItems) to high(posItems) do
   tList.Add(posItems[i].tagname);

  ppIEnumString:=TOPCStringsEnumerator.Create(tList);
 finally
  tList.Free;
 end;
end;

function TDA2.GetItemID(szItemDataID:POleStr; out szItemID:POleStr):HResult;stdcall;
var
 propID:integer;
begin
 result:=S_OK;
 if length(szItemDataID) = 0 then
  szItemID:=StringToLPOLESTR(szItemDataID)
 else
  begin
   propID:=ReturnPropIDFromTagname(szItemDataID);
   if propID = 0 then
    result:=OPC_E_UNKNOWNITEMID
   else
    szItemID:=StringToLPOLESTR(szItemDataID);
  end;
end;

function TDA2.BrowseAccessPaths(szItemID:POleStr; out ppIEnumString:IEnumString):HResult;stdcall;
begin
 result:=E_NOTIMPL;
end;

//{$INCLUDE IPersistFileIMPL}
function TDA2.GetClassID(out classID:TCLSID):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.IsDirty:HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.Load(pszFileName:POleStr; dwMode:Longint):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.Save(pszFileName:POleStr; fRemember:BOOL):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.SaveCompleted(pszFileName:POleStr):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.GetCurFile(out pszFileName:POleStr):HResult;stdcall;
begin
 result:=S_FALSE;
end;


function GetNextFreeServerSpot:integer;
var
 i:integer;
begin
 result:=-1;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] = nil then
   begin
    result:=i;
    Exit;
   end;
end;

function FindServerInArray(which:TDA2):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   if theServers[i] = which then
   begin
    result:=i;
    Exit;
   end;
end;

function ReturnServerCount:integer;
var
 i:integer;
begin
 result:=0;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   result:=succ(result);
end;

procedure KillServers;
var
 i:integer;
begin
 for i:= high(theServers) downTo low(theServers) do
  begin
   CoDisconnectObject(TDA2(theServers[i]) as IUnknown,0);
   TDA2(theServers[i]).Free;
  end;
 FreeAndNil(theServers);
end;

procedure TDA2.CreateGroups;
begin
 if grps <> nil then Exit;
 grps:=TList.Create;
 grps.Capacity:=255;
 pubGrps:=TList.Create;
 pubGrps.Capacity:=grps.Capacity;
end;

procedure TDA2.Initialize;
var
 i:integer;
begin
 i:=GetNextFreeServerSpot;
 if i = -1 then Exit;
 inherited Initialize;
 srvStarted:=Now;
 lastClientUpdate:=0;
 localID:=LOCALE_SYSTEM_DEFAULT;

 FIConnectionPoints:=TConnectionPoints.Create(self);
 FIOPCItemProperties:=TOPCItemProp.Create;

 FOnSDConnect:=ShutdownOnConnect;
 FIConnectionPoints.CreateConnectionPoint(IID_IOPCShutdown,ckSingle,FOnSDConnect);

 CreateGroups;

 //hook into Main program here    may have multiple servers
 theServers[i]:=self;
 MicroNetOPCForm.UpdateGroupCount;
end;

procedure TDA2.ShutdownOnConnect(const Sink: IUnknown; Connecting: Boolean);
begin
 if connecting then
  ClientIUnknown:=Sink
 else
  ClientIUnknown:=nil
end;

destructor TDA2.Destroy;
var
 i:integer;
begin
 if grps <> nil then
  for i:= 0 to grps.count-1 do
   TOPCGroup(grps.Items[i]).Free;
 grps.Free;
 if pubGrps <> nil then
  for i:= 0 to pubGrps.count-1 do
   TOPCGroup(pubGrps.Items[i]).Free;
 pubGrps.Free;
 i:=FindServerInArray(self);
 if i <> -1 then
  theServers[i]:=nil;                //the client has let us be free ;)

 if Assigned(FIConnectionPoints) then                   FIConnectionPoints.Free;
 if Assigned(FIOPCItemProperties) then                  FIOPCItemProperties.Free;
 MicroNetOPCForm.UpdateGroupCount;
 Inherited;
end;

function TDA2.GetNewGroupNumber:longword;
const
 grpIndex:longword = 1;             //Assignable Typed Constants gota lovem
begin
 grpIndex:=succ(grpIndex);         //get us a new reference number
 result:=grpIndex;
end;

function TDA2.GetNewItemNumber:longword;
const
 itemIndex:longword = 1;
begin
 itemIndex:=succ(itemIndex);
 result:=itemIndex;
end;

function TDA2.FindIndexViaGrpNumber(wGrp:TList;gNum:longword):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to wGrp.count-1 do
  if TOPCGroup(wGrp[i]).serverHandle = gNum then
   begin
    result:=i;
    Break;
   end;
end;

procedure TDA2.GroupRemovingSelf(wGrp:TList; gNum:integer);
var
 i:integer;
begin
 i:=FindIndexViaGrpNumber(wGrp,gNum);
 if (i <> -1) then
  wGrp.Delete(i);
 MicroNetOPCForm.UpdateGroupCount;
end;

function TDA2.GetGroupCount(gList:TList):integer;
begin
 result:=0;
 if gList = nil then Exit;
 result:=gList.count;
end;

function TDA2.CreateGrpNameList(gList:TList):TStringList;
var
 i:integer;
begin
 result:=nil;
 if gList = nil then Exit;
 result:=TStringList.Create;
 for i:= 0 to gList.count-1 do
  result.Add(TOPCGroup(gList.Items[i]).tagName);
 if result.count = 0 then
  begin
   result.Free;
   result:=nil;
  end;
end;

function TDA2.IsGroupNamePresent(gList:TList; theName:string):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to gList.count-1 do
  if theName = TOPCGroup(gList.Items[i]).tagName then
    begin
     result:=i;            Break;
    end;
end;

function TDA2.IsNameUsedInAnyGroup(theName:string):boolean;
var
 i:integer;
begin
 result:=false;
 i:=IsGroupNamePresent(grps,theName);
 if i <> -1 then
  begin
   result:=true;                     Exit;
  end;
 i:=IsGroupNamePresent(pubGrps,theName);
 if i <> -1 then
  result:=true;
end;

function TDA2.IsThisGroupPublic(aList:TList):boolean;
begin
 result:=boolean((pubGrps <> nil)    and
                 (pubGrps.count = 0) and
                 (aList <> nil)      and
                 (pubGrps = aList));
end;

procedure TDA2.TimeSlice(cTime:TDateTime);
var
 i:integer;
begin
 lastClientUpdate:=cTime;
 if Assigned(grps) then
  for i:= 0 to grps.count-1 do
   TOPCGroup(grps.Items[i]).TimeSlice(cTime);
 if Assigned(pubGrps) then
  for i:= 0 to pubGrps.count-1 do
   TOPCGroup(pubGrps.Items[i]).TimeSlice(cTime);
end;

function TDA2.CloneAGroup(szName:string;aGrp:TTypedComObject; out res:HResult):IUnknown;
var
 sGrp,dGrp:TOPCGroup;
begin
 sGrp:=TOPCGroup(aGrp);
 dGrp:=TOPCGroup.Create(self,grps);
 if dGrp = nil then
  begin
   result:=nil;
   res:=E_OUTOFMEMORY;      Exit;
  end;

 grps.Add(dGrp);
 dGrp.tagName:=szName;
 sGrp.CloneYourSelf(dGrp);
 result:=dGrp;
 res:=S_OK;
end;

initialization
  TAutoObjectFactory.Create(ComServer, TDA2, Class_OPCServer,
                            ciMultiInstance, tmApartment);

 //if an OPC client(s) is connected and the user has selected to quit after
 //the warning in the FormCloseQuery then do not let the system ask again in the
 //AutomationTerminateProc procedure in the VCL.
 ComServer.UIInteractive:=false;
finalization
 //if an OPC client is connected and this is a forced kill then if CoUnintialize
 //is not called here the OLE dll will generate an error when it is called after
 //we have killed the servers.
 CoUninitialize;
 KillServers;

end.
unit ServIMPL;

interface

uses Windows,ComObj,ActiveX,Axctrls,MicroNet_TLB,OPCDA,SysUtils,Dialogs,Classes,
     OPCCOMN,StdVCL,enumstring,Globals,OPCErrorStrings, EnumUnknown, OPCTypes;

type
  TOPCItemProp = class
  public
   function QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                          out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                          out ppvtDataTypes:PVarTypeList):HResult;stdcall;
   function GetItemProperties(szItemID:POleStr;
                              dwCount:DWORD;
                              pdwPropertyIDs:PDWORDARRAY;
                              out ppvData:POleVariantArray;
                              out ppErrors:PResultList):HResult;stdcall;
   function LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                           out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
end;


  TDA2 = class(TAutoObject,IDA2,IOPCServer,IOPCCommon,IOPCServerPublicGroups,
               IOPCBrowseServerAddressSpace,IPersist,IPersistFile,
               IConnectionPointContainer,IOPCItemProperties)
  private
   FIOPCItemProperties:TOPCItemProp;
   FIConnectionPoints:TConnectionPoints;
  protected
   property iFIConnectionPoints:TConnectionPoints read FIConnectionPoints
                          write FIConnectionPoints implements IConnectionPointContainer;
//IOPCServer begin
    function AddGroup(szName:POleStr;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                      hClientGroup:OPCHANDLE; pTimeBias:PLongint; pPercentDeadband:PSingle;
                      dwLCID:DWORD; out phServerGroup: OPCHANDLE;
                                    out pRevisedUpdateRate:DWORD;
                                    const riid: TIID;
                                    out ppUnk:IUnknown):HResult;stdcall;
    function GetErrorString(dwError:HResult; dwLocale:TLCID; out ppString:POleStr):HResult;overload; stdcall;
    function GetGroupByName(szName:POleStr; const riid: TIID; out ppUnk:IUnknown):HResult; stdcall;
    function GetStatus(out ppServerStatus:POPCSERVERSTATUS): HResult; stdcall;
    function RemoveGroup(hServerGroup: OPCHANDLE; bForce: BOOL): HResult; stdcall;
    function CreateGroupEnumerator(dwScope:OPCENUMSCOPE; const riid:TIID; out ppUnk:IUnknown):HResult; stdcall;
//IOPCServer end;

//IOPCCommon begin
    function SetLocaleID(dwLcid:TLCID):HResult;stdcall;
    function GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
    function QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
    function GetErrorString(dwError:HResult; out ppString:POleStr):HResult;overload;stdcall;
    function SetClientName(szName:POleStr):HResult;stdcall;
//IOPCCommon end

//IOPCServerPublicGroups begin
    function GetPublicGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult;stdcall;
    function RemovePublicGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
//IOPCServerPublicGroups end

//IOPCBrowseServerAddressSpace begin
   function QueryOrganization(out pNameSpaceType:OPCNAMESPACETYPE):HResult;stdcall;
   function ChangeBrowsePosition(dwBrowseDirection:OPCBROWSEDIRECTION;
                                 szString:POleStr):HResult;stdcall;
   function BrowseOPCItemIDs(dwBrowseFilterType:OPCBROWSETYPE; szFilterCriteria:POleStr;
                             vtDataTypeFilter:TVarType; dwAccessRightsFilter:DWORD;
                             out ppIEnumString:IEnumString):HResult;stdcall;
   function GetItemID(szItemDataID:POleStr; out szItemID:POleStr):HResult;stdcall;
   function BrowseAccessPaths(szItemID:POleStr; out ppIEnumString:IEnumString):HResult;stdcall;
//IOPCBrowseServerAddressSpace end

//IPersistFile begin
    function GetClassID(out classID: TCLSID):HResult;stdcall;
    function IsDirty:HResult;stdcall;
    function Load(pszFileName:POleStr; dwMode:Longint):HResult;stdcall;
    function Save(pszFileName:POleStr; fRemember:BOOL):HResult;stdcall;
    function SaveCompleted(pszFileName:POleStr):HResult;stdcall;
    function GetCurFile(out pszFileName:POleStr):HResult;stdcall;
//IPersistFile end
  public
   grps,pubGrps:TList;
   localID:longword;
   clientName,errString:string;
   srvStarted,lastClientUpdate:TDateTime;
   FOnSDConnect: TConnectEvent;
   ClientIUnknown:IUnknown;
   property iFIOPCItemProperties:TOPCItemProp read FIOPCItemProperties
                                              write FIOPCItemProperties
                                              implements IOPCItemProperties;

   procedure CreateGroups;
   procedure Initialize; override;
   procedure ShutdownOnConnect(const Sink: IUnknown; Connecting: Boolean);

   destructor Destroy;override;
   function GetNewGroupNumber:longword;
   function GetNewItemNumber:longword;
   function FindIndexViaGrpNumber(wGrp:TList;gNum:longword):integer;
   procedure GroupRemovingSelf(wGrp:TList;gNum:integer);
   function GetGroupCount(gList:TList):integer;
   function CreateGrpNameList(gList:TList):TStringList;
   function IsGroupNamePresent(gList:TList;theName:string):integer;
   function IsNameUsedInAnyGroup(theName:string):boolean;
   function IsThisGroupPublic(aList:TList):boolean;
   procedure TimeSlice(cTime:TDateTime);
   function CloneAGroup(szName:string;aGrp:TTypedComObject; out res:HResult):IUnknown;
  end;

var
 theServers:array [0..10] of TDA2;

implementation

uses ComServ,Main,GroupUnit;

function TOPCItemProp.QueryAvailableProperties(szItemID:POleStr; out pdwCount:DWORD;
                       out ppPropertyIDs:PDWORDARRAY; out ppDescriptions:POleStrList;
                       out ppvtDataTypes:PVarTypeList):HResult;stdcall;
var
 memErr:boolean;
 propID:longword;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 pdwCount:=1;
 memErr:=false;
 ppPropertyIDs:=PDWORDARRAY(CoTaskMemAlloc(pdwCount*sizeof(DWORD)));
 if ppPropertyIDs = nil then
  memErr:=true;
 if not memErr then
  ppDescriptions:=POleStrList(CoTaskMemAlloc(pdwCount*sizeof(POleStr)));
 if ppDescriptions = nil then
  memErr:=true;
 if not memErr then
  ppvtDataTypes:=PVarTypeList(CoTaskMemAlloc(pdwCount*sizeof(TVarType)));
 if ppvtDataTypes = nil then
  memErr:=true;

 if memErr then
  begin
   if ppPropertyIDs <> nil then  CoTaskMemFree(ppPropertyIDs);
   if ppDescriptions <> nil then  CoTaskMemFree(ppDescriptions);
   if ppvtDataTypes <> nil then  CoTaskMemFree(ppvtDataTypes);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ppPropertyIDs[0]:=propID;
 ppDescriptions[0]:=StringToLPOLESTR(ReturnTagnameFromPropID(propID));
 ppvtDataTypes[0]:=ReturnDataTypeFromPropID(propID);
 result:=S_OK;
end;

function TOPCItemProp.GetItemProperties(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                      out ppvData:POleVariantArray; out ppErrors:PResultList):HResult;stdcall;
var
 data:variant;
 i:integer;
 memErr:boolean;
 propID:longword;
 ppArray:PDWORDARRAY;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 memErr:=false;
 ppvData:=POleVariantArray(CoTaskMemAlloc(dwCount * sizeof(OleVariant)));
 if ppvData = nil then
  memErr:=true;
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount * sizeof(HRESULT)));
 if ppErrors = nil then
  memErr:=true;

 if memErr then
  begin
   if ppvData <> nil then  CoTaskMemFree(ppvData);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 ppArray:=@pdwPropertyIDs^;
 result:=S_OK;
 for i:= 0 to dwCount-1 do
  begin
   case ppArray[i] of
    1:  data:=ReturnDataTypeFromPropID(propID);   //ppArray[i]
    2:  data:=MicroNetOPCForm.ReturnItemValue(ppArray[i]);
    5:
     begin
      if CanPropIDBeWritten(propID) then
       data:=OPC_READABLE or OPC_WRITEABLE
      else
       data:=OPC_READABLE;
     end;
    //5000 - 5022
//    posItems[low(posItems)].PropID..posItems[high(posItems)].PropID:
//    5000..5022:
//     data:=MicroNetOPCForm.ReturnItemValue(ppArray[i] - posItems[low(posItems)].PropID);

    else
     begin
      ppErrors[i]:=OPC_E_INVALID_PID;
      result:=S_FALSE;
      Continue;
     end;
   end;
   ppvData[i]:=data;
   ppErrors[i]:=S_OK;
  end;

end;

function TOPCItemProp.LookupItemIDs(szItemID:POleStr; dwCount:DWORD; pdwPropertyIDs:PDWORDARRAY;
                      out ppszNewItemIDs:POleStrList;out ppErrors:PResultList): HResult; stdcall;
var
 i:integer;
 propID:longword;
 memErr:boolean;
begin
 propID:=ReturnPropIDFromTagname(szItemID);
 if propID = 0 then
  begin
   result:=OPC_E_INVALIDITEMID;     Exit;
  end;

 memErr:=false;
 ppszNewItemIDs:=POleStrList(CoTaskMemAlloc(dwCount*sizeof(POleStr)));
 if not memErr then
  ppErrors:=PResultList(CoTaskMemAlloc(dwCount*sizeof(HRESULT)));
 if ppErrors = nil then
  memErr:=true;

 if memErr then
  begin
   if ppszNewItemIDs <> nil then  CoTaskMemFree(ppszNewItemIDs);
   if ppErrors <> nil then  CoTaskMemFree(ppErrors);
   result:=E_OUTOFMEMORY;
   Exit;
  end;

 for i:= 0 to dwCount-1 do
  begin
   ppszNewItemIDs[i]:=StringToLPOLESTR(szItemID);
   ppErrors[i]:=S_OK;
  end;

 result:=S_OK;
end;

//{$INCLUDE IOPCServerIMPL}
function TDA2.AddGroup(szName:POleStr;bActive:BOOL; dwRequestedUpdateRate:DWORD;
                  hClientGroup:OPCHANDLE; pTimeBias:PLongint; pPercentDeadband:PSingle;
                  dwLCID:DWORD; out phServerGroup: OPCHANDLE;
                                out pRevisedUpdateRate:DWORD;
                                const riid: TIID;
                                out ppUnk:IUnknown):HResult;stdcall;
var
 s1:string;
 i:longint;
 aGrp:TOPCGroup;
 perDeadband:single;
begin
 result:=S_OK;
 s1:=szName;
 i:=0;

 if (s1 = '') then
  repeat
   s1:=s1 + IntToStr(GetTickCount);
   i:=succ(i);
  until (not IsNameUsedInAnyGroup(s1)) or (i > 9);

 if i > 9 then
  begin
   result:=OPC_E_DUPLICATENAME;
   phServerGroup:=0;
   Exit;
  end;

 if IsNameUsedInAnyGroup(s1) then
  begin
   result:=OPC_E_DUPLICATENAME;
   phServerGroup:=0;
   Exit;
  end;

 aGrp:=TOPCGroup.Create(self,grps);
 if aGrp = nil then
  begin
   result:=E_OUTOFMEMORY;
   phServerGroup:=0;
   Exit;
  end;

 grps.Add(aGrp);
 phServerGroup:=GetNewGroupNumber;
 if assigned(pPercentDeadband) then
  perDeadband:=pPercentDeadband^
 else
  perDeadband:=0;

 aGrp.SetUp(s1,bActive,dwRequestedUpdateRate,hClientGroup,0,
            perDeadband,dwLCID,phServerGroup);

 aGrp.ValidateTimeBias(pTimeBias);

 if dwRequestedUpdateRate <> aGrp.requestedUpdateRate then
  result:=OPC_S_UNSUPPORTEDRATE;

 pRevisedUpdateRate:=aGrp.requestedUpdateRate;

 MicroNetOPCForm.UpdateGroupCount;
 ppUnk:=aGrp;
end;

function TDA2.GetErrorString(dwError:HResult; dwLocale:TLCID; out ppString:POleStr):HResult; stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TDA2.GetGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult; stdcall;
var
 i,gNum:integer;
begin
 gNum:=IsGroupNamePresent(grps,szName);     //returns the index to the groups list
 if (addr(ppUnk) = nil) or (gNum = -1)then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 i:=gNum;
// i:=FindIndexViaGrpNumber(grps,gNum);
 if (i = -1) or (i > (grps.count-1) )  then
  begin
   Result:=E_FAIL;              Exit;
  end;
 result:=IUnknown(TOPCGroup(grps[i])).QueryInterface(riid,ppUnk);
end;

function TDA2.GetStatus(out ppServerStatus:POPCSERVERSTATUS):HResult;stdcall;
var
 aFileTime:TFileTime;
begin
 if (addr(ppServerStatus) = nil) then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 result:=S_OK;
 ppServerStatus:=POPCSERVERSTATUS(CoTaskMemAlloc(sizeof(OPCSERVERSTATUS)));
 if ppServerStatus = nil then
  begin
   ppServerStatus:=nil;         result:=E_OUTOFMEMORY;   Exit;
  end;

 DataTimeToOPCTime(srvStarted,aFileTime);
 CoFiletimeNow(ppServerStatus.ftCurrentTime);
 DataTimeToOPCTime(lastClientUpdate,ppServerStatus.ftLastUpdateTime);

 ppServerStatus.dwServerState:=OPC_STATUS_RUNNING;
 ppServerStatus.dwGroupCount:=GetGroupCount(grps) + GetGroupCount(pubGrps);
 ppServerStatus.dwBandWidth:=100;
 ppServerStatus.wMajorVersion:=1;
 ppServerStatus.wMinorVersion:=2;
 ppServerStatus.wBuildNumber:=5;
 ppServerStatus.szVendorInfo:=StringToLPOLESTR('MRD');
end;

function TDA2.RemoveGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
var
 i:integer;
 aGrp:TOPCGroup;
begin
 result:=S_OK;
 if hServerGroup < 1 then
  begin
   result:=E_INVALIDARG;
   Exit;
  end;

 i:=FindIndexViaGrpNumber(grps,hServerGroup);
 if i = -1 then            //the client as already freed the group
  Exit;                    //and we deleted it from the server list in the group destroy

 aGrp:=grps[i];
 if (aGrp.RefCount > 2) and not bForce then
  begin
   result:=OPC_S_INUSE;
   Exit;
  end;

 GroupRemovingSelf(grps,hServerGroup);
end;

function TDA2.CreateGroupEnumerator(dwScope:OPCENUMSCOPE; const riid:TIID;
                                         out ppUnk:IUnknown):HResult; stdcall;

 procedure EnumerateForStrings;
 var
  i:integer;
  pvList,pubList:TStringList;
 begin
  pvList:=nil;               pubList:=nil;
  try
   case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
     pvList:=CreateGrpNameList(grps);
    OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
     pubList:=CreateGrpNameList(pubGrps);
    OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
     begin
      pvList:=CreateGrpNameList(grps);
      pubList:=CreateGrpNameList(pubGrps);
     end
    else
     begin
      result:=E_INVALIDARG;
      Exit;
     end;
   end;

   if pvList = nil then
    pvList:=TStringList.Create;
   if pubList <> nil then
    if pubList.count > 0 then
     for i:= 0 to pubList.count-1 do
      pvList.Add(pubList[i]);

   result:=S_OK;
   ppUnk:=TOPCStringsEnumerator.Create(pvList);
  finally
   pvList.Free;
   pubList.Free;
  end;
 end;


 procedure EumerateForUnknown;
 var
  aList:TList;

 procedure AddAGroup(inList:TList);
 var
  i:integer;
  Obj:Pointer;
 begin
  for i:= 0 to inList.count -1 do
   begin
    Obj:=nil;
    IUnknown(TOPCGroup(inList[i])).QueryInterface(IUnknown,Obj);
    if Assigned(Obj) then
     aList.Add(Obj);
   end;
 end;

 begin
  aList:=nil;
  try
   aList:=TList.Create;
   case dwScope of
    OPC_ENUM_PRIVATE_CONNECTIONS,OPC_ENUM_PRIVATE:
     if Assigned(grps) then
      AddAGroup(grps);
    OPC_ENUM_PUBLIC_CONNECTIONS,OPC_ENUM_PUBLIC:
     if Assigned(pubGrps) then
      AddAGroup(pubGrps);
    OPC_ENUM_ALL_CONNECTIONS,OPC_ENUM_ALL:
     begin
     if Assigned(grps) then
      AddAGroup(grps);
     if Assigned(pubGrps) then
      AddAGroup(pubGrps);
     end
    else
     begin
      result:=E_INVALIDARG;
      Exit;
     end;
   end;

   if Assigned(aList) and (aList.count > 0) then
    ppUnk:=TS3UnknownEnumerator.Create(aList);
  finally
   aList.Free;
  end;
  result:=S_OK;
 end;

begin
 if not (IsEqualIID(riid,IEnumUnknown) or IsEqualIID(riid,IEnumString)) then
  begin
   result:=E_NOINTERFACE;
   Exit;
  end;

 if IsEqualIID(riid,IEnumString) then
  EnumerateForStrings
 else if IsEqualIID(riid,IEnumUnknown) then
  EumerateForUnknown
 else
  result:=E_FAIL;

end;


//{$INCLUDE IOPCCommonIMPL}
function TDA2.SetLocaleID(dwLcid:TLCID):HResult;stdcall;
begin
 if (dwLcid = LOCALE_SYSTEM_DEFAULT) or (dwLcid = LOCALE_USER_DEFAULT) then
  begin
   localID:=dwLcid;
   result:=S_OK;
  end
 else
  result:=E_INVALIDARG;
end;

function TDA2.GetLocaleID(out pdwLcid:TLCID):HResult;stdcall;
begin
 pdwLcid:=localID;
 result:=S_OK;
end;

function TDA2.QueryAvailableLocaleIDs(out pdwCount:UINT; out pdwLcid:PLCIDARRAY):HResult;stdcall;
begin
 pdwCount:=2;
 pdwLcid:=PLCIDARRAY(CoTaskMemAlloc(pdwCount*sizeof(LCID)));
 if (pdwLcid = nil) then
  begin
   if pdwLcid <> nil then  CoTaskMemFree(pdwLcid);
   result:=E_OUTOFMEMORY;
   Exit;
  end;
 pdwLcid[0]:=LOCALE_SYSTEM_DEFAULT;
 pdwLcid[1]:=LOCALE_USER_DEFAULT;
 result:=S_OK;
end;

function TDA2.GetErrorString(dwError:HResult; out ppString:POleStr):HResult;stdcall;
begin
 ppString:=StringToLPOLESTR(OPCErrorCodeToString(dwError));
 result:=S_OK;
end;

function TDA2.SetClientName(szName:POleStr):HResult;stdcall;
begin
 if (addr(szName) = nil) then
  begin
   Result:=E_INVALIDARG;              Exit;
  end;
 clientName:=szName;
 result:=S_OK;
end;

//{$INCLUDE IOPCServerPublicGroupsIMPL}
function TDA2.GetPublicGroupByName(szName:POleStr; const riid:TIID; out ppUnk:IUnknown):HResult;stdcall;
begin
 if IsGroupNamePresent(pubGrps,szName) <> -1 then
  begin
   result:=S_OK;
   ppUnk:=self;
  end
 else
 result:=OPC_E_NOTFOUND;
end;

function TDA2.RemovePublicGroup(hServerGroup:OPCHANDLE; bForce:BOOL):HResult;stdcall;
begin
 //do something with the forse if needed
 GroupRemovingSelf(pubGrps,hServerGroup);
 result:=S_OK;
end;

//{$INCLUDE IOPCBrowseServerAddressSpaceIMPL}
function TDA2.QueryOrganization(out pNameSpaceType:OPCNAMESPACETYPE):HResult;stdcall;
begin
 pNameSpaceType:=OPC_NS_FLAT;
 result:=S_OK;
end;

function TDA2.ChangeBrowsePosition(dwBrowseDirection:OPCBROWSEDIRECTION;
                              szString:POleStr):HResult;stdcall;
begin
 result:=E_FAIL
end;

function TDA2.BrowseOPCItemIDs(dwBrowseFilterType:OPCBROWSETYPE; szFilterCriteria:POleStr;
                          vtDataTypeFilter:TVarType; dwAccessRightsFilter:DWORD;
                          out ppIEnumString:IEnumString):HResult;stdcall;
var
 i:integer;
 tList:TStringList;
begin
//add filter support
 result:=S_OK;
 tList:=nil;
 try
  tList:=TStringList.Create;
  if tList = nil then
   begin
    result:=E_OUTOFMEMORY;
    Exit;
   end;

  for i:= low(posItems) to high(posItems) do
   tList.Add(posItems[i].tagname);

  ppIEnumString:=TOPCStringsEnumerator.Create(tList);
 finally
  tList.Free;
 end;
end;

function TDA2.GetItemID(szItemDataID:POleStr; out szItemID:POleStr):HResult;stdcall;
var
 propID:integer;
begin
 result:=S_OK;
 if length(szItemDataID) = 0 then
  szItemID:=StringToLPOLESTR(szItemDataID)
 else
  begin
   propID:=ReturnPropIDFromTagname(szItemDataID);
   if propID = 0 then
    result:=OPC_E_UNKNOWNITEMID
   else
    szItemID:=StringToLPOLESTR(szItemDataID);
  end;
end;

function TDA2.BrowseAccessPaths(szItemID:POleStr; out ppIEnumString:IEnumString):HResult;stdcall;
begin
 result:=E_NOTIMPL;
end;

//{$INCLUDE IPersistFileIMPL}
function TDA2.GetClassID(out classID:TCLSID):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.IsDirty:HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.Load(pszFileName:POleStr; dwMode:Longint):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.Save(pszFileName:POleStr; fRemember:BOOL):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.SaveCompleted(pszFileName:POleStr):HResult;stdcall;
begin
 result:=S_FALSE;
end;

function TDA2.GetCurFile(out pszFileName:POleStr):HResult;stdcall;
begin
 result:=S_FALSE;
end;


function GetNextFreeServerSpot:integer;
var
 i:integer;
begin
 result:=-1;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] = nil then
   begin
    result:=i;
    Exit;
   end;
end;

function FindServerInArray(which:TDA2):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   if theServers[i] = which then
   begin
    result:=i;
    Exit;
   end;
end;

function ReturnServerCount:integer;
var
 i:integer;
begin
 result:=0;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   result:=succ(result);
end;

procedure KillServers;
var
 i:integer;
begin
 for i:= high(theServers) downTo low(theServers) do
  begin
   CoDisconnectObject(TDA2(theServers[i]) as IUnknown,0);
   TDA2(theServers[i]).Free;
  end;
 FreeAndNil(theServers);
end;

procedure TDA2.CreateGroups;
begin
 if grps <> nil then Exit;
 grps:=TList.Create;
 grps.Capacity:=255;
 pubGrps:=TList.Create;
 pubGrps.Capacity:=grps.Capacity;
end;

procedure TDA2.Initialize;
var
 i:integer;
begin
 i:=GetNextFreeServerSpot;
 if i = -1 then Exit;
 inherited Initialize;
 srvStarted:=Now;
 lastClientUpdate:=0;
 localID:=LOCALE_SYSTEM_DEFAULT;

 FIConnectionPoints:=TConnectionPoints.Create(self);
 FIOPCItemProperties:=TOPCItemProp.Create;

 FOnSDConnect:=ShutdownOnConnect;
 FIConnectionPoints.CreateConnectionPoint(IID_IOPCShutdown,ckSingle,FOnSDConnect);

 CreateGroups;

 //hook into Main program here    may have multiple servers
 theServers[i]:=self;
 MicroNetOPCForm.UpdateGroupCount;
end;

procedure TDA2.ShutdownOnConnect(const Sink: IUnknown; Connecting: Boolean);
begin
 if connecting then
  ClientIUnknown:=Sink
 else
  ClientIUnknown:=nil
end;

destructor TDA2.Destroy;
var
 i:integer;
begin
 if grps <> nil then
  for i:= 0 to grps.count-1 do
   TOPCGroup(grps.Items[i]).Free;
 grps.Free;
 if pubGrps <> nil then
  for i:= 0 to pubGrps.count-1 do
   TOPCGroup(pubGrps.Items[i]).Free;
 pubGrps.Free;
 i:=FindServerInArray(self);
 if i <> -1 then
  theServers[i]:=nil;                //the client has let us be free ;)

 if Assigned(FIConnectionPoints) then                   FIConnectionPoints.Free;
 if Assigned(FIOPCItemProperties) then                  FIOPCItemProperties.Free;
 MicroNetOPCForm.UpdateGroupCount;
 Inherited;
end;

function TDA2.GetNewGroupNumber:longword;
const
 grpIndex:longword = 1;             //Assignable Typed Constants gota lovem
begin
 grpIndex:=succ(grpIndex);         //get us a new reference number
 result:=grpIndex;
end;

function TDA2.GetNewItemNumber:longword;
const
 itemIndex:longword = 1;
begin
 itemIndex:=succ(itemIndex);
 result:=itemIndex;
end;

function TDA2.FindIndexViaGrpNumber(wGrp:TList;gNum:longword):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to wGrp.count-1 do
  if TOPCGroup(wGrp[i]).serverHandle = gNum then
   begin
    result:=i;
    Break;
   end;
end;

procedure TDA2.GroupRemovingSelf(wGrp:TList; gNum:integer);
var
 i:integer;
begin
 i:=FindIndexViaGrpNumber(wGrp,gNum);
 if (i <> -1) then
  wGrp.Delete(i);
 MicroNetOPCForm.UpdateGroupCount;
end;

function TDA2.GetGroupCount(gList:TList):integer;
begin
 result:=0;
 if gList = nil then Exit;
 result:=gList.count;
end;

function TDA2.CreateGrpNameList(gList:TList):TStringList;
var
 i:integer;
begin
 result:=nil;
 if gList = nil then Exit;
 result:=TStringList.Create;
 for i:= 0 to gList.count-1 do
  result.Add(TOPCGroup(gList.Items[i]).tagName);
 if result.count = 0 then
  begin
   result.Free;
   result:=nil;
  end;
end;

function TDA2.IsGroupNamePresent(gList:TList; theName:string):integer;
var
 i:integer;
begin
 result:=-1;
 for i:= 0 to gList.count-1 do
  if theName = TOPCGroup(gList.Items[i]).tagName then
    begin
     result:=i;            Break;
    end;
end;

function TDA2.IsNameUsedInAnyGroup(theName:string):boolean;
var
 i:integer;
begin
 result:=false;
 i:=IsGroupNamePresent(grps,theName);
 if i <> -1 then
  begin
   result:=true;                     Exit;
  end;
 i:=IsGroupNamePresent(pubGrps,theName);
 if i <> -1 then
  result:=true;
end;

function TDA2.IsThisGroupPublic(aList:TList):boolean;
begin
 result:=boolean((pubGrps <> nil)    and
                 (pubGrps.count = 0) and
                 (aList <> nil)      and
                 (pubGrps = aList));
end;

procedure TDA2.TimeSlice(cTime:TDateTime);
var
 i:integer;
begin
 lastClientUpdate:=cTime;
 if Assigned(grps) then
  for i:= 0 to grps.count-1 do
   TOPCGroup(grps.Items[i]).TimeSlice(cTime);
 if Assigned(pubGrps) then
  for i:= 0 to pubGrps.count-1 do
   TOPCGroup(pubGrps.Items[i]).TimeSlice(cTime);
end;

function TDA2.CloneAGroup(szName:string;aGrp:TTypedComObject; out res:HResult):IUnknown;
var
 sGrp,dGrp:TOPCGroup;
begin
 sGrp:=TOPCGroup(aGrp);
 dGrp:=TOPCGroup.Create(self,grps);
 if dGrp = nil then
  begin
   result:=nil;
   res:=E_OUTOFMEMORY;      Exit;
  end;

 grps.Add(dGrp);
 dGrp.tagName:=szName;
 sGrp.CloneYourSelf(dGrp);
 result:=dGrp;
 res:=S_OK;
end;

initialization
  TAutoObjectFactory.Create(ComServer, TDA2, Class_OPCServer,
                            ciMultiInstance, tmApartment);

 //if an OPC client(s) is connected and the user has selected to quit after
 //the warning in the FormCloseQuery then do not let the system ask again in the
 //AutomationTerminateProc procedure in the VCL.
 ComServer.UIInteractive:=false;
finalization
 //if an OPC client is connected and this is a forced kill then if CoUnintialize
 //is not called here the OLE dll will generate an error when it is called after
 //we have killed the servers.
 CoUninitialize;
 KillServers;

end.
