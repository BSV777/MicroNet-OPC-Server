unit ItemsUnit;

interface

uses Windows,ActiveX,ComObj,MicroNet_TLB,SysUtils,Dialogs,Classes,ServIMPL,OPCDA,
     axctrls,Globals,ItemAttributesOPC;

type
  TOPCItem = class
  public
   servObj:TDA2;                         //the owner
   quality:word;
   strID:string;
   pBlob:PByteArray;
   bActive,isWriteAble:longbool;
   currentValue,oldValue:variant;
   serverItemNum,clientNum,itemIndex:longword;
   vtReqDataType,canonicalDataType:TVarType;
   constructor Create;
   destructor Destroy;override;
   procedure CopyYourSelf(dItem:TOPCItem);
   procedure SetActiveState(state:longbool);
   function GetActiveState:longbool;
   function GetClientHandle:longword;
   function GetQuality:word;
   procedure SetClientHandle(h:longword);
   procedure ResolveQuality(source:word);
   function ReturnCurrentValue(source:integer):variant;
   procedure ReadItemValueStateTime(source:word; var pStateRec:OPCITEMSTATE);
   procedure WriteItemValue(toWhat:longword);
   procedure FillInOPCItemObject(aObj:TOPCItemAttributes);
   procedure SetReqDataType(aType:TVarType);
   procedure CallBackRead(var cHandle:longword; var cValue:OleVariant; var q:word);virtual;
   function UpdateYourSelf:boolean;
   procedure SetOldValue;
  end;

implementation

uses Main;

constructor TOPCItem.Create;
begin
 Inherited;
 quality:=OPC_QUALITY_GOOD;
 isWriteAble:=false;
end;

destructor TOPCItem.Destroy;
begin
 Inherited;
end;

procedure TOPCItem.CopyYourSelf(dItem:TOPCItem);
begin
 dItem.servObj:=servObj;
 dItem.quality:=quality;
 dItem.strID:=strID;
 dItem.pBlob:=pBlob;
 dItem.bActive:= bActive;
 dItem.isWriteAble:=isWriteAble;
 dItem.currentValue:=currentValue;
 dItem.oldValue:=oldValue;
 dItem.serverItemNum:=serverItemNum;
 dItem.clientNum:=clientNum;
 dItem.itemIndex:=itemIndex;
 dItem.vtReqDataType:=vtReqDataType;
 dItem.canonicalDataType:=canonicalDataType;
end;

procedure TOPCItem.SetActiveState(state:longbool);
begin
 bActive:=state;
 if not bActive then
  quality:=OPC_QUALITY_OUT_OF_SERVICE
 else
  quality:=OPC_QUALITY_GOOD;
end;

function TOPCItem.GetActiveState:longbool;
begin
 result:=bActive;
end;

function TOPCItem.GetQuality:word;
begin
 result:=quality;
end;

function TOPCItem.GetClientHandle:longword;
begin
 result:=clientNum;
end;

procedure TOPCItem.SetClientHandle(h:longword);
begin
 clientNum:=h;
end;

procedure TOPCItem.ResolveQuality(source:word);
begin
 if source = OPC_DS_CACHE then
  begin
   if bActive then                //in service so is it good
    quality:=OPC_QUALITY_GOOD
   else
    quality:=OPC_QUALITY_OUT_OF_SERVICE;
  end
 else                             //device
  quality:=OPC_QUALITY_GOOD;
end;

function TOPCItem.ReturnCurrentValue(source:integer):variant;
begin
 if vtReqDataType <> canonicalDataType then
  result:=ConvertVariant(currentValue,vtReqDataType)
 else
  result:=currentValue;
end;

procedure TOPCItem.ReadItemValueStateTime(source:word;
                                          var pStateRec:OPCITEMSTATE);
begin
 ResolveQuality(source);
 pStateRec.wQuality:=quality;
 pStateRec.hClient:=clientNum;
 pStateRec.vDataValue:=ReturnCurrentValue(source);
 DataTimeToOPCTime(servObj.lastClientUpdate,pStateRec.ftTimeStamp);
end;

procedure TOPCItem.WriteItemValue(toWhat:longword);
begin
  if (toWhat = 0) and (itemIndex in [94..125]) then MicroNetOPCServer.AsyncCmd[itemIndex - 93] := $4B52;  //������� ���������� ������� �����
  if (toWhat = 0) and (itemIndex in [1..32]) then MicroNetOPCServer.Calibr[itemIndex] := $01;             //������ ����������
  if (toWhat <> 0) and (itemIndex in [187..217]) then MicroNetOPCServer.AsyncCmd[itemIndex - 186] := $4B54; //������� �����
  if (toWhat = 0) and (itemIndex in [187..217]) then MicroNetOPCServer.AsyncCmd[itemIndex - 186] := $4B58;//������� ����
end;

procedure TOPCItem.FillInOPCItemObject(aObj:TOPCItemAttributes);
begin
 if aObj = nil then Exit;
 aObj.szAccessPath:='';
 aObj.szItemID:=strID;
 aObj.bActive:=bActive;
 aObj.hClient:=clientNum;
 aObj.hServer:=serverItemNum;
 if isWriteAble then
  aObj.dwAccessRights:=OPC_READABLE or OPC_WRITEABLE
 else
  aObj.dwAccessRights:=OPC_READABLE;

 aObj.vtRequestedDataType:=vtReqDataType;
 aObj.vtCanonicalDataType:=canonicalDataType;
 aObj.dwEUType:=OPC_NOENUM;
 aObj.vEUInfo:=VT_EMPTY;
end;

procedure TOPCItem.SetReqDataType(aType:TVarType);
begin
 if aType = VT_EMPTY then
  vtReqDataType:=canonicalDataType
 else
  vtReqDataType:=aType;
end;

procedure TOPCItem.CallBackRead(var cHandle:longword; var cValue:OleVariant; var q:word);
begin
 cHandle:=GetClientHandle;
 q:=quality;
 cValue:=ReturnCurrentValue(0);
end;

function TOPCItem.UpdateYourSelf:boolean;
begin
 currentValue:=MicroNetOPCForm.ReturnItemValue(itemIndex);
 result:=(currentValue <> oldValue);
 oldValue:=currentValue;
end;

procedure TOPCItem.SetOldValue;
begin
 if canonicalDataType = VT_BSTR then
  oldValue:=''
 else
  oldValue:=-1;
end;

end.
