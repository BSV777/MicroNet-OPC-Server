unit Globals;

interface

uses
  Windows,Messages,SysUtils,Classes,Graphics,StdCtrls,Forms,Dialogs,Controls,
  ShellAPI,ActiveX,OPCDA;

type
 itemIDStrings = record
  trunk,branch,leaf:string[255];
 end;

type
  itemProps = record
    PropID: longword;
    tagname: string[64];
    dataType:integer;
  end;

const
 IID_IUnknown: TIID = '{00000000-0000-0000-C000-000000000046}';
 io2Read    = 1;
 io2Write   = 2;
 io2Refresh = 3;
 io2Change  = 4;

var
 posItems: array[0..310] of itemProps;
 itemValues:array[0..310] of word;

function ScanToChar(const theString:string; var start:integer;theChar:char):string;
function ReturnPropIDFromTagname(const s1:string):longword;
function ReturnTagnameFromPropID(PropID:longword):string;
function CanPropIDBeWritten(i:longword):boolean;
function ReturnDataTypeFromPropID(i:longword):integer;
procedure DataTimeToOPCTime(cTime:TDateTime; var OPCTime:TFileTime);
function ConvertVariant(cv:variant; reqDataType:TVarType):variant;

implementation

function ScanToChar(const theString:string; var start:integer;theChar:char):string;
var
 tempS:string;
 finish:boolean;
 nextloc,strLength: integer;
begin
 {$R-}
 strLength := length(theString);
 finish := false;
 SetLength(tempS,strLength);
 result := tempS;
 nextloc := 1;
 while not finish do
  begin
   if (start < 256) and (theString[start] <> theChar) and
      (theString[start] <> chr(13)) and (start <= strLength) then
    begin
     tempS[nextloc] := theString[start];
     nextloc := succ(nextloc);
     start := succ(start);
    end
   else
    begin
     SetLength(tempS,nextloc-1);      {this sets the length of the string}
     finish:=true;                    {exit the loop}
     result:=tempS;                   {return the value}
    end;
  end;
 {$R+}
end;

function ReturnPropIDFromTagname(const s1:string):longword;
var
 i:integer;
begin
 result:=0;
 for i:= low(posItems) to high(posItems) do
  if posItems[i].tagname = s1 then
   begin
    result:=posItems[i].PropID;
    Exit;
   end;
end;

function ReturnTagnameFromPropID(PropID:longword):string;
var
 i:integer;
begin
 result:='';
 for i:= low(posItems) to high(posItems) do
  if posItems[i].PropID = PropID then
   begin
    result:=posItems[i].tagname;
    Exit;
   end;
end;

function CanPropIDBeWritten(i:longword):boolean;
begin
 i:= i - posItems[low(posItems)].PropID;
 result := boolean(i in [94..125]) or boolean(i in [1..32]) or boolean(i in [186..217]);  //Права на запись тегов
end;

function ReturnDataTypeFromPropID(i:longword):integer;
var
 x:longword;
begin
 x:= i - posItems[low(posItems)].PropID;
 if (x <= high(posItems)) then
  result:=posItems[x].dataType
 else
  result:=VT_UI2;
end;

procedure DataTimeToOPCTime(cTime:TDateTime; var OPCTime:TFileTime);
var
 sTime:TSystemTime;
begin
 DateTimeToSystemTime(cTime,sTime);
 SystemTimeToFileTime(sTime,OPCTime);
 LocalFileTimeToFileTime(OPCTime,OPCTime);
end;

function ConvertVariant(cv:variant; reqDataType:TVarType):variant;
begin
 try
  result:=VarAsType(cv,reqDataType)
 except
  on EVariantError do   result:=DISP_E_TYPEMISMATCH;
 end;
end;

end.
