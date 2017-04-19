unit OPCErrorStrings;

interface

uses Windows,SysUtils;

function OPCErrorCodeToString(x:HResult):string;

const
  OPC_E_INVALIDHANDLE = HResult($C0040001);
  OPC_E_BADTYPE = HResult($C0040004);
  OPC_E_PUBLIC = HResult($C0040005);
  OPC_E_BADRIGHTS = HResult($C0040006);
  OPC_E_UNKNOWNITEMID = HResult($C0040007);
  OPC_E_INVALIDITEMID = HResult($C0040008);
  OPC_E_INVALIDFILTER = HResult($C0040009);
  OPC_E_UNKNOWNPATH = HResult($C004000A);
  OPC_E_RANGE = HResult($C004000B);
  OPC_E_DUPLICATENAME = HResult($C004000C);
  OPC_S_UNSUPPORTEDRATE = HResult($0004000D);
  OPC_S_CLAMP = HResult($0004000E);
  OPC_S_INUSE = HResult($0004000F);
  OPC_E_INVALIDCONFIGFILE = HResult($C0040010);
  OPC_E_NOTFOUND = HResult($C0040011);
  OPC_E_INVALID_PID = HResult($C0040203);
  OPC_S_ALREADYACKED = HResult($00040200);
  OPC_S_INVALIDBUFFERTIME = HResult($00040201);
  OPC_S_INVALIDMAXSIZE = HResult($00040202);
  OPC_E_INVALIDBRANCHNAME = HResult($C0040203);
  OPC_E_INVALIDTIME = HResult($C0040204);
  OPC_E_BUSY = HResult($C0040205);
  OPC_E_NOINFO = HResult($C0040206);
  OPC_E_MAXEXCEEDED = HResult($C0041001);
  OPC_S_NODATA = HResult($40041002);
  OPC_S_MOREDATA = HResult($40041003);
  OPC_E_INVALIDAGGREGATE = HResult($C0041004);
  OPC_S_CURRENTVALUE = HResult($40041005);
  OPC_S_EXTRADATA = HResult($40041006);
  OPC_W_NOFILTER = HResult($80041007);
  OPC_E_UNKNOWNATTRID = HResult($C0041008);
  OPC_E_NOT_AVAIL = HResult($C0041009);
  OPC_E_INVALIDDATATYPE = HResult($C004100A);
  OPC_E_DATAEXISTS = HResult($C004100B);
  OPC_E_INVALIDATTRID = HResult($C004100C);
  OPC_E_NODATAEXISTS = HResult($C004100D);
  OPC_S_INSERTED = HResult($4004100E);
  OPC_S_REPLACED = HResult($4004100F);
  OPC_E_PRIVATE_ACTIVE = HResult($C0040301);
  OPC_E_LOW_IMPERS_LEVEL = HResult($C0040302);
  OPC_S_LOW_AUTHN_LEVEL = HResult($00040303);

implementation

function OPCErrorCodeToString(x:HResult):string;
var
 len:integer;
 buffer:array[0..255] of char;
begin
 result:='';
 case x of
  OPC_E_INVALIDHANDLE:
   result:='The value of the handle is invalid.';
  OPC_E_BADTYPE:
   result:='The server cannot convert the data between the requested data type and the canonical data type.';
  OPC_E_PUBLIC:
   result:='The requested operation cannot be done on a public group.';
  OPC_E_BADRIGHTS:
   result:='The Items AccessRights do not allow the operation.';
  OPC_E_UNKNOWNITEMID:
   result:='The item is no longer available in the server address space.';
  OPC_E_INVALIDITEMID:
   result:='The item definition does not conform to the server''s syntax.';
  OPC_E_INVALIDFILTER:
   result:='The filter string was not valid.';
  OPC_E_UNKNOWNPATH:
   result:='The item''s access path is not known to the server.';
  OPC_E_RANGE:
   result:='The value was out of range.';
  OPC_E_DUPLICATENAME:
   result:='Duplicate name not allowed.';
  OPC_S_UNSUPPORTEDRATE:
   result:='The server does not support the requested data rate but will use the closest available rate.';
  OPC_S_CLAMP:
   result:='A value passed to WRITE was accepted but the output was clamped.';
  OPC_S_INUSE:
   result:='The operation cannot be completed because the object still has references that exist.';
  OPC_E_INVALIDCONFIGFILE:
   result:='The server''s configuration file is an invalid format.';
  OPC_E_NOTFOUND:
   result:='The server could not locate the requested object.';
  OPC_E_INVALID_PID:
   result:='The server does not recognise the passed property ID.';
  OPC_S_ALREADYACKED:
   result:='The condition has already been acknowleged.';
  OPC_S_INVALIDBUFFERTIME:
   result:='The buffer time parameter was invalid.';
  OPC_S_INVALIDMAXSIZE:
   result:='The max size parameter was invalid.';
//  OPC_E_INVALIDBRANCHNAME:
//   result:='The string was not recognized as an area name.';
  OPC_E_INVALIDTIME:
   result:='The time does not match the latest active time.';
  OPC_E_BUSY:
   result:='A refresh is currently in progress.';
  OPC_E_NOINFO:
   result:='Information is not available.';
 end;

 if length(result) <> 0 then
  Exit;

 len:=FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_ARGUMENT_ARRAY,
                       nil,longword(x),0, Buffer,sizeOf(buffer), nil);
 while (Len > 0) and (Buffer[Len - 1] in [#0..#32, '.']) do
  Dec(Len);
 SetString(result,buffer,len);
end;

end.
unit OPCErrorStrings;

interface

uses Windows,SysUtils;

function OPCErrorCodeToString(x:HResult):string;

const
  OPC_E_INVALIDHANDLE = HResult($C0040001);
  OPC_E_BADTYPE = HResult($C0040004);
  OPC_E_PUBLIC = HResult($C0040005);
  OPC_E_BADRIGHTS = HResult($C0040006);
  OPC_E_UNKNOWNITEMID = HResult($C0040007);
  OPC_E_INVALIDITEMID = HResult($C0040008);
  OPC_E_INVALIDFILTER = HResult($C0040009);
  OPC_E_UNKNOWNPATH = HResult($C004000A);
  OPC_E_RANGE = HResult($C004000B);
  OPC_E_DUPLICATENAME = HResult($C004000C);
  OPC_S_UNSUPPORTEDRATE = HResult($0004000D);
  OPC_S_CLAMP = HResult($0004000E);
  OPC_S_INUSE = HResult($0004000F);
  OPC_E_INVALIDCONFIGFILE = HResult($C0040010);
  OPC_E_NOTFOUND = HResult($C0040011);
  OPC_E_INVALID_PID = HResult($C0040203);
  OPC_S_ALREADYACKED = HResult($00040200);
  OPC_S_INVALIDBUFFERTIME = HResult($00040201);
  OPC_S_INVALIDMAXSIZE = HResult($00040202);
  OPC_E_INVALIDBRANCHNAME = HResult($C0040203);
  OPC_E_INVALIDTIME = HResult($C0040204);
  OPC_E_BUSY = HResult($C0040205);
  OPC_E_NOINFO = HResult($C0040206);
  OPC_E_MAXEXCEEDED = HResult($C0041001);
  OPC_S_NODATA = HResult($40041002);
  OPC_S_MOREDATA = HResult($40041003);
  OPC_E_INVALIDAGGREGATE = HResult($C0041004);
  OPC_S_CURRENTVALUE = HResult($40041005);
  OPC_S_EXTRADATA = HResult($40041006);
  OPC_W_NOFILTER = HResult($80041007);
  OPC_E_UNKNOWNATTRID = HResult($C0041008);
  OPC_E_NOT_AVAIL = HResult($C0041009);
  OPC_E_INVALIDDATATYPE = HResult($C004100A);
  OPC_E_DATAEXISTS = HResult($C004100B);
  OPC_E_INVALIDATTRID = HResult($C004100C);
  OPC_E_NODATAEXISTS = HResult($C004100D);
  OPC_S_INSERTED = HResult($4004100E);
  OPC_S_REPLACED = HResult($4004100F);
  OPC_E_PRIVATE_ACTIVE = HResult($C0040301);
  OPC_E_LOW_IMPERS_LEVEL = HResult($C0040302);
  OPC_S_LOW_AUTHN_LEVEL = HResult($00040303);

implementation

function OPCErrorCodeToString(x:HResult):string;
var
 len:integer;
 buffer:array[0..255] of char;
begin
 result:='';
 case x of
  OPC_E_INVALIDHANDLE:
   result:='The value of the handle is invalid.';
  OPC_E_BADTYPE:
   result:='The server cannot convert the data between the requested data type and the canonical data type.';
  OPC_E_PUBLIC:
   result:='The requested operation cannot be done on a public group.';
  OPC_E_BADRIGHTS:
   result:='The Items AccessRights do not allow the operation.';
  OPC_E_UNKNOWNITEMID:
   result:='The item is no longer available in the server address space.';
  OPC_E_INVALIDITEMID:
   result:='The item definition does not conform to the server''s syntax.';
  OPC_E_INVALIDFILTER:
   result:='The filter string was not valid.';
  OPC_E_UNKNOWNPATH:
   result:='The item''s access path is not known to the server.';
  OPC_E_RANGE:
   result:='The value was out of range.';
  OPC_E_DUPLICATENAME:
   result:='Duplicate name not allowed.';
  OPC_S_UNSUPPORTEDRATE:
   result:='The server does not support the requested data rate but will use the closest available rate.';
  OPC_S_CLAMP:
   result:='A value passed to WRITE was accepted but the output was clamped.';
  OPC_S_INUSE:
   result:='The operation cannot be completed because the object still has references that exist.';
  OPC_E_INVALIDCONFIGFILE:
   result:='The server''s configuration file is an invalid format.';
  OPC_E_NOTFOUND:
   result:='The server could not locate the requested object.';
  OPC_E_INVALID_PID:
   result:='The server does not recognise the passed property ID.';
  OPC_S_ALREADYACKED:
   result:='The condition has already been acknowleged.';
  OPC_S_INVALIDBUFFERTIME:
   result:='The buffer time parameter was invalid.';
  OPC_S_INVALIDMAXSIZE:
   result:='The max size parameter was invalid.';
//  OPC_E_INVALIDBRANCHNAME:
//   result:='The string was not recognized as an area name.';
  OPC_E_INVALIDTIME:
   result:='The time does not match the latest active time.';
  OPC_E_BUSY:
   result:='A refresh is currently in progress.';
  OPC_E_NOINFO:
   result:='Information is not available.';
 end;

 if length(result) <> 0 then
  Exit;

 len:=FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM or FORMAT_MESSAGE_ARGUMENT_ARRAY,
                       nil,longword(x),0, Buffer,sizeOf(buffer), nil);
 while (Len > 0) and (Buffer[Len - 1] in [#0..#32, '.']) do
  Dec(Len);
 SetString(result,buffer,len);
end;

end.
