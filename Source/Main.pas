unit Main;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls,
     AxCtrls, ExtCtrls, MicroNetOPC, RegDeRegServer,ShellAPI, Menus;

const
 serverName = 'MicroNet.OPCServer';
 WM_ICONNOTIFY = WM_USER + 1234;

type
  TMicroNetOPCForm = class(TForm)
    Label2: TLabel;
    ClientConLbl: TLabel;
    Label1: TLabel;
    GrpCountLbl: TLabel;
    PopupMenu1: TPopupMenu;
    mnOpen: TMenuItem;
    mnBreak: TMenuItem;
    mnExit: TMenuItem;
    Image1: TImage;
    Label3: TLabel;
    lbCopyright: TLabel;
    lbTrademark: TLabel;
    WatchDogTimer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure mnOpenClick(Sender: TObject);
    procedure mnExitClick(Sender: TObject);
    procedure WMIconNotify(var Message:TMessage); message WM_ICONNOTIFY;
    procedure WMClose(var Message:TMessage); message WM_CLOSE;
    procedure lbCopyrightClick(Sender: TObject);
    procedure lbTrademarkClick(Sender: TObject);
    procedure WatchDogTimerTimer(Sender: TObject);
  private
    FHI:Ticon;
    FNID:TNotifyIconData;
  public
   constructor Create(AOwner:TComponent); override;
   destructor Destroy; override;
   function ReturnItemValue(ii:integer):variant;
   procedure UpdateGroupCount;
  end;

var
  MicroNetOPCForm : TMicroNetOPCForm;
  MicroNetOPCServer : TMicroNetOPC;
  clientsConnected : integer;

implementation

{$R *.RES}

{$R *.DFM}

uses ServImpl,OPCCOMN,Globals,ActiveX;

function TMicroNetOPCForm.ReturnItemValue(ii:integer):variant;
begin
  result:=itemValues[ii];
end;

procedure TMicroNetOPCForm.FormCreate(Sender: TObject);
var
 s1:string;
begin
 if ParamCount <> 0 then
  begin
   s1:=LowerCase(ParamStr(1));
   if s1 = 'regserver' then
    begin
     RegisterTheServer(serverName);
     PostMessage(self.handle,WM_CLOSE,0,0);
     Exit;
    end
   else if s1 = 'unregserver' then
    begin
     UnRegisterTheServer(serverName);
     PostMessage(self.handle,WM_CLOSE,0,0);
     Exit;
    end;
  end;
  lbCopyright.ShowHint := True;
  lbTrademark.ShowHint := True;  
  ServerPath := ExtractFilePath(Application.EXEName);
  MicroNetOPCServer := TMicroNetOPC.Create(Self);
  UpdateGroupCount;
end;

procedure TMicroNetOPCForm.UpdateGroupCount;
var
 i,g:integer;
begin
 if Application.Terminated then Exit;
 clientsConnected:=0;
 g:=0;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   begin
    clientsConnected:=succ(clientsConnected);
    if Assigned(theServers[i].grps) then
     g:=g + theServers[i].grps.count;
    if Assigned(theServers[i].pubGrps) then
     g:=g + theServers[i].pubGrps.count;
   end;
 ClientConLbl.caption:=IntToStr(clientsConnected);
 GrpCountLbl.caption:=IntToStr(g);
end;

procedure TMicroNetOPCForm.mnOpenClick(Sender: TObject);
var
   PT : TPoint;
begin
  GetCursorPos(PT);
  if PT.X > Screen.DesktopWidth / 2 then Left := PT.X - 272 else Left := PT.X;
  if PT.Y > Screen.DesktopHeight / 2 then Top := PT.Y - 177 else Top := PT.Y;
  Show;
  BringWindowToTop(Handle);
end;

procedure TMicroNetOPCForm.mnExitClick(Sender: TObject);
const
  OneSecond = 1 / (24 * 60 * 60);
var
  i:integer;
  Obj: Pointer;
  StartTime: TDateTime;
begin
  if clientsConnected = 0 then Application.Terminate else
    if (MessageDlg('Прекращение работы сервера повлияет на подключенных к нему клиентов.' +
      #13 + #10 + 'Остановить MicroNet OPC-сервер?', mtConfirmation,[mbYes,mbNo],0) =  mrYes) then
      begin
        for i:= low(theServers) to high(theServers) do
          begin
            if theServers[i] <> nil then
              if theServers[i].ClientIUnknown <> nil then
                if Succeeded(theServers[i].ClientIUnknown.QueryInterface(IOPCShutdown,Obj)) then
                  IOPCShutdown(Obj).ShutdownRequest('MicroNet IOPCShutdown request.');
          end;
        StartTime := Now;
        while (Now - StartTime) < OneSecond do Application.ProcessMessages;
        Close;
      end;
end;

procedure TMicroNetOPCForm.WMIconNotify(var Message:TMessage);
var
   PT:TPoint;
begin
   if Message.lParam=WM_LBUTTONDOWN then mnOpenClick(Self)
    else if Message.lParam=WM_RBUTTONDOWN then
      begin
        GetCursorPos(PT);
        PopupMenu1.Popup(PT.X,PT.Y);
      end;
end;

constructor TMicroNetOPCForm.Create(AOwner:TComponent);
begin
   inherited Create(AOwner);
   FHI:=TIcon.Create;
   FHI.Handle:=LoadIcon(HInstance,'TRAYICON');
   FNID.cbSize:=sizeof(FNID);
   FNID.Wnd:=Handle;
   FNID.uID:=1;
   FNID.uCallbackMessage:=WM_ICONNOTIFY;
   FNID.HIcon:=FHI.Handle;
   FNID.szTip:='MicroNet OPC-Server';
   FNID.uFlags:=nif_Message or nif_Icon or nif_Tip;
   Shell_NotifyIcon(NIM_ADD,@FNID);
end;

destructor TMicroNetOPCForm.Destroy;
begin
   FNID.uFlags:=0;
   Shell_NotifyIcon(NIM_DELETE,@FNID);
   FHI.Free;
   inherited Destroy;
end;

procedure TMicroNetOPCForm.WMClose(var Message:TMessage);
begin
   Message.Result:=0;
   Hide;
end;

procedure TMicroNetOPCForm.lbCopyrightClick(Sender: TObject);
var
  TempString : array[0..79] of char;
begin
  StrPCopy(TempString, 'mailto:admin@company.ru');
  ShellExecute(0, Nil, TempString, Nil, Nil, SW_NORMAL);
end;

procedure TMicroNetOPCForm.lbTrademarkClick(Sender: TObject);
var
  TempString : array[0..79] of char;
begin
  StrPCopy(TempString, 'http://www.opcfoundation.org');
  ShellExecute(0, Nil, TempString, Nil, Nil, SW_NORMAL);
end;

procedure TMicroNetOPCForm.WatchDogTimerTimer(Sender: TObject);
begin
  if Alive = True then Alive := False else Close;
end;

end.
unit Main;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, StdCtrls,
     AxCtrls, ExtCtrls, MicroNetOPC, RegDeRegServer,ShellAPI, Menus;

const
 serverName = 'MicroNet.OPCServer';
 WM_ICONNOTIFY = WM_USER + 1234;

type
  TMicroNetOPCForm = class(TForm)
    Label2: TLabel;
    ClientConLbl: TLabel;
    Label1: TLabel;
    GrpCountLbl: TLabel;
    PopupMenu1: TPopupMenu;
    mnOpen: TMenuItem;
    mnBreak: TMenuItem;
    mnExit: TMenuItem;
    Image1: TImage;
    Label3: TLabel;
    lbCopyright: TLabel;
    lbTrademark: TLabel;
    WatchDogTimer: TTimer;
    procedure FormCreate(Sender: TObject);
    procedure mnOpenClick(Sender: TObject);
    procedure mnExitClick(Sender: TObject);
    procedure WMIconNotify(var Message:TMessage); message WM_ICONNOTIFY;
    procedure WMClose(var Message:TMessage); message WM_CLOSE;
    procedure lbCopyrightClick(Sender: TObject);
    procedure lbTrademarkClick(Sender: TObject);
    procedure WatchDogTimerTimer(Sender: TObject);
  private
    FHI:Ticon;
    FNID:TNotifyIconData;
  public
   constructor Create(AOwner:TComponent); override;
   destructor Destroy; override;
   function ReturnItemValue(ii:integer):variant;
   procedure UpdateGroupCount;
  end;

var
  MicroNetOPCForm : TMicroNetOPCForm;
  MicroNetOPCServer : TMicroNetOPC;
  clientsConnected : integer;

implementation

{$R *.RES}

{$R *.DFM}

uses ServImpl,OPCCOMN,Globals,ActiveX;

function TMicroNetOPCForm.ReturnItemValue(ii:integer):variant;
begin
  result:=itemValues[ii];
end;

procedure TMicroNetOPCForm.FormCreate(Sender: TObject);
var
 s1:string;
begin
 if ParamCount <> 0 then
  begin
   s1:=LowerCase(ParamStr(1));
   if s1 = 'regserver' then
    begin
     RegisterTheServer(serverName);
     PostMessage(self.handle,WM_CLOSE,0,0);
     Exit;
    end
   else if s1 = 'unregserver' then
    begin
     UnRegisterTheServer(serverName);
     PostMessage(self.handle,WM_CLOSE,0,0);
     Exit;
    end;
  end;
  lbCopyright.ShowHint := True;
  lbTrademark.ShowHint := True;  
  ServerPath := ExtractFilePath(Application.EXEName);
  MicroNetOPCServer := TMicroNetOPC.Create(Self);
  UpdateGroupCount;
end;

procedure TMicroNetOPCForm.UpdateGroupCount;
var
 i,g:integer;
begin
 if Application.Terminated then Exit;
 clientsConnected:=0;
 g:=0;
 for i:= low(theServers) to high(theServers) do
  if theServers[i] <> nil then
   begin
    clientsConnected:=succ(clientsConnected);
    if Assigned(theServers[i].grps) then
     g:=g + theServers[i].grps.count;
    if Assigned(theServers[i].pubGrps) then
     g:=g + theServers[i].pubGrps.count;
   end;
 ClientConLbl.caption:=IntToStr(clientsConnected);
 GrpCountLbl.caption:=IntToStr(g);
end;

procedure TMicroNetOPCForm.mnOpenClick(Sender: TObject);
var
   PT : TPoint;
begin
  GetCursorPos(PT);
  if PT.X > Screen.DesktopWidth / 2 then Left := PT.X - 272 else Left := PT.X;
  if PT.Y > Screen.DesktopHeight / 2 then Top := PT.Y - 177 else Top := PT.Y;
  Show;
  BringWindowToTop(Handle);
end;

procedure TMicroNetOPCForm.mnExitClick(Sender: TObject);
const
  OneSecond = 1 / (24 * 60 * 60);
var
  i:integer;
  Obj: Pointer;
  StartTime: TDateTime;
begin
  if clientsConnected = 0 then Application.Terminate else
    if (MessageDlg('Прекращение работы сервера повлияет на подключенных к нему клиентов.' +
      #13 + #10 + 'Остановить MicroNet OPC-сервер?', mtConfirmation,[mbYes,mbNo],0) =  mrYes) then
      begin
        for i:= low(theServers) to high(theServers) do
          begin
            if theServers[i] <> nil then
              if theServers[i].ClientIUnknown <> nil then
                if Succeeded(theServers[i].ClientIUnknown.QueryInterface(IOPCShutdown,Obj)) then
                  IOPCShutdown(Obj).ShutdownRequest('MicroNet IOPCShutdown request.');
          end;
        StartTime := Now;
        while (Now - StartTime) < OneSecond do Application.ProcessMessages;
        Close;
      end;
end;

procedure TMicroNetOPCForm.WMIconNotify(var Message:TMessage);
var
   PT:TPoint;
begin
   if Message.lParam=WM_LBUTTONDOWN then mnOpenClick(Self)
    else if Message.lParam=WM_RBUTTONDOWN then
      begin
        GetCursorPos(PT);
        PopupMenu1.Popup(PT.X,PT.Y);
      end;
end;

constructor TMicroNetOPCForm.Create(AOwner:TComponent);
begin
   inherited Create(AOwner);
   FHI:=TIcon.Create;
   FHI.Handle:=LoadIcon(HInstance,'TRAYICON');
   FNID.cbSize:=sizeof(FNID);
   FNID.Wnd:=Handle;
   FNID.uID:=1;
   FNID.uCallbackMessage:=WM_ICONNOTIFY;
   FNID.HIcon:=FHI.Handle;
   FNID.szTip:='MicroNet OPC-Server';
   FNID.uFlags:=nif_Message or nif_Icon or nif_Tip;
   Shell_NotifyIcon(NIM_ADD,@FNID);
end;

destructor TMicroNetOPCForm.Destroy;
begin
   FNID.uFlags:=0;
   Shell_NotifyIcon(NIM_DELETE,@FNID);
   FHI.Free;
   inherited Destroy;
end;

procedure TMicroNetOPCForm.WMClose(var Message:TMessage);
begin
   Message.Result:=0;
   Hide;
end;

procedure TMicroNetOPCForm.lbCopyrightClick(Sender: TObject);
var
  TempString : array[0..79] of char;
begin
  StrPCopy(TempString, 'mailto:admin@company.ru');
  ShellExecute(0, Nil, TempString, Nil, Nil, SW_NORMAL);
end;

procedure TMicroNetOPCForm.lbTrademarkClick(Sender: TObject);
var
  TempString : array[0..79] of char;
begin
  StrPCopy(TempString, 'http://www.opcfoundation.org');
  ShellExecute(0, Nil, TempString, Nil, Nil, SW_NORMAL);
end;

procedure TMicroNetOPCForm.WatchDogTimerTimer(Sender: TObject);
begin
  if Alive = True then Alive := False else Close;
end;

end.
