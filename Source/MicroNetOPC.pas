unit MicroNetOPC;

interface

uses SysUtils, Classes, Controls, StdCtrls, AxCtrls, ExtCtrls,
    CommPort, Globals, ActiveX, ServIMPL, IniFiles, Forms;

type
  TState = (stOk, stSenErr, stNetErr, stNotMnt);
  CommBuffer = array[0..99] of byte;

  PackagesState = record
    PakType : byte;
    SendPaks : integer;     //���������� ������� 0..1000
    ErrPaks : integer;      //������� ��������� ������� �� ����� SendPaks
    LostPaks : integer;     //�������� ������� �� ����� SendPaks
    AllErr : integer;       //����� ������ ����� 0..10 AllErr=10 -> State:=stNetErr
  end;

  AverageStruct = record
    Prod : array[1..9] of single;
    Count : byte;
  end;

  AverageStructPointer = ^AverageStruct;

TMicroNetOPC = class(TComponent)
  private
    Comm1 : TComm;                          //��������� ������� � COM-����� Async32
    RequestTimer : TTimer;                  //������ �������� ����� ���������
    WatchDogTimer : TTimer;                 //������ ��������� �������� ������
    FPackages : array[1..31] of PackagesState;
    FState : array[1..31] of TState;            //������ �������
    FCurrentDev : byte;                         //��������� ������
    FCurrentCmd : byte;                         //������� �������
    FCurrentDat : byte;                         //��������� ������� �������
    FAsyncCmd : array[1..31] of word;           //����������� �������
    FCalibr : array[1..31] of byte;             //������������������ ������ ����������
    FRequest : array of byte;                   //����� �������
    FAnswer : array[0..99] of byte;             //����� ������
    FIndex : byte;                              //������� � ������ ������
    FLogFile : TextFile;                        //���������� ���� ����
    FBatchMode : array[1..31] of boolean;       //����� - "+�������"
    procedure RequestTimerTimer(Sender : TObject);
    procedure SendRequest;                      //�������� ���������� �������
    procedure SendPackage(const Dev: byte; const Cmd: byte; const Dat : byte); //������������ � �������� ������
    procedure Comm1RxChar(Sender : TObject; Count : Integer);  //����� ������ �� ����������
    procedure Parser;                           //������ ����������� ������
    procedure PacketLost(Sender : TObject);     //����� �������
    procedure CloseLogFile;
    procedure SetAsyncCmd(Index : byte; Value : word);
    procedure SetCalibr(Index : byte; Value : byte);
    function GetAverage(Values : AverageStructPointer) : single;
  protected
    { Protected declarations }
  public
    property AsyncCmd[Index : byte] : word write SetAsyncCmd;
    property Calibr[Index : byte] : byte write SetCalibr;
    constructor Create(AOwner:TComponent); override;
    destructor Destroy; override;
  end;

var
  ServerPath : string;
  Recordation : boolean;
  PortReady : boolean;
  LogFileName : string;
  AverageCollection : array[1..31] of AverageStruct;
  pAverageStruct : AverageStructPointer;
  Alive : boolean;

implementation

constructor TMicroNetOPC.Create(AOwner:TComponent);
var
  i : byte;
  ItP : itemProps;
  PkSt : ^PackagesState;
  MicroNetIni: TIniFile;
  IniString : string;
  Svr : integer;
begin
  inherited Create(AOwner);
  RequestTimer := TTimer.Create(Self);      //������ �������� ����� ���������
  RequestTimer.Interval := 20;
  RequestTimer.OnTimer := RequestTimerTimer;
  RequestTimer.Enabled := False;
  WatchDogTimer := TTimer.Create(Self);     //������ ��������� �������� ������
  WatchDogTimer.Interval := 80;
  WatchDogTimer.Enabled := False;
  WatchDogTimer.OnTimer := PacketLost;
  Comm1 := TComm.Create(Self);              //��������� ������� � COM-����� Async32
  MicroNetIni := TIniFile.Create(ServerPath + 'MicroNet.ini');
  IniString := MicroNetIni.ReadString('Port', 'PortName', 'Com1');
  if IniString = 'Com2' then Comm1.DeviceName :='Com2' else
    begin
      Comm1.DeviceName :='Com1';
      MicroNetIni.WriteString('Port', 'PortName', 'Com1');
    end;
  IniString := MicroNetIni.ReadString('Port', 'BaudRate', '9600');
  if IniString = '2400' then Comm1.BaudRate := br2400 else
  if IniString = '4800' then Comm1.BaudRate := br4800 else
  if IniString = '14400' then Comm1.BaudRate := br14400 else
  if IniString = '19200' then Comm1.BaudRate := br19200 else
    begin
      Comm1.BaudRate := br9600;
      MicroNetIni.WriteString('Port', 'BaudRate', '9600');
    end;
  Comm1.Parity := paNone;
  Comm1.Stopbits := sb20;
  Comm1.Databits := da8;
  Comm1.MonitorEvents := [evRxChar];
  Comm1.OnRxChar := Comm1RxChar;
  ItP.PropID :=  5000;
  ItP.tagname := 'none';
  ItP.dataType := VT_UI2;
  posItems[0] := ItP;
  for i := 1 to 31 do
    begin
      IniString := MicroNetIni.ReadString('Devices', 'ScanDevice' + IntToStr(i), 'Yes');
      if IniString = 'No' then
        begin
          FState[i] := stNotMnt;
          itemValues[i + 155] := 3;
        end
          else
        begin
          FState[i] := stOk;
          itemValues[i + 155] := 0;
          MicroNetIni.WriteString('Devices', 'ScanDevice' + IntToStr(i), 'Yes');
        end;
      ItP.PropID :=  5000 + i;
      ItP.tagname := 'Prod' + IntToStr(i);
      ItP.dataType := VT_I2;
      posItems[i] := ItP;
      itemValues[i] := 0;
      ItP.PropID :=  5031 + i;
      ItP.tagname := 'Average' + IntToStr(i);
      ItP.dataType := VT_I2;
      posItems[i + 31] := ItP;
      itemValues[i + 31] := 0;
      ItP.PropID :=  5062 + i;
      ItP.tagname := 'Weight' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 62] := ItP;
      itemValues[i + 62] := 0;
      ItP.PropID :=  5093 + i;
      ItP.tagname := 'Summ' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 93] := ItP;
      itemValues[i + 93] := 0;
      ItP.PropID :=  5124 + i;
      ItP.tagname := 'Time' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 124] := ItP;
      itemValues[i + 124] := 0;
      ItP.PropID :=  5155 + i;
      ItP.tagname := 'State' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 155] := ItP;
      ItP.PropID :=  5186 + i;
      ItP.tagname := 'Process' + IntToStr(i);
      ItP.dataType := VT_BOOL;
      posItems[i + 186] := ItP;
      itemValues[i + 186] := 0;
      ItP.PropID :=  5217 + i;
      ItP.tagname := 'SendPaks' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 217] := ItP;
      itemValues[i + 217] := 0;
      ItP.PropID :=  5248 + i;
      ItP.tagname := 'ErrPaks' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 248] := ItP;
      itemValues[i + 248] := 0;
      ItP.PropID :=  5279 + i;
      ItP.tagname := 'LostPaks' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 279] := ItP;
      itemValues[i + 279] := 0;
      PkSt := @FPackages[i];
      PkSt.PakType := 0;
      PkSt.SendPaks := 0;
      PkSt.ErrPaks := 0;
      PkSt.LostPaks := 0;
      PkSt.AllErr := 0;
      FAsyncCmd[i] := 0;                       //����� ����������� �������
      FCalibr[i] := 0;
      FBatchMode[i] := True;
    end;
  for Svr:= low(theServers) to high(theServers) do
    if theServers[Svr] <> nil then theServers[Svr].TimeSlice(Now);
  FCurrentDev := 0;
  IniString := MicroNetIni.ReadString('Settings', 'CreateLogFile', 'Yes');
  if IniString = 'No' then Recordation := False else
    begin
      Recordation := True;
      MicroNetIni.WriteString('Settings', 'CreateLogFile', 'Yes');
      LogFileName := 'MicroNet ' + FormatDateTime('dd"-"mm"-"yyyy" "hh"-"nn', Now) + '.log';
      AssignFile(FLogFile, ServerPath + LogFileName);     //������� ���������� ����
      Rewrite(FLogFile);
      Writeln(FLogFile, '������ ������: ' + FormatDateTime('dd/mmmm/yyyy hh:nn:ss', Now));
    end;
  MicroNetIni.Free;
  try
    Comm1.Open;
    Comm1.SetRTSState(True);
    Comm1.SetDTRState(True);
    RequestTimer.Enabled := True;
    PortReady := True;
  except
    on ECommError do
      begin
        PortReady := False;
        Application.Terminate;
      end;
  end;
end;

destructor TMicroNetOPC.Destroy;
begin
  if PortReady then
    begin
      RequestTimer.Enabled := False;
      Comm1.OnRxChar := nil;
      Comm1.SetRTSState(False);
      Comm1.SetDTRState(False);
      if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
      if Comm1.OutQueCount <> 0 then Comm1.PurgeOut;
      Comm1.Close;                              //������� ����
    end;
  if Recordation then CloseLogFile;
  Comm1.Destroy;
  inherited Destroy;
end;

procedure TMicroNetOPC.RequestTimerTimer(Sender: TObject);
begin
  SendRequest;
end;

procedure TMicroNetOPC.SendRequest;
var
  Svr : integer;
  PkSt : ^PackagesState;
begin
  WatchDogTimer.Enabled := False;   //��������� ������ �������� ������
  RequestTimer.Enabled := False;        //����� ������� �������
  if FCurrentDev <> 31 then FCurrentDev := FCurrentDev + 1 else FCurrentDev := 1; //����������� ���������
  while (FState[FCurrentDev] = stNotMnt) and (FCurrentDev <> 31) do FCurrentDev := FCurrentDev + 1;
  if FCurrentDev = 31 then
    begin
      for Svr := low(theServers) to high(theServers) do
        if theServers[Svr] <> nil then theServers[Svr].TimeSlice(Now);
        Application.ProcessMessages;
    end;
  RequestTimer.Enabled := True;
  if FState[FCurrentDev] <> stNotMnt then  //������ ����������
    begin
      if FCalibr[FCurrentDev] <> 0 then
        begin
          case FCalibr[FCurrentDev] of
            1:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $2C;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            2:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            3:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            4:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $64;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            5:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            6:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            7:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $42;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            8:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            9:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            10:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $64;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := 0;
                end;
          end;
        end else
        begin
          PkSt := @FPackages[FCurrentDev];
          if FAsyncCmd[FCurrentDev] <> 0 then //����������� ����������� �������
            begin
              if FAsyncCmd[FCurrentDev] = $4B52 then
                begin
                  FCurrentCmd := $4B;   //�����
                  FCurrentDat := $52;   //
                end;
              if FAsyncCmd[FCurrentDev] = $4B51 then
                begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $51;   //������� ��������� ��������
                end;
              if FAsyncCmd[FCurrentDev] = $4B54 then
                begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $54;   //������� �����
                end;
              if FAsyncCmd[FCurrentDev] = $4B58 then
                begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $58;   //������� ����
                end;
            end else
              begin
                if PkSt.PakType < 7 then
                  begin
                    FCurrentCmd := $56; //������ �������� �� ���-�����
                    FCurrentDat := $47; //���-�����: �����, ������������������, �����, ���������
                    PkSt.PakType := PkSt.PakType + 1;
                  end else
                  begin
                    FCurrentCmd := $2E;   //������ ����� �� ���-�����
                    FCurrentDat := $A6;   //���-�����: ������, �����, ���������, ������ 485
                    PkSt.PakType := 0;
                  end;
              end;
          PkSt.SendPaks := PkSt.SendPaks + 1;
          itemValues[FCurrentDev + 217] := PkSt.SendPaks;
          if PkSt.SendPaks >= 1000 then
            begin
              PkSt.SendPaks := 0;
              PkSt.ErrPaks := 0;
              PkSt.LostPaks := 0;
            end;
           SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
        end;
    end;
end;

procedure TMicroNetOPC.SendPackage(const Dev : byte; const Cmd : byte; const Dat : byte);
var
  CRC : byte;     //����������� �����
  i : byte;       //������� � ������ �������
  S : string;     //���������� ������������� ������ �������
  TempDat : byte; //�������� �������
  Buffer1 : ^CommBuffer;
begin
  if Dat <> 0 then SetLength(FRequest, 7) else SetLength(FRequest, 6);
  FRequest[0] := $FF;         //������ ������
  FRequest[1] := $20 + Dev;   //����� �������
  FRequest[2] := $20;         //����� �������
  FRequest[3] := Cmd;         //�������
  i := 4;
  if Dat <> 0 then
    begin
    //������ ����������� �����
      CRC := $FF xor ($20 + Dev) xor $20 xor Cmd xor Dat;
      TempDat := Dat;
      if TempDat in [$FF, $03, $10] then  //��������� ������ ��� �������������
        begin
          SetLength(FRequest, High(FRequest) + 2);
          FRequest[i] := $10;
          i := i + 1;
          TempDat := $FF - TempDat;
        end;
      FRequest[i] := TempDat;
      i := i + 1;
    end else CRC := $FF xor ($20 + Dev) xor $20 xor Cmd;
  if CRC in [$FF, $03, $10] then  //��������� ����������� ����� ��� �������������
    begin
      SetLength(FRequest, High(FRequest) + 2);
      FRequest[i] := $10;
      i := i + 1;
      CRC := $FF - CRC;
    end;
  FRequest[i] := CRC;             //����������� �����
  i := i + 1;
  FRequest[i] := $03;             //����� ������
  GetMem(Buffer1, High(FRequest) + 2);
  S := '';
  for i := 0 to High(FRequest) do
    begin
      S := S + IntToHex(FRequest[i], 2) + ' ';  //������������ ����������� ������������� ������ �������
      CommBuffer(Buffer1^)[i] := FRequest[i];
    end;
  Comm1.Write(Buffer1^, High(FRequest) + 1);  //�������� ������� ����������
  FreeMem(Buffer1);
  Comm1.SetDTRState(True);
  if Recordation then Writeln(FLogFile, FormatDateTime('dd/mm/yyyy hh:nn:ss:zzz', Now) + ';0;' + S);  //����������������
  FIndex := 0;                          //����� ������� � ������ ������
  for i := 0 to 99 do FAnswer[i] := 0;  //������� ������ ������
  RequestTimer.Enabled := False;        //����� ������� �������
  WatchDogTimer.Enabled := True;        //������ ������� �������� ������
end;

procedure TMicroNetOPC.Comm1RxChar(Sender: TObject; Count: Integer);
var
  Buffer2 : ^CommBuffer;
  Bytes : integer;
  i : byte;
begin
      GetMem(Buffer2, Comm1.ReadBufSize);
      try
        Bytes := Comm1.Read(Buffer2^, Count);
        if Bytes <> - 1 then
          begin
            i := 0;
            repeat
              FAnswer[FIndex] := CommBuffer(Buffer2^)[i];
              FIndex := FIndex + 1;
              i := i + 1;
            until (i = Bytes) or (FAnswer[FIndex - 1] = $03);
            if FAnswer[FIndex - 1] = $03 then Parser else
              if FIndex > 30 then   //��������� �������� ������ ��� ���� ����������
                begin
                  if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
                  FIndex := 0;                            //����� ������� � ������ ������
                  for i := 0 to 99 do FAnswer[i] := 0;    //������� ������ ������
                end;
          end;
      finally
        FreeMem(Buffer2);
      end;
end;

procedure TMicroNetOPC.Parser;
var
  CleanAnswer : array of byte;  //����� ������ ��� ������������
  CleanIndex : byte;            //������� � ������ CleanAnswer
  S : string;
  CRC, i : byte;
  TempInt : word;
  TempLong : Longword;
  TempLong1 : Longword;
  pSingle : ^Single;
  pSmallint : ^Smallint;
  PkSt : ^PackagesState;
  Aver : single;
begin
  if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
  WatchDogTimer.Enabled := False;   //��������� ������ �������� ������
  PkSt := @FPackages[FCurrentDev];
  FIndex := FIndex - 1;             //����� ���������� ����� ������ ������
  S := '';
  for i := 0 to FIndex do S := S + IntToHex(FAnswer[i], 2) + ' ';
  if Recordation then Write(FLogFile, FormatDateTime('dd/mm/yyyy hh:nn:ss:zzz', Now) + ';1;' + S + ';');  //����������������
  SetLength(CleanAnswer, 1);
  CleanIndex := 0;
  i := 0;
  repeat
    if FAnswer[i] <> $10 then
      begin
        CleanAnswer[CleanIndex] := FAnswer[i];
        SetLength(CleanAnswer, High(CleanAnswer) + 2);
        CleanIndex := CleanIndex + 1;
        i := i + 1;
      end
      else
      begin
        CleanAnswer[CleanIndex] := $FF - FAnswer[i + 1];
        SetLength(CleanAnswer, High(CleanAnswer) + 2);
        CleanIndex := CleanIndex + 1;
        i := i + 2;
      end;
  until i > FIndex;
  CleanIndex := CleanIndex - 1;       //����� ���������� ����� ������ CleanAnswer
  CRC := 0;
  for i := 0 to CleanIndex - 2 do CRC := CRC xor CleanAnswer[i];
  if (CRC <> CleanAnswer[CleanIndex - 1]) or (CleanAnswer[2] <> FCurrentDev + $20) then
    begin
      if Recordation then Writeln(FLogFile, '��������� �����!!! CRC = ' + IntToHex(CRC, 2) + 'h');
      PkSt.ErrPaks := PkSt.ErrPaks + 1;
      PkSt.AllErr := PkSt.AllErr + 1;
      if PkSt.AllErr > 10 then  //������ ����� � ��������
        begin
          FState[FCurrentDev] := stNetErr;
          if FAsyncCmd[FCurrentDev] <> 0 then
            begin
              FAsyncCmd[FCurrentDev] := 0;       //����� ����� ����������� ������� (�� ���������)
            end;
        end;
    end else
    begin
      if Recordation then Writeln(FLogFile, 'OK');
      if CleanAnswer[3] = $4B then FAsyncCmd[FCurrentDev] := 0;  //����� ����� ����������� ������� (���������)
      PkSt.AllErr := 0;
      if CleanAnswer[3] = $2E then
        begin
          if (CleanAnswer[13] in [$96, $97]) then
            FState[FCurrentDev] := stSenErr else FState[FCurrentDev] := stOk; //�������� ����������� �������
          //���� ������ � �����, �� �������� �� � ������ "+�������" - ������� � �����
          if (FState[FCurrentDev] = stOk) and (not FBatchMode[FCurrentDev]) then FAsyncCmd[FCurrentDev] := $4B51;
          TempInt := 256 * CleanAnswer[8] + CleanAnswer[9];
          pSmallint := @TempInt;
          if pSmallint^ > 0 then itemValues[62 + FCurrentDev] := pSmallint^ else itemValues[62 + FCurrentDev] := 0;
        end;
      if CleanAnswer[3] = $56 then
        begin
          //��������� �������� �����
          try
            if CleanAnswer[5] < 127 then
              begin
                TempLong := 256 * CleanAnswer[5] + CleanAnswer[6];
                TempLong1 := 256 * CleanAnswer[7] + CleanAnswer[8];
                TempLong := TempLong * 65536 + TempLong1;
              end else TempLong := 0;
            pSingle := @TempLong;
            if pSingle^ > 0 then pSingle^ := pSingle^ / 1000 else pSingle^ := 0;
            if pSingle^ < 65535 then itemValues[93 + FCurrentDev] := Round(pSingle^) else itemValues[93 + FCurrentDev] := 0;
          except
            itemValues[93 + FCurrentDev] := 0;
          end;

          //��������� �������� ������������������
          try
            TempLong := 256 * CleanAnswer[9] + CleanAnswer[10];
            TempLong1 := 256 * CleanAnswer[11] + CleanAnswer[12];
            TempLong := TempLong * 65536 + TempLong1;
            pSingle := @TempLong;
            if (pSingle^ > -32767) and (pSingle^ < 32767) then
              begin
                if pSingle^ >= 0 then itemValues[FCurrentDev] := Round(pSingle^) else
                  itemValues[FCurrentDev] := Round(Abs(pSingle^ + 65536));
                pAverageStruct := @AverageCollection[FCurrentDev];
                if itemValues[186 + FCurrentDev] = 1 then
                  begin
                    for i := High(pAverageStruct.Prod) - 1 downto 1 do pAverageStruct.Prod[i + 1] := pAverageStruct.Prod[i];
                    pAverageStruct.Prod[1] := pSingle^;
                    if pAverageStruct.Count < High(pAverageStruct.Prod) then pAverageStruct.Count := pAverageStruct.Count + 1;
                    Aver := GetAverage(pAverageStruct);
                    if Aver >= 0 then itemValues[31 + FCurrentDev] := Round(Aver) else
                      itemValues[31 + FCurrentDev] := Round(Abs(Aver + 65536));
                  end else
                  begin
                    pAverageStruct.Count := 0;
                    itemValues[31 + FCurrentDev] := 0;
                  end;
              end else itemValues[FCurrentDev] := 0;
          except
            itemValues[FCurrentDev] := 0;
          end;

          //��������� �������� �������
          try
            if CleanAnswer[13] < 127 then
              begin
                TempLong := 256 * CleanAnswer[13] + CleanAnswer[14];
                TempLong1 := 256 * CleanAnswer[15] + CleanAnswer[16];
                TempLong := TempLong * 65536 + TempLong1;
              end else TempLong := 0;
            if (TempLong > 0) and (TempLong div 228 < 65535) then
              itemValues[FCurrentDev + 124] := Round(TempLong div 228) else
              itemValues[FCurrentDev + 124] := 0;
          except
            itemValues[FCurrentDev + 124] := 0;
          end;
          //�������� ��������� "����"
          if CleanAnswer[18] = $02 then itemValues[186 + FCurrentDev] := 1 else itemValues[186 + FCurrentDev] := 0;
          //�������� ��������� "+�������"
          if CleanAnswer[18] = $00 then
            begin
              FBatchMode[FCurrentDev] := False;
              FState[FCurrentDev] := stSenErr;
            end else FBatchMode[FCurrentDev] := True;
        end;
    end;
  if FState[FCurrentDev] = stOk then itemValues[155 + FCurrentDev] := 0;
  if FState[FCurrentDev] = stSenErr then itemValues[155 + FCurrentDev] := 1;
  if FState[FCurrentDev] = stNetErr then itemValues[155 + FCurrentDev] := 2;
  itemValues[248 + FCurrentDev] := PkSt.ErrPaks;
  itemValues[279 + FCurrentDev] := PkSt.LostPaks;
  Alive := True;
  RequestTimer.Enabled := True;
end;

procedure TMicroNetOPC.PacketLost(Sender : TObject);
var
  PkSt : ^PackagesState;
begin
  if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
  PkSt := @FPackages[FCurrentDev];
  PkSt.LostPaks := PkSt.LostPaks + 1;
  PkSt.AllErr := PkSt.AllErr + 1;
  if PkSt.AllErr > 10 then  //������ ����� � ��������
    begin
      FState[FCurrentDev] := stNetErr;
      if FAsyncCmd[FCurrentDev] <> 0 then
        begin
          FAsyncCmd[FCurrentDev] := 0;       //����� ����� ����������� ������� (�� ���������)
        end;
    end;
  if FState[FCurrentDev] = stOk then itemValues[155 + FCurrentDev] := 0;
  if FState[FCurrentDev] = stSenErr then itemValues[155 + FCurrentDev] := 1;
  if FState[FCurrentDev] = stNetErr then itemValues[155 + FCurrentDev] := 2;
  if FState[FCurrentDev] = stNotMnt then itemValues[155 + FCurrentDev] := 3;
  itemValues[279 + FCurrentDev] := PkSt.LostPaks;
  RequestTimer.Enabled := True;     //�������� ������ �������
  WatchDogTimer.Enabled := False;   //��������� ������ �������� ������
end;

procedure TMicroNetOPC.CloseLogFile;
begin
  Recordation := False;
  Writeln(FLogFile, '���������� ������: ' + FormatDateTime('dd/mmmm/yyyy hh:nn:ss', Now));
  CloseFile(FLogFile);                      //������� ���������� ����
end;

procedure TMicroNetOPC.SetAsyncCmd(Index : byte; Value : word);
begin
  FAsyncCmd[Index] := Value;
end;

procedure TMicroNetOPC.SetCalibr(Index : byte; Value : byte);
begin
  if itemValues[Index + 186] = 0 then FCalibr[Index] := Value;
end;

function TMicroNetOPC.GetAverage(Values : AverageStructPointer) : single;
var
  SortCounter : byte;
  SortTemp : single;
  i, ArrayLength : byte;
  TempVar : single;
  TempArray : array of single;
begin
  ArrayLength := High(pAverageStruct.Prod);
  if pAverageStruct.Count < ArrayLength then
    begin
      TempVar := 0;
      for i := 1 to pAverageStruct.Count do TempVar := TempVar + pAverageStruct.Prod[i];
      Result := TempVar / pAverageStruct.Count;
    end else
    begin
      SetLength(TempArray, ArrayLength + 1);
      for i := 1 to ArrayLength do TempArray[i] := pAverageStruct.Prod[i];
      repeat //����������
      SortCounter := 0;
      for i := 1 to ArrayLength - 1 do
        begin
          if TempArray[i] > TempArray[i + 1] then
            begin
              SortTemp := TempArray[i];
              TempArray[i] := TempArray[i + 1];
              TempArray[i + 1] := SortTemp;
              SortCounter := SortCounter + 1;
            end;
        end;
      until SortCounter = 0;
      Result := (TempArray[ArrayLength div 2 - 2] + TempArray[ArrayLength div 2 - 1] + 
        TempArray[ArrayLength div 2] + TempArray[ArrayLength div 2 + 1] + 
        TempArray[ArrayLength div 2 + 2]) / 5;
    end;
end;

end.
unit MicroNetOPC;

interface

uses SysUtils, Classes, Controls, StdCtrls, AxCtrls, ExtCtrls,
    CommPort, Globals, ActiveX, ServIMPL, IniFiles, Forms;

type
  TState = (stOk, stSenErr, stNetErr, stNotMnt);
  CommBuffer = array[0..99] of byte;

  PackagesState = record
    PakType : byte;
    SendPaks : integer;     //���������� ������� 0..1000
    ErrPaks : integer;      //������� ��������� ������� �� ����� SendPaks
    LostPaks : integer;     //�������� ������� �� ����� SendPaks
    AllErr : integer;       //����� ������ ����� 0..10 AllErr=10 -> State:=stNetErr
  end;

  AverageStruct = record
    Prod : array[1..9] of single;
    Count : byte;
  end;

  AverageStructPointer = ^AverageStruct;

TMicroNetOPC = class(TComponent)
  private
    Comm1 : TComm;                          //��������� ������� � COM-����� Async32
    RequestTimer : TTimer;                  //������ �������� ����� ���������
    WatchDogTimer : TTimer;                 //������ ��������� �������� ������
    FPackages : array[1..31] of PackagesState;
    FState : array[1..31] of TState;            //������ �������
    FCurrentDev : byte;                         //��������� ������
    FCurrentCmd : byte;                         //������� �������
    FCurrentDat : byte;                         //��������� ������� �������
    FAsyncCmd : array[1..31] of word;           //����������� �������
    FCalibr : array[1..31] of byte;             //������������������ ������ ����������
    FRequest : array of byte;                   //����� �������
    FAnswer : array[0..99] of byte;             //����� ������
    FIndex : byte;                              //������� � ������ ������
    FLogFile : TextFile;                        //���������� ���� ����
    FBatchMode : array[1..31] of boolean;       //����� - "+�������"
    procedure RequestTimerTimer(Sender : TObject);
    procedure SendRequest;                      //�������� ���������� �������
    procedure SendPackage(const Dev: byte; const Cmd: byte; const Dat : byte); //������������ � �������� ������
    procedure Comm1RxChar(Sender : TObject; Count : Integer);  //����� ������ �� ����������
    procedure Parser;                           //������ ����������� ������
    procedure PacketLost(Sender : TObject);     //����� �������
    procedure CloseLogFile;
    procedure SetAsyncCmd(Index : byte; Value : word);
    procedure SetCalibr(Index : byte; Value : byte);
    function GetAverage(Values : AverageStructPointer) : single;
  protected
    { Protected declarations }
  public
    property AsyncCmd[Index : byte] : word write SetAsyncCmd;
    property Calibr[Index : byte] : byte write SetCalibr;
    constructor Create(AOwner:TComponent); override;
    destructor Destroy; override;
  end;

var
  ServerPath : string;
  Recordation : boolean;
  PortReady : boolean;
  LogFileName : string;
  AverageCollection : array[1..31] of AverageStruct;
  pAverageStruct : AverageStructPointer;
  Alive : boolean;

implementation

constructor TMicroNetOPC.Create(AOwner:TComponent);
var
  i : byte;
  ItP : itemProps;
  PkSt : ^PackagesState;
  MicroNetIni: TIniFile;
  IniString : string;
  Svr : integer;
begin
  inherited Create(AOwner);
  RequestTimer := TTimer.Create(Self);      //������ �������� ����� ���������
  RequestTimer.Interval := 20;
  RequestTimer.OnTimer := RequestTimerTimer;
  RequestTimer.Enabled := False;
  WatchDogTimer := TTimer.Create(Self);     //������ ��������� �������� ������
  WatchDogTimer.Interval := 80;
  WatchDogTimer.Enabled := False;
  WatchDogTimer.OnTimer := PacketLost;
  Comm1 := TComm.Create(Self);              //��������� ������� � COM-����� Async32
  MicroNetIni := TIniFile.Create(ServerPath + 'MicroNet.ini');
  IniString := MicroNetIni.ReadString('Port', 'PortName', 'Com1');
  if IniString = 'Com2' then Comm1.DeviceName :='Com2' else
    begin
      Comm1.DeviceName :='Com1';
      MicroNetIni.WriteString('Port', 'PortName', 'Com1');
    end;
  IniString := MicroNetIni.ReadString('Port', 'BaudRate', '9600');
  if IniString = '2400' then Comm1.BaudRate := br2400 else
  if IniString = '4800' then Comm1.BaudRate := br4800 else
  if IniString = '14400' then Comm1.BaudRate := br14400 else
  if IniString = '19200' then Comm1.BaudRate := br19200 else
    begin
      Comm1.BaudRate := br9600;
      MicroNetIni.WriteString('Port', 'BaudRate', '9600');
    end;
  Comm1.Parity := paNone;
  Comm1.Stopbits := sb20;
  Comm1.Databits := da8;
  Comm1.MonitorEvents := [evRxChar];
  Comm1.OnRxChar := Comm1RxChar;
  ItP.PropID :=  5000;
  ItP.tagname := 'none';
  ItP.dataType := VT_UI2;
  posItems[0] := ItP;
  for i := 1 to 31 do
    begin
      IniString := MicroNetIni.ReadString('Devices', 'ScanDevice' + IntToStr(i), 'Yes');
      if IniString = 'No' then
        begin
          FState[i] := stNotMnt;
          itemValues[i + 155] := 3;
        end
          else
        begin
          FState[i] := stOk;
          itemValues[i + 155] := 0;
          MicroNetIni.WriteString('Devices', 'ScanDevice' + IntToStr(i), 'Yes');
        end;
      ItP.PropID :=  5000 + i;
      ItP.tagname := 'Prod' + IntToStr(i);
      ItP.dataType := VT_I2;
      posItems[i] := ItP;
      itemValues[i] := 0;
      ItP.PropID :=  5031 + i;
      ItP.tagname := 'Average' + IntToStr(i);
      ItP.dataType := VT_I2;
      posItems[i + 31] := ItP;
      itemValues[i + 31] := 0;
      ItP.PropID :=  5062 + i;
      ItP.tagname := 'Weight' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 62] := ItP;
      itemValues[i + 62] := 0;
      ItP.PropID :=  5093 + i;
      ItP.tagname := 'Summ' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 93] := ItP;
      itemValues[i + 93] := 0;
      ItP.PropID :=  5124 + i;
      ItP.tagname := 'Time' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 124] := ItP;
      itemValues[i + 124] := 0;
      ItP.PropID :=  5155 + i;
      ItP.tagname := 'State' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 155] := ItP;
      ItP.PropID :=  5186 + i;
      ItP.tagname := 'Process' + IntToStr(i);
      ItP.dataType := VT_BOOL;
      posItems[i + 186] := ItP;
      itemValues[i + 186] := 0;
      ItP.PropID :=  5217 + i;
      ItP.tagname := 'SendPaks' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 217] := ItP;
      itemValues[i + 217] := 0;
      ItP.PropID :=  5248 + i;
      ItP.tagname := 'ErrPaks' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 248] := ItP;
      itemValues[i + 248] := 0;
      ItP.PropID :=  5279 + i;
      ItP.tagname := 'LostPaks' + IntToStr(i);
      ItP.dataType := VT_UI2;
      posItems[i + 279] := ItP;
      itemValues[i + 279] := 0;
      PkSt := @FPackages[i];
      PkSt.PakType := 0;
      PkSt.SendPaks := 0;
      PkSt.ErrPaks := 0;
      PkSt.LostPaks := 0;
      PkSt.AllErr := 0;
      FAsyncCmd[i] := 0;                       //����� ����������� �������
      FCalibr[i] := 0;
      FBatchMode[i] := True;
    end;
  for Svr:= low(theServers) to high(theServers) do
    if theServers[Svr] <> nil then theServers[Svr].TimeSlice(Now);
  FCurrentDev := 0;
  IniString := MicroNetIni.ReadString('Settings', 'CreateLogFile', 'Yes');
  if IniString = 'No' then Recordation := False else
    begin
      Recordation := True;
      MicroNetIni.WriteString('Settings', 'CreateLogFile', 'Yes');
      LogFileName := 'MicroNet ' + FormatDateTime('dd"-"mm"-"yyyy" "hh"-"nn', Now) + '.log';
      AssignFile(FLogFile, ServerPath + LogFileName);     //������� ���������� ����
      Rewrite(FLogFile);
      Writeln(FLogFile, '������ ������: ' + FormatDateTime('dd/mmmm/yyyy hh:nn:ss', Now));
    end;
  MicroNetIni.Free;
  try
    Comm1.Open;
    Comm1.SetRTSState(True);
    Comm1.SetDTRState(True);
    RequestTimer.Enabled := True;
    PortReady := True;
  except
    on ECommError do
      begin
        PortReady := False;
        Application.Terminate;
      end;
  end;
end;

destructor TMicroNetOPC.Destroy;
begin
  if PortReady then
    begin
      RequestTimer.Enabled := False;
      Comm1.OnRxChar := nil;
      Comm1.SetRTSState(False);
      Comm1.SetDTRState(False);
      if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
      if Comm1.OutQueCount <> 0 then Comm1.PurgeOut;
      Comm1.Close;                              //������� ����
    end;
  if Recordation then CloseLogFile;
  Comm1.Destroy;
  inherited Destroy;
end;

procedure TMicroNetOPC.RequestTimerTimer(Sender: TObject);
begin
  SendRequest;
end;

procedure TMicroNetOPC.SendRequest;
var
  Svr : integer;
  PkSt : ^PackagesState;
begin
  WatchDogTimer.Enabled := False;   //��������� ������ �������� ������
  RequestTimer.Enabled := False;        //����� ������� �������
  if FCurrentDev <> 31 then FCurrentDev := FCurrentDev + 1 else FCurrentDev := 1; //����������� ���������
  while (FState[FCurrentDev] = stNotMnt) and (FCurrentDev <> 31) do FCurrentDev := FCurrentDev + 1;
  if FCurrentDev = 31 then
    begin
      for Svr := low(theServers) to high(theServers) do
        if theServers[Svr] <> nil then theServers[Svr].TimeSlice(Now);
        Application.ProcessMessages;
    end;
  RequestTimer.Enabled := True;
  if FState[FCurrentDev] <> stNotMnt then  //������ ����������
    begin
      if FCalibr[FCurrentDev] <> 0 then
        begin
          case FCalibr[FCurrentDev] of
            1:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $2C;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            2:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            3:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            4:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $64;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            5:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            6:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            7:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $42;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            8:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            9:  begin
                  FCalibr[FCurrentDev] := FCalibr[FCurrentDev] + 1;
                end;
            10:  begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $64;   //
                  SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
                  FCalibr[FCurrentDev] := 0;
                end;
          end;
        end else
        begin
          PkSt := @FPackages[FCurrentDev];
          if FAsyncCmd[FCurrentDev] <> 0 then //����������� ����������� �������
            begin
              if FAsyncCmd[FCurrentDev] = $4B52 then
                begin
                  FCurrentCmd := $4B;   //�����
                  FCurrentDat := $52;   //
                end;
              if FAsyncCmd[FCurrentDev] = $4B51 then
                begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $51;   //������� ��������� ��������
                end;
              if FAsyncCmd[FCurrentDev] = $4B54 then
                begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $54;   //������� �����
                end;
              if FAsyncCmd[FCurrentDev] = $4B58 then
                begin
                  FCurrentCmd := $4B;   //
                  FCurrentDat := $58;   //������� ����
                end;
            end else
              begin
                if PkSt.PakType < 7 then
                  begin
                    FCurrentCmd := $56; //������ �������� �� ���-�����
                    FCurrentDat := $47; //���-�����: �����, ������������������, �����, ���������
                    PkSt.PakType := PkSt.PakType + 1;
                  end else
                  begin
                    FCurrentCmd := $2E;   //������ ����� �� ���-�����
                    FCurrentDat := $A6;   //���-�����: ������, �����, ���������, ������ 485
                    PkSt.PakType := 0;
                  end;
              end;
          PkSt.SendPaks := PkSt.SendPaks + 1;
          itemValues[FCurrentDev + 217] := PkSt.SendPaks;
          if PkSt.SendPaks >= 1000 then
            begin
              PkSt.SendPaks := 0;
              PkSt.ErrPaks := 0;
              PkSt.LostPaks := 0;
            end;
           SendPackage(FCurrentDev, FCurrentCmd, FCurrentDat);
        end;
    end;
end;

procedure TMicroNetOPC.SendPackage(const Dev : byte; const Cmd : byte; const Dat : byte);
var
  CRC : byte;     //����������� �����
  i : byte;       //������� � ������ �������
  S : string;     //���������� ������������� ������ �������
  TempDat : byte; //�������� �������
  Buffer1 : ^CommBuffer;
begin
  if Dat <> 0 then SetLength(FRequest, 7) else SetLength(FRequest, 6);
  FRequest[0] := $FF;         //������ ������
  FRequest[1] := $20 + Dev;   //����� �������
  FRequest[2] := $20;         //����� �������
  FRequest[3] := Cmd;         //�������
  i := 4;
  if Dat <> 0 then
    begin
    //������ ����������� �����
      CRC := $FF xor ($20 + Dev) xor $20 xor Cmd xor Dat;
      TempDat := Dat;
      if TempDat in [$FF, $03, $10] then  //��������� ������ ��� �������������
        begin
          SetLength(FRequest, High(FRequest) + 2);
          FRequest[i] := $10;
          i := i + 1;
          TempDat := $FF - TempDat;
        end;
      FRequest[i] := TempDat;
      i := i + 1;
    end else CRC := $FF xor ($20 + Dev) xor $20 xor Cmd;
  if CRC in [$FF, $03, $10] then  //��������� ����������� ����� ��� �������������
    begin
      SetLength(FRequest, High(FRequest) + 2);
      FRequest[i] := $10;
      i := i + 1;
      CRC := $FF - CRC;
    end;
  FRequest[i] := CRC;             //����������� �����
  i := i + 1;
  FRequest[i] := $03;             //����� ������
  GetMem(Buffer1, High(FRequest) + 2);
  S := '';
  for i := 0 to High(FRequest) do
    begin
      S := S + IntToHex(FRequest[i], 2) + ' ';  //������������ ����������� ������������� ������ �������
      CommBuffer(Buffer1^)[i] := FRequest[i];
    end;
  Comm1.Write(Buffer1^, High(FRequest) + 1);  //�������� ������� ����������
  FreeMem(Buffer1);
  Comm1.SetDTRState(True);
  if Recordation then Writeln(FLogFile, FormatDateTime('dd/mm/yyyy hh:nn:ss:zzz', Now) + ';0;' + S);  //����������������
  FIndex := 0;                          //����� ������� � ������ ������
  for i := 0 to 99 do FAnswer[i] := 0;  //������� ������ ������
  RequestTimer.Enabled := False;        //����� ������� �������
  WatchDogTimer.Enabled := True;        //������ ������� �������� ������
end;

procedure TMicroNetOPC.Comm1RxChar(Sender: TObject; Count: Integer);
var
  Buffer2 : ^CommBuffer;
  Bytes : integer;
  i : byte;
begin
      GetMem(Buffer2, Comm1.ReadBufSize);
      try
        Bytes := Comm1.Read(Buffer2^, Count);
        if Bytes <> - 1 then
          begin
            i := 0;
            repeat
              FAnswer[FIndex] := CommBuffer(Buffer2^)[i];
              FIndex := FIndex + 1;
              i := i + 1;
            until (i = Bytes) or (FAnswer[FIndex - 1] = $03);
            if FAnswer[FIndex - 1] = $03 then Parser else
              if FIndex > 30 then   //��������� �������� ������ ��� ���� ����������
                begin
                  if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
                  FIndex := 0;                            //����� ������� � ������ ������
                  for i := 0 to 99 do FAnswer[i] := 0;    //������� ������ ������
                end;
          end;
      finally
        FreeMem(Buffer2);
      end;
end;

procedure TMicroNetOPC.Parser;
var
  CleanAnswer : array of byte;  //����� ������ ��� ������������
  CleanIndex : byte;            //������� � ������ CleanAnswer
  S : string;
  CRC, i : byte;
  TempInt : word;
  TempLong : Longword;
  TempLong1 : Longword;
  pSingle : ^Single;
  pSmallint : ^Smallint;
  PkSt : ^PackagesState;
  Aver : single;
begin
  if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
  WatchDogTimer.Enabled := False;   //��������� ������ �������� ������
  PkSt := @FPackages[FCurrentDev];
  FIndex := FIndex - 1;             //����� ���������� ����� ������ ������
  S := '';
  for i := 0 to FIndex do S := S + IntToHex(FAnswer[i], 2) + ' ';
  if Recordation then Write(FLogFile, FormatDateTime('dd/mm/yyyy hh:nn:ss:zzz', Now) + ';1;' + S + ';');  //����������������
  SetLength(CleanAnswer, 1);
  CleanIndex := 0;
  i := 0;
  repeat
    if FAnswer[i] <> $10 then
      begin
        CleanAnswer[CleanIndex] := FAnswer[i];
        SetLength(CleanAnswer, High(CleanAnswer) + 2);
        CleanIndex := CleanIndex + 1;
        i := i + 1;
      end
      else
      begin
        CleanAnswer[CleanIndex] := $FF - FAnswer[i + 1];
        SetLength(CleanAnswer, High(CleanAnswer) + 2);
        CleanIndex := CleanIndex + 1;
        i := i + 2;
      end;
  until i > FIndex;
  CleanIndex := CleanIndex - 1;       //����� ���������� ����� ������ CleanAnswer
  CRC := 0;
  for i := 0 to CleanIndex - 2 do CRC := CRC xor CleanAnswer[i];
  if (CRC <> CleanAnswer[CleanIndex - 1]) or (CleanAnswer[2] <> FCurrentDev + $20) then
    begin
      if Recordation then Writeln(FLogFile, '��������� �����!!! CRC = ' + IntToHex(CRC, 2) + 'h');
      PkSt.ErrPaks := PkSt.ErrPaks + 1;
      PkSt.AllErr := PkSt.AllErr + 1;
      if PkSt.AllErr > 10 then  //������ ����� � ��������
        begin
          FState[FCurrentDev] := stNetErr;
          if FAsyncCmd[FCurrentDev] <> 0 then
            begin
              FAsyncCmd[FCurrentDev] := 0;       //����� ����� ����������� ������� (�� ���������)
            end;
        end;
    end else
    begin
      if Recordation then Writeln(FLogFile, 'OK');
      if CleanAnswer[3] = $4B then FAsyncCmd[FCurrentDev] := 0;  //����� ����� ����������� ������� (���������)
      PkSt.AllErr := 0;
      if CleanAnswer[3] = $2E then
        begin
          if (CleanAnswer[13] in [$96, $97]) then
            FState[FCurrentDev] := stSenErr else FState[FCurrentDev] := stOk; //�������� ����������� �������
          //���� ������ � �����, �� �������� �� � ������ "+�������" - ������� � �����
          if (FState[FCurrentDev] = stOk) and (not FBatchMode[FCurrentDev]) then FAsyncCmd[FCurrentDev] := $4B51;
          TempInt := 256 * CleanAnswer[8] + CleanAnswer[9];
          pSmallint := @TempInt;
          if pSmallint^ > 0 then itemValues[62 + FCurrentDev] := pSmallint^ else itemValues[62 + FCurrentDev] := 0;
        end;
      if CleanAnswer[3] = $56 then
        begin
          //��������� �������� �����
          try
            if CleanAnswer[5] < 127 then
              begin
                TempLong := 256 * CleanAnswer[5] + CleanAnswer[6];
                TempLong1 := 256 * CleanAnswer[7] + CleanAnswer[8];
                TempLong := TempLong * 65536 + TempLong1;
              end else TempLong := 0;
            pSingle := @TempLong;
            if pSingle^ > 0 then pSingle^ := pSingle^ / 1000 else pSingle^ := 0;
            if pSingle^ < 65535 then itemValues[93 + FCurrentDev] := Round(pSingle^) else itemValues[93 + FCurrentDev] := 0;
          except
            itemValues[93 + FCurrentDev] := 0;
          end;

          //��������� �������� ������������������
          try
            TempLong := 256 * CleanAnswer[9] + CleanAnswer[10];
            TempLong1 := 256 * CleanAnswer[11] + CleanAnswer[12];
            TempLong := TempLong * 65536 + TempLong1;
            pSingle := @TempLong;
            if (pSingle^ > -32767) and (pSingle^ < 32767) then
              begin
                if pSingle^ >= 0 then itemValues[FCurrentDev] := Round(pSingle^) else
                  itemValues[FCurrentDev] := Round(Abs(pSingle^ + 65536));
                pAverageStruct := @AverageCollection[FCurrentDev];
                if itemValues[186 + FCurrentDev] = 1 then
                  begin
                    for i := High(pAverageStruct.Prod) - 1 downto 1 do pAverageStruct.Prod[i + 1] := pAverageStruct.Prod[i];
                    pAverageStruct.Prod[1] := pSingle^;
                    if pAverageStruct.Count < High(pAverageStruct.Prod) then pAverageStruct.Count := pAverageStruct.Count + 1;
                    Aver := GetAverage(pAverageStruct);
                    if Aver >= 0 then itemValues[31 + FCurrentDev] := Round(Aver) else
                      itemValues[31 + FCurrentDev] := Round(Abs(Aver + 65536));
                  end else
                  begin
                    pAverageStruct.Count := 0;
                    itemValues[31 + FCurrentDev] := 0;
                  end;
              end else itemValues[FCurrentDev] := 0;
          except
            itemValues[FCurrentDev] := 0;
          end;

          //��������� �������� �������
          try
            if CleanAnswer[13] < 127 then
              begin
                TempLong := 256 * CleanAnswer[13] + CleanAnswer[14];
                TempLong1 := 256 * CleanAnswer[15] + CleanAnswer[16];
                TempLong := TempLong * 65536 + TempLong1;
              end else TempLong := 0;
            if (TempLong > 0) and (TempLong div 228 < 65535) then
              itemValues[FCurrentDev + 124] := Round(TempLong div 228) else
              itemValues[FCurrentDev + 124] := 0;
          except
            itemValues[FCurrentDev + 124] := 0;
          end;
          //�������� ��������� "����"
          if CleanAnswer[18] = $02 then itemValues[186 + FCurrentDev] := 1 else itemValues[186 + FCurrentDev] := 0;
          //�������� ��������� "+�������"
          if CleanAnswer[18] = $00 then
            begin
              FBatchMode[FCurrentDev] := False;
              FState[FCurrentDev] := stSenErr;
            end else FBatchMode[FCurrentDev] := True;
        end;
    end;
  if FState[FCurrentDev] = stOk then itemValues[155 + FCurrentDev] := 0;
  if FState[FCurrentDev] = stSenErr then itemValues[155 + FCurrentDev] := 1;
  if FState[FCurrentDev] = stNetErr then itemValues[155 + FCurrentDev] := 2;
  itemValues[248 + FCurrentDev] := PkSt.ErrPaks;
  itemValues[279 + FCurrentDev] := PkSt.LostPaks;
  Alive := True;
  RequestTimer.Enabled := True;
end;

procedure TMicroNetOPC.PacketLost(Sender : TObject);
var
  PkSt : ^PackagesState;
begin
  if Comm1.InQueCount <> 0 then Comm1.PurgeIn;
  PkSt := @FPackages[FCurrentDev];
  PkSt.LostPaks := PkSt.LostPaks + 1;
  PkSt.AllErr := PkSt.AllErr + 1;
  if PkSt.AllErr > 10 then  //������ ����� � ��������
    begin
      FState[FCurrentDev] := stNetErr;
      if FAsyncCmd[FCurrentDev] <> 0 then
        begin
          FAsyncCmd[FCurrentDev] := 0;       //����� ����� ����������� ������� (�� ���������)
        end;
    end;
  if FState[FCurrentDev] = stOk then itemValues[155 + FCurrentDev] := 0;
  if FState[FCurrentDev] = stSenErr then itemValues[155 + FCurrentDev] := 1;
  if FState[FCurrentDev] = stNetErr then itemValues[155 + FCurrentDev] := 2;
  if FState[FCurrentDev] = stNotMnt then itemValues[155 + FCurrentDev] := 3;
  itemValues[279 + FCurrentDev] := PkSt.LostPaks;
  RequestTimer.Enabled := True;     //�������� ������ �������
  WatchDogTimer.Enabled := False;   //��������� ������ �������� ������
end;

procedure TMicroNetOPC.CloseLogFile;
begin
  Recordation := False;
  Writeln(FLogFile, '���������� ������: ' + FormatDateTime('dd/mmmm/yyyy hh:nn:ss', Now));
  CloseFile(FLogFile);                      //������� ���������� ����
end;

procedure TMicroNetOPC.SetAsyncCmd(Index : byte; Value : word);
begin
  FAsyncCmd[Index] := Value;
end;

procedure TMicroNetOPC.SetCalibr(Index : byte; Value : byte);
begin
  if itemValues[Index + 186] = 0 then FCalibr[Index] := Value;
end;

function TMicroNetOPC.GetAverage(Values : AverageStructPointer) : single;
var
  SortCounter : byte;
  SortTemp : single;
  i, ArrayLength : byte;
  TempVar : single;
  TempArray : array of single;
begin
  ArrayLength := High(pAverageStruct.Prod);
  if pAverageStruct.Count < ArrayLength then
    begin
      TempVar := 0;
      for i := 1 to pAverageStruct.Count do TempVar := TempVar + pAverageStruct.Prod[i];
      Result := TempVar / pAverageStruct.Count;
    end else
    begin
      SetLength(TempArray, ArrayLength + 1);
      for i := 1 to ArrayLength do TempArray[i] := pAverageStruct.Prod[i];
      repeat //����������
      SortCounter := 0;
      for i := 1 to ArrayLength - 1 do
        begin
          if TempArray[i] > TempArray[i + 1] then
            begin
              SortTemp := TempArray[i];
              TempArray[i] := TempArray[i + 1];
              TempArray[i + 1] := SortTemp;
              SortCounter := SortCounter + 1;
            end;
        end;
      until SortCounter = 0;
      Result := (TempArray[ArrayLength div 2 - 2] + TempArray[ArrayLength div 2 - 1] + 
        TempArray[ArrayLength div 2] + TempArray[ArrayLength div 2 + 1] + 
        TempArray[ArrayLength div 2 + 2]) / 5;
    end;
end;

end.
