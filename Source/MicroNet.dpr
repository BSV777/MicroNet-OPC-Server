program MicroNet;

uses
  Forms,
  Main in 'Main.pas' {MicroNetOPCForm},
  MicroNet_TLB in 'MicroNet_TLB.pas',
  ServIMPL in 'ServIMPL.pas',
  RegDeRegServer in 'RegDeRegServer.pas',
  ComCat in 'comcat.pas',
  Enumstring in 'Enumstring.pas',
  GroupUnit in 'GroupUnit.pas',
  Globals in 'Globals.pas',
  ItemsUnit in 'ItemsUnit.pas',
  ItemAttributesOPC in 'ItemAttributesOPC.pas',
  EnumItemAtt in 'EnumItemAtt.pas',
  AsyncUnit in 'AsyncUnit.pas',
  OPCErrorStrings in 'OPCErrorStrings.pas',
  EnumUnknown in 'EnumUnknown.pas',
  OPCCOMN,
  CommPort in 'CommPort.pas',
  MicroNetOPC in 'MicroNetOPC.pas',
  SingleInst in '..\..\Share\SingleInst.pas',
  OPCDA in '..\..\Share\OPCDA.pas',
  OPCtypes in '..\..\Share\OPCtypes.pas';

{$R *.TLB}

{$R *.RES}

begin
  if not ActivatePrevInstance(TMicroNetOPCForm.ClassName, '') then
    begin
      Application.Initialize;
      Application.Title := 'MicroNet OPC-Server';
      Application.ShowMainForm := False;
      Application.CreateForm(TMicroNetOPCForm, MicroNetOPCForm);
  Application.Run;
    end;
end.
program MicroNet;

uses
  Forms,
  Main in 'Main.pas' {MicroNetOPCForm},
  MicroNet_TLB in 'MicroNet_TLB.pas',
  ServIMPL in 'ServIMPL.pas',
  RegDeRegServer in 'RegDeRegServer.pas',
  ComCat in 'comcat.pas',
  Enumstring in 'Enumstring.pas',
  GroupUnit in 'GroupUnit.pas',
  Globals in 'Globals.pas',
  ItemsUnit in 'ItemsUnit.pas',
  ItemAttributesOPC in 'ItemAttributesOPC.pas',
  EnumItemAtt in 'EnumItemAtt.pas',
  AsyncUnit in 'AsyncUnit.pas',
  OPCErrorStrings in 'OPCErrorStrings.pas',
  EnumUnknown in 'EnumUnknown.pas',
  OPCCOMN,
  CommPort in 'CommPort.pas',
  MicroNetOPC in 'MicroNetOPC.pas',
  SingleInst in '..\..\Share\SingleInst.pas',
  OPCDA in '..\..\Share\OPCDA.pas',
  OPCtypes in '..\..\Share\OPCtypes.pas';

{$R *.TLB}

{$R *.RES}

begin
  if not ActivatePrevInstance(TMicroNetOPCForm.ClassName, '') then
    begin
      Application.Initialize;
      Application.Title := 'MicroNet OPC-Server';
      Application.ShowMainForm := False;
      Application.CreateForm(TMicroNetOPCForm, MicroNetOPCForm);
  Application.Run;
    end;
end.
