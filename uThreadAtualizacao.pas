unit uThreadAtualizacao;

interface

uses
  Classes, Global, SysUtils, XMLIntf, XMLDoc, Variants, Forms, Windows, ActiveX,
  uAdminPriv;

type
  TThreadAtualizacao = class(TThread)
  private
    FOnUpdateReady: TNotifyEvent;
    FVersaoAtual: string;
    function ModoDebug : Boolean;
    function GetCNPJ : string;
    function DoVerificarAtualizacao : Boolean;
    procedure CreateJsonFile(const AFileName : string);
    procedure CreateStopBatFile(const AFileName : string);
  public
    constructor Create; reintroduce;
    procedure Execute; override;

    function IsRunning : Boolean;

    property VersaoAtual : string read FVersaoAtual write FVersaoAtual;
    property OnUpdateReady : TNotifyEvent read FOnUpdateReady write FOnUpdateReady;
  end;


implementation

{ TThreadAtualizacao }

constructor TThreadAtualizacao.Create;
begin
  inherited Create(True);
  FreeOnTerminate := False;
end;

procedure TThreadAtualizacao.CreateJsonFile(const AFileName: string);
var
  lStringList : TStringList;
  lPos: Integer;
begin
  lStringList := TStringList.Create;

  // 1.2.3.4
  lPos := LastDelimiter('.', FVersaoAtual);
  FVersaoAtual[lPos] := '+';

  try
    lStringList.Add('{');
    lStringList.Add('	"id" : "fbdatacopy",');
    lStringList.Add('	"version" : "' + FVersaoAtual + '",');
    lStringList.Add('	"ignore" : [],');
    lStringList.Add('	"clean" : false,');
    lStringList.Add('	"scripts" : {');
    lStringList.Add('		"install" : "",');
    lStringList.Add('		"stop" : "prepare_update.bat",');
    lStringList.Add('		"update" : ""');
    lStringList.Add('	}');
    lStringList.Add('}');

    lStringList.SaveToFile(AFileName);
  finally
    FreeAndNil(lStringList);
  end;
end;

procedure TThreadAtualizacao.CreateStopBatFile(const AFileName: string);
var
  lStringList : TStringList;
begin
  lStringList := TStringList.Create;

  try
    lStringList.Add('@echo off');
    lStringList.Add('if exist ".\FBDataCopy.old" (');
    lStringList.Add('    del .\FBDataCopy.old');
    lStringList.Add(')');

    lStringList.SaveToFile(AFileName);
  finally
    FreeAndNil(lStringList);
  end;
end;

function TThreadAtualizacao.DoVerificarAtualizacao: Boolean;
var
  lDirApp: string;
  lCNPJ: string;
  lJanela: TExecState;
begin
  if ModoDebug then
    lJanela := esNormal
  else
    lJanela := esHidden;

  lCNPJ := Trim(GetCNPJ);
  if lCNPJ <> EmptyStr then
    lCNPJ := ' -cliente ' + lCNPJ;

  lDirApp := IncludeTrailingPathDelimiter( ExtractFilePath(ParamStr(0)) );
  if not FileExists(lDirApp + '\jetpack.json') then
  begin
    CreateJsonFile(lDirApp + '\jetpack.json');
    CreateStopBatFile(lDirApp + '\prepare_update.bat');
  end else
    FileExecuteWait(Global.RetornaCaminhoInstalacao + 'SpiGet\spiget.exe', ' update ' + lCNPJ, lDirApp, lJanela);

  lDirApp := StringReplace(Application.ExeName, '.exe', '.new', [rfIgnoreCase]);
  Result := FileExists(lDirApp);
end;

procedure TThreadAtualizacao.Execute;
var
  lCount: Integer;
begin
  inherited;
  lCount := 0;

  if not TAdminPriv.IsAdmin then
    Terminate
  else begin
    while not Terminated do
    begin
      Inc(lCount);
      if (lCount >= 5) then
      begin
        if DoVerificarAtualizacao and Assigned(FOnUpdateReady) then
          FOnUpdateReady(nil);
        Terminate;
      end;
      Sleep(1000);
    end;
  end;
end;

function TThreadAtualizacao.GetCNPJ: string;
var
  lFileNameRes : TFileNameResolver;
  lConfig: string;
  lXmlDoc: IXMLDocument;
  lXmlNode: IXMLNode;
begin
  CoInitialize(nil);
  lFileNameRes := TFileNameResolver.Create;
  try
    lConfig := lFileNameRes.GetConfigFileName('SpiServiceControl.xml');

    lXmlDoc := TXMLDocument.Create(nil);
    lXmlDoc.LoadFromFile(lConfig);

    lXmlNode := lXmlDoc.ChildNodes.FindNode('FileControl').ChildNodes.FindNode('Geral');
    Result := VarToStrDef( lXmlNode['CNPJEmpresaLocal'], '' );
  finally
    FreeAndNil(lFileNameRes);
    CoUninitialize;
  end;
end;

function TThreadAtualizacao.IsRunning: Boolean;
begin
  Result := not Self.Terminated;
end;

function TThreadAtualizacao.ModoDebug: Boolean;
var
  i: Integer;
begin
  Result := False;

  for i := 1 to ParamCount do
    if ParamStr(i) = '-debug' then
    begin
      Result := True;
      Break;
    end;
end;

end.
