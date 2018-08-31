//{$D-}
unit LogHook;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  SysUtils, Forms, IniFiles, Windows, Classes, Graphics, Jpeg,
  JclDebug, JclHookExcept, ExtCtrls, JclSysInfo, SyncObjs;

type
  TWindowsVersion = class
  public
    class function GetWindowsVersionName : string;
    class function GetProgramVersion : string;
  end;

  TUserInfo = packed record
    UserName : string;
    IP : string;
    AppParams : string;
    ComputerName : string;
    SistemaOperacional : string;
    Versao: string;
    procedure Load;
    function GetInfo : string;
  end;

  TLogConfig = class
  private
    FActive: Boolean;
    FIgnoredTypes: string;
    FIgnoredMessages: string;
    FIgnoredMensList: TStringList;
    FMaxLogSize: Integer;
    FLifeTime: Integer;
    FDirLog: string;
    FTratandoErro : Boolean;
    FErroDetectado : Boolean;
    FNomeIniConfig : TFileName;
    FDirProgramData: string;
//    FServerURL: string;
    procedure Load(const AFirstLoading : Boolean = True);
  public
    constructor Create;
    destructor Destroy; override;

    property DirLog : string read FDirLog;
    property DirProgramData : string read FDirProgramData;

    property Active : Boolean read FActive;
    property IgnoredTypes : string read FIgnoredTypes;
    property IgnoredMessages : string read FIgnoredMessages write FIgnoredMessages;
    property MaxLogSize : Integer read FMaxLogSize;
    property LifeTime : Integer read FLifeTime write FLifeTime;

    function GetNextExceptionNumber : Integer;
    procedure ReLoadConfig;
  end;

  TLogControl = class
  private
    FUserInfo : TUserInfo;
    FConfig: TLogConfig;
    FLastStackTrace: TStringList;
    FLastLogValid: Boolean;
    FInCriticalArea: Boolean;
    procedure Logar(Message: string; const NomeArquivo : String;
      const AAdicionarNroException: Boolean = True; const ALineBreak : Boolean = True);
    procedure ClearErrorDetected;
    function FileSize(const aFilename: String): Int64;
    function GetLogBkpFileName(const AFileName : string) : string;
  public
    constructor Create;
    destructor Destroy; override;

    property UserInfo : TUserInfo read FUserInfo;
    property Config : TLogConfig read FConfig;
    property LastStackTrace : TStringList read FLastStackTrace;

    property LastLogValid : Boolean read FLastLogValid;

    function CanLog(const AErrorClassName, AErrorMessage : string) : Boolean;

    procedure WriteLog(const AMessage : string; const AFileName : string; const ALineBreak : Boolean = True);

    procedure EnterCriticalArea;
    procedure LeaveCriticalArea;
  end;

var
  _LogControl: TLogControl;
  _LogHookCS: TCriticalSection;

function GetLogControl : TLogControl;

// define se o log será capturado ou não, quando o parametro Active = true, mas em
// em código que o erro é esperado
procedure EnterCriticalArea;
procedure LeaveCriticalArea;

function LogAtivo: Boolean;
function GetLogDir: string;
function GetLogFileName: string;

implementation

const
  ConstDirLog = 'Logs\';
  ConstDefaultIgnoreList = ';EResNotFound;EAbort;EDBClient;EExternalException;EIB_ISCError;EDatabaseError;EIdSocketError;EIdConnClosedGracefully;'; // ;Exception;
  ConstSection_LogConfig = 'LogConfig';
  ConstSection_Log = 'Log';

  ConstIdent_Active         = 'Active';
  ConstIdent_IgnoreList     = 'IgnoreLogTypes';
  ConstIdent_IgnoreMessages = 'IgnoreLogMsges';
  ConstIdent_LogSize        = 'LogSize';
  ConstIdent_Count          = 'LogCount';
  ConstIdent_TratandoErro   = 'TratandoErro';
  ConstIdent_ErrorDetected  = 'ErrorDetected';
  ConstIdent_LogLifeTime    = 'LogLifeTime';

//  ConstErroHostFB = 'Unable to complete network request to host';
//  ConstErroTCPClient = '<<>>Connection Closed Gracefully';

function GetLogControl : TLogControl;
begin
  Result := _LogControl;
end;  

procedure EnterCriticalArea;
begin
  _LogControl.EnterCriticalArea;
end;

procedure LeaveCriticalArea;
begin
  _LogControl.LeaveCriticalArea;
end;

function LogAtivo : Boolean;
begin
  Result := _LogControl.Config.Active;
end;

function GetLogDir: string;
begin
  Result := IncludeTrailingPathDelimiter(_LogControl.Config.DirLog);
end;

function GetLogFileName : string;
begin
  Result := _LogControl.Config.DirLog
          + StringReplace(ExtractFileName(Application.ExeName), '.exe', '', [rfIgnoreCase])
          + FormatDateTime('_YYYY-MM-DD_"E"rror', Date) + '.log';
end; 

// Função principal para captura do stacktrace
procedure LogExceptionHook(ExceptObj: TObject; ExceptAddr: Pointer; OSException: Boolean);
var
  lNomeArquivo : String;

begin
  if (not _LogControl.Config.FTratandoErro)
  and (_LogControl.Config.Active)
  and _LogControl.CanLog(Exception(ExceptObj).ClassName, Exception(ExceptObj).Message) then
  begin
    try
      _LogControl.LastStackTrace.BeginUpdate;
      _LogControl.LastStackTrace.Clear;
      _LogControl.LastStackTrace.Add(Format('''%s'': %s.' + sLineBreak + '%s', [
        Exception(ExceptObj).ClassName,
        Exception(ExceptObj).Message,
        _LogControl.UserInfo.GetInfo]));
      JclLastExceptStackListToStrings(_LogControl.LastStackTrace, True);

      ForceDirectories(_LogControl.Config.DirLog);

      lNomeArquivo := GetLogFileName;
      _LogControl.Logar(_LogControl.LastStackTrace.Text, lNomeArquivo);
      _LogControl.Config.FErroDetectado := True;
    finally
      _LogControl.LastStackTrace.EndUpdate;
    end;
  end;
end;

{ TUserInfo }

function TUserInfo.GetInfo: string;
begin
  Result := Format('Usuário....: %s' + sLineBreak
                 + 'Endereço IP: %s' + sLineBreak
                 + 'OS.........: %s' + sLineBreak
                 + 'Parâmetros.: %s' + sLineBreak
                 + 'Versão.....: %s' + sLineBreak
                 + 'Data e Hora: %s', [
                   UserName,
                   ComputerName,
                   SistemaOperacional,
                   AppParams,
                   Versao,
                   FormatDateTime('dd/mm/yyyy hh:nn:ss:zzz', Now) ]);

  Result := StringOfChar('-', 80) + sLineBreak + Result
          + sLineBreak + StringOfChar('-', 80);
end;

procedure TUserInfo.Load;
var Ind : Integer;
begin
  AppParams := '';
  for Ind := 0 to ParamCount do
    AppParams := AppParams + '"' + ParamStr(Ind) + '"  ';
  AppParams := Trim(AppParams);

  UserName := GetLocalUserName;
  ComputerName := GetLocalComputerName;
  IP := GetIPAddress(ComputerName);

  SistemaOperacional := TWindowsVersion.GetWindowsVersionName;

  Versao := TWindowsVersion.GetProgramVersion;
end;

{ TLogControl }

function TLogControl.CanLog(const AErrorClassName, AErrorMessage: string): Boolean;
var
  i: Integer;
  lMens: string;
begin
  // testa a classe do erro
  Result := (not FInCriticalArea)
        and (Pos(';' + AErrorClassName + ';', FConfig.IgnoredTypes) <= 0);

//    Result := (Pos(ConstErroHostFB, AErrorMessage) <= 0)
//          and (Pos(ConstErroTCPClient, AErrorMessage) <= 0);
  // testa a mensagem do erro
  if Result then
  begin
    for i := 0 to FConfig.FIgnoredMensList.Count - 1 do
    begin
      lMens := Trim(FConfig.FIgnoredMensList[i]);
      if (lMens <> EmptyStr) and (Pos(lMens, AErrorMessage) > 0) then
      begin
        Result := False;
        Break;
      end;
    end;
  end;
        
  FLastLogValid := Result;
end;

procedure TLogControl.ClearErrorDetected;
var
  lLista : TStringList;

  procedure FindLogFiles(const AMask : string);
  var
    lRec: TSearchRec;
  begin
    if FindFirst(FConfig.DirLog + AMask, faAnyFile, lRec) = 0 then
    begin
      repeat
        if FileExists(FConfig.DirLog + lRec.Name) then
          lLista.Add(FConfig.DirLog + lRec.Name );
      until FindNext(lRec) <> 0;
    end;
    SysUtils.FindClose(lRec);
  end;

  procedure ApagaArquivoLog;
  var
    DataArq : TDateTime;
    lFileAtr: TWin32FileAttributeData;
    SystemTime, LocalTime: TSystemTime;
    i : Integer;
  begin
    for i := 0 to lLista.Count - 1 do
    begin
      if GetFileAttributesEx(PChar(lLista.Strings[i]), GetFileExInfoStandard, @lFileAtr) then
      begin
        if FileTimeToSystemTime(lFileAtr.ftCreationTime, SystemTime) and SystemTimeToTzSpecificLocalTime(nil, SystemTime, LocalTime) then
        begin
          DataArq := SystemTimeToDateTime(LocalTime);
          if DataArq < (Now - FConfig.LifeTime) then
            DeleteFile(PChar(lLista.Strings[i]));
        end
        else
          DeleteFile( PChar(lLista.Strings[i]) )
      end
      else
        DeleteFile( PChar(lLista.Strings[i]) );
    end;
  end;

begin
  FConfig.FTratandoErro := False;
  FConfig.FErroDetectado := False;

  if FConfig.LifeTime > 0 then
  try
    lLista := TStringList.Create;
    { logs de erros }
    FindLogFiles('*.log');
    FindLogFiles('*.bkp*');

    ApagaArquivoLog;
  finally
    FreeAndNil(lLista);
  end;
end;

constructor TLogControl.Create;
begin
  inherited;
  FLastLogValid := False;
  FLastStackTrace := TStringList.Create;
  FUserInfo.Load;
  FConfig := TLogConfig.Create;
end;

destructor TLogControl.Destroy;
begin
  FreeAndNil(FLastStackTrace);
  FreeAndNil(FConfig);
  inherited;
end;

procedure TLogControl.EnterCriticalArea;
begin
  FInCriticalArea := True;
end;

function TLogControl.FileSize(const aFilename: String): Int64;
var
  lInfo: TWin32FileAttributeData;
  lOk: Boolean;
begin
  Result := -1;
  {$IFDEF CONSOLE}
  lOk := GetFileAttributesEx(PAnsiChar(aFileName), GetFileExInfoStandard, @lInfo);
  {$ELSE}
    {$IFDEF DLLMODE}
    lOk := GetFileAttributesEx(PAnsiChar(aFileName), GetFileExInfoStandard, @lInfo);
    {$ELSE}
      {$IFNDEF UNICODE}
      lOk := GetFileAttributesEx(PAnsiChar(aFileName), GetFileExInfoStandard, @lInfo);
      {$ELSE}
      lOk := GetFileAttributesEx(PWideChar(aFileName), GetFileExInfoStandard, @lInfo);
      {$ENDIF}
    {$ENDIF}
  {$ENDIF}

  if not lOk then
    Exit;

  Result := lInfo.nFileSizeLow or (lInfo.nFileSizeHigh shl 32);
end;

function TLogControl.GetLogBkpFileName(const AFileName: string): string;
var
  lNomeBase: string;
  lCount: Integer;
  lRec: TSearchRec;
begin
  lNomeBase := StringReplace(AFileName, '.log', '', [rfIgnoreCase]);
  lCount := 0;
  // renomeia o arquivo para o próximo arquivo de backup
  if FindFirst(lNomeBase + '.bkp*', faAnyFile, lRec) = 0 then
  begin
    repeat
      Inc(lCount);
    until FindNext(lRec) <> 0;
  end;
  Inc(lCount);

  Result := lNomeBase + '.bkp' + IntToStr(lCount);
end;

procedure TLogControl.LeaveCriticalArea;
begin
  FInCriticalArea := False;
end;

procedure TLogControl.Logar(Message: string; const NomeArquivo: String;
  const AAdicionarNroException: Boolean; const ALineBreak : Boolean);
var
  LogFile: Text;
  lFileOpen: Boolean;
  lControl: Integer;
  lRealLogName: string;
begin
  _LogHookCS.Enter;
  try
    lFileOpen := False;
    try
      // garante que os arquivos fiquem na mesma pasta, mesmo quando o nome é
      // informado incorretamente para outra pasta
      lRealLogName := FConfig.DirLog + ExtractFileName(NomeArquivo);
      ForceDirectories(FConfig.DirLog);

      if FileExists(lRealLogName) then
      begin
        AssignFile(LogFile, lRealLogName);
        if FileSize(lRealLogName) > (1024 * FConfig.MaxLogSize) then
        begin
          CopyFile(PChar(lRealLogName), PChar( GetLogBkpFileName(lRealLogName) ), False );
          DeleteFile(PChar(lRealLogName));
          ReWrite(LogFile);
          lFileOpen := True;
        end else
        begin
          // alguem pode estar utilizando o log neste momento
          lControl := 0;
          repeat
            try
              Inc(lControl);
              Append(LogFile);
              lFileOpen := True;
            except
              OutputDebugString(PChar('Arquivo de Log em Uso [' + lRealLogName + ']'));
              Sleep(500);
            end;
          until lFileOpen or (lControl > 120); //max 1min tentando liberar o arquivo
        end;
      end
      else
      begin
        AssignFile(LogFile, lRealLogName);
        ReWrite(LogFile);
        lFileOpen := True;
      end;

      if AAdicionarNroException then
        Writeln(LogFile, Format('Exceção N°: %d)', [_LogControl.Config.GetNextExceptionNumber]));

      if ALineBreak then
        WriteLn(LogFile, Message)
      else
        Write(LogFile, Message);
    finally
      if lFileOpen then
        CloseFile(LogFile);
    end;
  finally
    _LogHookCS.Leave;
  end;
  
  if DebugHook = 1 then
    OutputDebugString( PChar(Message) );
end;

procedure TLogControl.WriteLog(const AMessage, AFileName: string; const ALineBreak : Boolean);
begin
  Logar(AMessage, AFileName, False, ALineBreak);
end;

{ TLogConfig }

constructor TLogConfig.Create;
begin
  FDirProgramData := EmptyStr;
  FIgnoredMensList := TStringList.Create;
  FIgnoredMensList.StrictDelimiter := True;
  FIgnoredMensList.Delimiter := ';';

  Load;
  
  FErroDetectado := False;
  FTratandoErro := False;
end;

destructor TLogConfig.Destroy;
begin
  FreeAndNil(FIgnoredMensList);
  
  inherited;
end;

function TLogConfig.GetNextExceptionNumber: Integer;
var
  lIni : TIniFile;
begin
  lIni := TIniFile.Create(FNomeIniConfig);
  try
    Result := lIni.ReadInteger(ConstSection_LogConfig, ConstIdent_Count, 0) + 1;
    lIni.WriteInteger(ConstSection_LogConfig, ConstIdent_Count, Result);
  finally
    FreeAndNil(lIni);
  end;
end;

procedure TLogConfig.Load(const AFirstLoading : Boolean);
var
  lConfigOk: Boolean;
  i: Integer;

  procedure FindBestPath(const APaths : array of string);
  var
    i: Integer;
  begin
    for i := 0 to High(APaths) do
      if DirectoryExists(APaths[i]) then
      begin
        FDirLog := APaths[i];
        Break;
      end;
  end;

begin
  if AFirstLoading or (FDirProgramData = EmptyStr) then
  begin
    if GetCommonAppdataFolder <> '' then
      FDirProgramData := IncludeTrailingPathDelimiter(GetCommonAppdataFolder) + 'SoftPlus\'
    else
      FDirProgramData := IncludeTrailingPathDelimiter(GetWindowsSystemFolder);
  end;

  with TIniFile.Create(FDirProgramData + 'softfire.ini') do
  try
    FDirLog := ReadString('Service', 'RootDir', EmptyStr);

    if FDirLog = EmptyStr then
      FindBestPath(['C:\Program Files (x86)\Softplus',
                    'C:\Arquivos de Programas (x86)\Softplus',
                    'C:\Program Files\Softplus',
                    'C:\Arquivos de Programas\Softplus']);

    lConfigOk := FDirLog <> EmptyStr;

    if not lConfigOk then
    begin
      FDirLog := SysUtils.GetEnvironmentVariable('ProgramFiles(x86)');
      if (FDirLog <> EmptyStr) and (DirectoryExists(FDirLog)) then
        FDirLog := IncludeTrailingPathDelimiter(FDirLog)
      else
        FDirLog := IncludeTrailingPathDelimiter(SysUtils.GetEnvironmentVariable('ProgramFiles'));

      FDirLog := FDirLog + 'SoftPlus';
      lConfigOk := True;
    end;
  finally
    Free;
  end;

  if lConfigOk then
  begin
    FDirLog := IncludeTrailingPathDelimiter(FDirLog);
    FNomeIniConfig := FDirLog + 'LogConfigV4.ini';
    FDirLog := FDirLog + ConstDirLog;

    if AFirstLoading then
      ForceDirectories(FDirLog);

    with TIniFile.Create( FNomeIniConfig ) do
    try
      // grava os valores default de configuracao
      if AFirstLoading and (not FileExists(FNomeIniConfig)) then
      begin
        WriteBool(ConstSection_LogConfig,    ConstIdent_Active,      False);
        WriteString(ConstSection_LogConfig,  ConstIdent_IgnoreList,  ConstDefaultIgnoreList);
        WriteInteger(ConstSection_LogConfig, ConstIdent_LogSize,     1024);
        WriteInteger(ConstSection_LogConfig, ConstIdent_LogLifeTime, 7);

        WriteBool(ConstSection_Log, ConstIdent_TratandoErro,  False);
        WriteBool(ConstSection_Log, ConstIdent_ErrorDetected, False);
      end;

      FActive          := ReadBool(ConstSection_LogConfig,    ConstIdent_Active,      False);
      FIgnoredTypes    := ReadString(ConstSection_LogConfig,  ConstIdent_IgnoreList,  ConstDefaultIgnoreList);
      FIgnoredMessages := ReadString(ConstSection_LogConfig,  ConstIdent_IgnoreMessages, '');

      FIgnoredMensList.Clear;
      FIgnoredMensList.DelimitedText := FIgnoredMessages;

      for i := FIgnoredMensList.Count - 1 downto 0 do
        if FIgnoredMensList[i] = EmptyStr then
          FIgnoredMensList.Delete(i);

      if FIgnoredMensList.Count = 0 then
        FIgnoredMensList.DelimitedText := 'Unable to complete network request to host;<<>>Connection Closed Gracefully';

      // garante que não se exclua todas as Exceptions
      if AFirstLoading and (Pos(';Exception;', FIgnoredTypes) > 0) or (Pos('Exception; ', FIgnoredTypes) > 0) then
      begin
        WriteString(ConstSection_LogConfig,  ConstIdent_IgnoreList, ConstDefaultIgnoreList);
        FIgnoredTypes := ConstDefaultIgnoreList;
      end;
      FIgnoredTypes := FIgnoredTypes + ';ExceptionNoInfo;';

      FMaxLogSize    := ReadInteger(ConstSection_LogConfig, ConstIdent_LogSize,     1024);
      FLifeTime      := ReadInteger(ConstSection_LogConfig, ConstIdent_LogLifeTime, 7);

      if (FMaxLogSize <= 0) or (FMaxLogSize > 10240) then // 10Mb
        FMaxLogSize := 1024;
    finally
      Free;
    end;
  end else
    FActive := False;
end;

procedure TLogConfig.ReLoadConfig;
begin
  Load(False);
end;

{ TWindowsVersion }

class function TWindowsVersion.GetProgramVersion: string;
var
  v1, v2, v3, v4: Word;
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  Result := EmptyStr;

  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  if VerInfoSize > 0 then
  begin
    GetMem(VerInfo, VerInfoSize);
    try
      if GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo) then
      begin
        VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
        with VerValue^ do
        begin
          V1 := dwFileVersionMS shr 16;
          V2 := dwFileVersionMS and $FFFF;
          V3 := dwFileVersionLS shr 16;
          V4 := dwFileVersionLS and $FFFF;
        end;
        Result := Format('%d.%d.%d.%d', [ v1, v2, v3, v4]);
      end;
    finally
      FreeMem(VerInfo, VerInfoSize);
    end;
  end;
end;

class function TWindowsVersion.GetWindowsVersionName: string;
var
  lVersion : TOSVersionInfo;
begin
  lVersion.dwOSVersionInfoSize := SizeOf(lVersion);
  GetVersionEx(lVersion);
  Result := '';
  case lVersion.dwPlatformId of
    1:
      case lVersion.dwMinorVersion of
         0: Result := 'Windows 95';
        10: Result := 'Windows 98';
        90: Result := 'Windows Me';
      end;
    2:
      case lVersion.dwMajorVersion of
        3: Result := 'Windows NT 3.51';
        4: Result := 'Windows NT 4.0';
        5: case lVersion.dwMinorVersion of
             0: Result := 'Windows 2000';
             1: Result := 'Windows XP';
             2: Result := 'Windows Server 2003';
           end;
        6: case lVersion.dwMinorVersion of
             0 : Result := 'Windows Vista';
             1 : Result := 'Windows 7';
             2 : Result := 'Windows 8';
             3 : Result := 'Windows 8.1';
           end;
        10: Result := 'Windows 10';
        else
          Result := 'Windows 10 ou Superior';
      end;
  end;

  if (Result = '') then
    Result := 'Sistema operacional desconhecido.'
  else
    Result := Result + ' ' + Trim(lVersion.szCSDVersion);
end;

initialization
  _LogHookCS := TCriticalSection.Create;
  _LogControl := TLogControl.Create;
  _LogControl.ClearErrorDetected;
//  JclStackTrackingOptions := [stStack, stExceptFrame, stRawMode, stAllModules, stStaticModuleList];
  JclStackTrackingOptions := [stStack, stRawMode];
  JclStartExceptionTracking;
  JclAddExceptNotifier(LogExceptionHook);

finalization
  JclRemoveExceptNotifier(LogExceptionHook);
  FreeAndNil(_LogControl);
  FreeAndNil(_LogHookCS);

end.
