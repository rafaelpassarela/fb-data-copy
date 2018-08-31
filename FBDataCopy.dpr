program FBDataCopy;



uses
  Forms,
  uFormDataCopy in 'uFormDataCopy.pas' {FormDataCopy},
  uFormStatus in 'uFormStatus.pas' {FormStatus},
  uModel in 'uModel.pas',
  uFormSQL in 'uFormSQL.pas' {FormSQL},
  uAdminPriv in 'uAdminPriv.pas',
  uThreadAtualizacao in 'uThreadAtualizacao.pas',
  uMapeamentoDLLCrypto in '..\..\dlls\common\mapeamento\uMapeamentoDLLCrypto.pas',
  LogHook in '..\..\common\util\LogHook.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'DataCopy';

  {$WARN SYMBOL_PLATFORM OFF}
  if DebugHook > 0 then
    ReportMemoryLeaksOnShutdown := True;
  {$WARN SYMBOL_PLATFORM ON}

  Application.CreateForm(TFormDataCopy, FormDataCopy);
  Application.Run;
end.
