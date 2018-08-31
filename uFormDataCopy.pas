//{$DEFINE TESTE}
unit uFormDataCopy;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit,
  cxButtonEdit, Buttons, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, DB, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView,  cxGridDBTableView, cxGrid, ImgList,
  DBClient, cxImageComboBox, IBODataset, IB_Components, ComCtrls, cxCheckBox,
  cxLookAndFeels, cxLookAndFeelPainters, IB_Access, uFormStatus, IBOExtract,
  IB_Process, IB_Script, Generics.Collections, StrUtils, uModel, uSPITCPBase,
  cxDropDownEdit, Math, cxCalendar, cxBlobEdit, cxMemo, uFormSQL, ShellAPI,
  uAdminPriv;

const
  WM_AFTER_SHOW = WM_USER + $1123;
  ConstSkipCreate : array[0..3] of string = (
    'CREATE DATABASE',
    'DECLARE EXTERNAL FUNCTION RDB$GET_CONTEXT',
    'DECLARE EXTERNAL FUNCTION RDB$SET_CONTEXT',
    'RDB$ADMIN');
  ConstTagSELECT = '{SELECT}';
  ConstTagDATA   = '{DATA}';
  C_CONECTAR = '&Conectar';
  C_DESCONECTAR = 'Des&conectar';
  C_FILE_DESC   = 'Arquivo de Definições (*.xdef)|*.xdef|Arquivo de Definições Antigos (*.def)|*.def';

type
  TTipoStatus = (stRegistro, stErro, stTabela, stReset);

  TFieldIdx = class(TDictionary<string,Integer>)
  end;

  TFormDataCopy = class(TForm)
    LabelDestino: TLabel;
    LabelOrigem: TLabel;
    EditArquivoOrigem: TEdit;
    EditDestino: TEdit;
    BitBtnConectar: TBitBtn;
    ClientDataSetTabelas: TClientDataSet;
    DataSourceTabelas: TDataSource;
    ImageList1: TImageList;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    ClientDataSetTabelasSTATUS: TIntegerField;
    ClientDataSetTabelasNOME: TStringField;
    ClientDataSetTabelasREGISTROS: TIntegerField;
    cxGrid1DBTableView1STATUS: TcxGridDBColumn;
    cxGrid1DBTableView1NOME: TcxGridDBColumn;
    cxGrid1DBTableView1REGISTROS: TcxGridDBColumn;
    IBODatabaseOrigem: TIBODatabase;
    IB_TransactionOrigem: TIB_Transaction;
    IBOQueryTabelas: TIBOQuery;
    IBOQueryRegistros: TIBOQuery;
    BitBtnIniciar: TBitBtn;
    OpenDialogScript: TOpenDialog;
    IBODatabaseDestino: TIBODatabase;
    IB_TransactionDestino: TIB_Transaction;
    IB_DSQLDestino: TIB_DSQL;
    IBOQueryChave: TIBOQuery;
    IBOQueryCampos: TIBOQuery;
    cdsCampos: TClientDataSet;
    cdsCamposNome: TStringField;
    cdsCamposPrimaryKey: TBooleanField;
    MemoLog: TMemo;
    StatusBar1: TStatusBar;
    ClientDataSetTabelasPESO: TIntegerField;
    ClientDataSetTabelasIMPORT: TBooleanField;
    cxGrid1DBTableView1IMPORT: TcxGridDBColumn;
    IBOQueryPesos: TIBOQuery;
    IBOExtract1: TIBOExtract;
    ButtonBanco: TButton;
    IB_ScriptNovo: TIB_Script;
    IBOQueryTriggers: TIBOQuery;
    BitBtnLoad: TBitBtn;
    BitBtnSave: TBitBtn;
    ClientDataSetTabelasCAMPODATA: TStringField;
    cxGrid1DBTableView1CAMPODATA: TcxGridDBColumn;
    ClientDataSetTabelasLISTACAMPODATA: TStringField;
    IBOQueryGeneratorsOrigem: TIBOQuery;
    GroupBoxParametros: TGroupBox;
    LabelDataFiltro: TLabel;
    cxDateEditFiltro: TcxDateEdit;
    CheckBoxGenerators: TCheckBox;
    CheckBoxIgnorar: TCheckBox;
    CheckBoxNovo: TCheckBox;
    CheckBoxDesativaTriggers: TCheckBox;
    cxGrid1DBTableView1CUSTOMSQL: TcxGridDBColumn;
    ClientDataSetTabelasCUSTOMSQL: TStringField;
    ClientDataSetTabelasPAGINACAO: TIntegerField;
    cxGrid1DBTableView1PAGINACAO: TcxGridDBColumn;
    CheckBoxDependencia: TCheckBox;
    LabelVersao: TLabel;
    BitBtnReconectar: TBitBtn;
    IB_QueryDadosOrigem: TIB_Query;
    IBOQueryGeneratorsDestino: TIBOQuery;
    procedure FormCreate(Sender: TObject);
    procedure ClientDataSetTabelasNewRecord(DataSet: TDataSet);
    procedure BitBtnConectarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BitBtnIniciarClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ButtonBancoClick(Sender: TObject);
    procedure IB_ScriptNovoStatement(Sender: TIB_Script; var Statement: string;
      var SkipIt: Boolean);
    procedure BitBtnSaveClick(Sender: TObject);
    procedure BitBtnLoadClick(Sender: TObject);
    procedure cxGrid1DBTableView1IMPORTHeaderClick(Sender: TObject);
    procedure ClientDataSetTabelasAfterScroll(DataSet: TDataSet);
    procedure cxGrid1DBTableView1CUSTOMSQLPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure EditArquivoOrigemEnter(Sender: TObject);
    procedure EditArquivoOrigemExit(Sender: TObject);
    procedure BitBtnReconectarClick(Sender: TObject);
  private
    { Private declarations }
    FImportados, FErros, FTabelas : Int64;
    FNomeLog : string;
    FStatus: TFormStatus;
    FAplicarFiltroData: Boolean;
    FConfigFileName: string;
    FTmpConfigFileName: string;
    FConfig: TConfig;

    procedure WmAfterShow(var Dummy : Boolean); message WM_AFTER_SHOW;
    procedure ReOpenApp;
    procedure AddFileSizeToLog(const ANewDBFileName, AOldDBFileName : string);

    procedure MensIgnorarTabela;
    procedure DoSaveConfig(const AFileName : string);
    procedure DoLoadConfig;
    procedure DoSetGenerators;
    procedure ResetTableStatus;
    procedure AddLog(const AMens : string; const AComment : Boolean = True; const ADataHora : Boolean = True);
    procedure DoConnectDatabase;
    procedure DoDesconectarDatabase;
    procedure DoEnableFormForImport(const Value : Boolean);
    procedure LoadVersion;

    function ReConnect(const AQry : TComponent; const AException: string; const ARestorePos : Boolean = True) : Boolean;
    function DoReconnect : Boolean;
    function IsSaveEnabled : Boolean;

    function FileSize(const aFilename: String): Int64;

    function GetWindowsDir : string;
    function GetTableRecordCount(const ANomeTabela : string) : Integer;
    function GetTableRecordCountEx(const ANomeTabela : string) : Integer;
    function GetTableDateList(const ANomeTabela : string) : string;
    function CriaNovoBanco(const NomeBanco : TFileName) : Boolean;
    function ProcessarRegistrosTabela(const ATotalRegistros : Int64;
      const ANomeTabela : string) : Boolean;
    function ImportaTabela(const ANomeTabela : string; const ATotalRegistros : Integer;
      const AIsDependencia : Boolean) : Boolean;

    function GetInsertSQL(const NomeTabela : string) : string;
    function GetInsertUpdateSQL(const NomeTabela : string) : string;
    procedure AtualizaStatus(const ATipo : TTipoStatus);
    procedure SetTriggersOnOff(const AValue : Boolean);

    procedure MostraErro(const AMens : string; const AAbort : Boolean);
    procedure MostraPerigo(const AMens : string);
    function MostraPergunta(const APergunta : string) : TModalResult;
  public
    { Public declarations }
  end;

var
  FormDataCopy: TFormDataCopy;

implementation

{$R *.dfm}

procedure TFormDataCopy.AddFileSizeToLog(const ANewDBFileName, AOldDBFileName: string);
var
  lDbSizeAfter: Int64;
  lDbSizeBefore: Int64;
  lPercent: Extended;

  function SizeToStr(const ASize : Int64) : string;
  const
    C_BASE = 1024;
  var
    lCalc: Extended;
  begin
    if ASize < C_BASE then
      Result := FormatFloat('#0.000 bytes', ASize)
    else begin
      lCalc := ASize / 1024;
      if lCalc < C_BASE then
        Result := FormatFloat('#0.000 KB', lCalc)
      else begin
        lCalc := lCalc / 1024;
        if lCalc < C_BASE then
          Result := FormatFloat('#0.000 MB', lCalc)
        else begin
          lCalc := lCalc / 1024;
          if lCalc < C_BASE then
            Result := FormatFloat('#0.000 GB', lCalc)
          else
            Result := FormatFloat('#0.000 TB', lCalc / 1024)
        end;
      end;
    end;
  end;

begin
  lDbSizeAfter := FileSize(ANewDBFileName);
  lDbSizeBefore:= FileSize(AOldDBFileName);

  if (lDbSizeAfter <> -1) and (lDbSizeBefore <> -1) then
  begin
    lPercent := (lDbSizeAfter * 100) / lDbSizeBefore;
    lPercent := 100 - lPercent;

    AddLog('O Banco de Dados ficou ' + FormatFloat('#0.000', lPercent) + '% menor.');
    AddLog('Banco Original: ' + SizeToStr(lDbSizeBefore));
    AddLog('Banco Novo: ' + SizeToStr(lDbSizeAfter));
  end;
end;

procedure TFormDataCopy.AddLog(const AMens: string; const AComment, ADataHora: Boolean);
var
  lMens : string;
begin
  lMens := AMens;
  if ADataHora then
    lMens := FormatDateTime('hh:nn:ss:zzz - ', Now) + lMens;

  if AComment then
    lMens := '/* ' + lMens + ' */';

  MemoLog.Lines.Add( lMens );
end;

procedure TFormDataCopy.AtualizaStatus(const ATipo : TTipoStatus);
var
  i: Integer;
begin
  case ATipo of
    stRegistro: StatusBar1.Panels[0].Text := 'Total de Registros Importados: ' + IntToStr(FImportados);
    stErro:     StatusBar1.Panels[1].Text := 'Erros de Registro: ' + IntToStr(FErros);
    stTabela:   StatusBar1.Panels[2].Text := 'Tabelas Proc.: ' + IntToStr(FTabelas) + ' de ' + IntToStr(ClientDataSetTabelas.RecordCount);
    stReset:
    begin
      for i := 0 to StatusBar1.Panels.Count -1 do
        StatusBar1.Panels[i].Text := EmptyStr;
    end;
  end;
  StatusBar1.Refresh;
end;

procedure TFormDataCopy.BitBtnConectarClick(Sender: TObject);
begin
  BitBtnConectar.Enabled := False;
  try
    if BitBtnConectar.Caption = C_CONECTAR then
      DoConnectDatabase
    else
      DoDesconectarDatabase;
  finally
    BitBtnConectar.Enabled := True;
  end;
end;

procedure TFormDataCopy.BitBtnIniciarClick(Sender: TObject);
var
  NomeNovoBanco : TFileName;
  CriaBanco : Boolean;
  ExisteBanco : Boolean;
  BancoOk : Boolean;
  UltimoRegistro : string;
  lStart: TDateTime;

  procedure LogarHoraInicio;
  begin
    lStart := Now;
    AddLog('Iniciando Cópia do Banco de Dados', True, True);
  end;

begin
  BitBtnIniciar.Enabled := False;
  BitBtnConectar.Enabled := False;

  AtualizaStatus(stReset);
  MemoLog.Clear;

  NomeNovoBanco :=  ExtractFilePath(EditArquivoOrigem.Text);
  // Log
  FNomeLog := StringReplace(Application.ExeName, '.exe', '.log', [rfIgnoreCase]);
  // Banco
  NomeNovoBanco :=  IncludeTrailingPathDelimiter(NomeNovoBanco) + EditDestino.Text;
  BancoOk := False;
  FErros := 0;
  FImportados := 0;
  ExisteBanco := FileExists(NomeNovoBanco);
  FAplicarFiltroData := (cxDateEditFiltro.Date > StrToDate('01/01/1950'));
  IBOQueryRegistros.Close;

  ResetTableStatus;

  if CheckBoxNovo.Checked then
  begin
    if ExisteBanco then
    begin
      CriaBanco := MostraPergunta('O arquivo já existe, deseja criar um novo?') = idYes;
      if CriaBanco then
      begin
        if DeleteFile(NomeNovoBanco) then
          ExisteBanco := False
        else
          MostraErro('Não foi possivel excluir o arquivo do novo banco de dados.', True);
      end else
      begin
        BitBtnIniciar.Enabled := True;
        BitBtnConectar.Enabled := True;
        Exit;
      end;
    end;

    LogarHoraInicio;

    if (not ExisteBanco) then
    begin
      DoEnableFormForImport(False);
      BancoOk := CriaNovoBanco( NomeNovoBanco );
    end;
  end else
  begin
    if not ExisteBanco then
      MostraErro('Banco de dados de destino não encontrado.', True)
    else begin
      BancoOk := True;
      LogarHoraInicio;
    end;
  end;

  { Conecta banco }
  if BancoOk then
  begin
    DoEnableFormForImport(False);

    if IBODatabaseDestino.Connected then
      IBODatabaseDestino.Disconnect;

    IBODatabaseDestino.Server := 'LOCALHOST';
    IBODatabaseDestino.PageSize := 8192;
    IBODatabaseDestino.Username := 'SYSDBA';
    IBODatabaseDestino.Password := 'masterkey';
    IBODatabaseDestino.DatabaseName := NomeNovoBanco;
    IBODatabaseDestino.Connect;

    if CheckBoxDesativaTriggers.Checked then
      SetTriggersOnOff(False);

    FTabelas := 0;
    if Pos('.def', LowerCase(FConfigFileName)) > 0 then
    begin
      // se o arquivo ja é temporario, remove o indicador do nome
      FTmpConfigFileName := StringReplace(FConfigFileName, '_tmp.def', '.def', [rfIgnoreCase]);
      // atualiza o nome do arquivo para indicar que é temporário
      FTmpConfigFileName := StringReplace(FTmpConfigFileName, '.def', '_tmp.def', [rfIgnoreCase]);
    end else
    begin
      // se o arquivo ja é temporario, remove o indicador do nome
      FTmpConfigFileName := StringReplace(FConfigFileName, '_tmp.xdef', '.xdef', [rfIgnoreCase]);
      // atualiza o nome do arquivo para indicar que é temporário
      FTmpConfigFileName := StringReplace(FTmpConfigFileName, '.xdef', '_tmp.xdef', [rfIgnoreCase]);
    end;

    with ClientDataSetTabelas do
    begin
      First;
      while not Eof do
      begin
        if FieldByName('STATUS').AsInteger <> 0 then
        begin
          MensIgnorarTabela;
          Next;
        end else
        begin
          // atualiza para executando
          UltimoRegistro := FieldByName('NOME').AsString;
          if FieldByName('IMPORT').AsBoolean then
          begin
            if ImportaTabela(UltimoRegistro, FieldByName('REGISTROS').AsInteger, False) then
            begin // Importa certo
              Edit;
              if FieldByName('STATUS').AsInteger <> 4 then
                FieldByName('STATUS').AsInteger := 2;
              Post;
            end
            else
            begin // Erro
              Edit;
              if FieldByName('STATUS').AsInteger <> 4 then
                FieldByName('STATUS').AsInteger := 3;
              Post;
            end;
          end else
          begin
            MensIgnorarTabela;
            Edit;
            FieldByName('STATUS').AsInteger := 4;
            Post;
          end;
          DoSaveConfig(FTmpConfigFileName);

          Inc(FTabelas);
          AtualizaStatus(stTabela);
          // Localiza o ultimo registro chamado localmente
          Locate('NOME', UltimoRegistro, []);
          {$IFDEF TESTE}
          OutputDebugString(PChar('Ultimo = ' + UltimoRegistro));
          {$ENDIF}
          Next;
        end;
      end;
    end;

    if CheckBoxGenerators.Checked then
      DoSetGenerators
    else begin
      MostraPerigo('Os GENERATORS devem ser ajustados manualmente.');
      AddLog('Os GENERATORS devem ser ajustados manualmente.');
    end;
  end
  else
    MostraErro('Ocorreu algum erro ao criar o arquivo de banco de dados novo.', False);

  if CheckBoxDesativaTriggers.Checked then
    SetTriggersOnOff(True);

  AddLog('Cópia Finalizada !');
  AddLog('Tempo de Operação: ' + FormatDateTime('hh:nn:ss', Now - lStart));
  AddLog('Processo finalizado, ' + IntToStr(FImportados + FErros) + ' processados. '
       + ' Ok = ' + IntToStr(FImportados)
       + IfThen(FErros > 0, ' | Erros = ' + IntToStr(FErros) ) );

  AddFileSizeToLog(NomeNovoBanco, EditArquivoOrigem.Text);

  Application.MessageBox('Processo finalizado. Verifique o arquivo de log para mais informações.', 'Aviso', MB_ICONINFORMATION + MB_OK);
  DoEnableFormForImport(True);

  if IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.Commit;
  if IBODatabaseDestino.Connected then
    IBODatabaseDestino.Disconnect;

  if IBODatabaseOrigem.Connected then
    IBODatabaseOrigem.Disconnect;

  BitBtnIniciar.Enabled := True;
  BitBtnConectar.Enabled := True;
end;

procedure TFormDataCopy.BitBtnLoadClick(Sender: TObject);
begin
  with TOpenDialog.Create(Self) do
  try
    Filter := C_FILE_DESC;
    DefaultExt := '*.xdef';
    FileName := FConfigFileName;
    if Execute then
    begin
      FConfigFileName := FileName;
      DoLoadConfig;
    end;
  finally
    Free;
  end;
end;

procedure TFormDataCopy.BitBtnReconectarClick(Sender: TObject);
begin
  DoReconnect;
end;

procedure TFormDataCopy.BitBtnSaveClick(Sender: TObject);
begin
  with TSaveDialog.Create(Self) do
  try
    Filter := C_FILE_DESC;
    DefaultExt := '*.xdef';
    FileName := FConfigFileName;
    if Execute then
    begin
      FConfigFileName := FileName;
      Self.Caption := 'DataCopy - ' + FConfigFileName;
      DoSaveConfig(FConfigFileName);
    end;
  finally
    Free;
  end;
end;

procedure TFormDataCopy.ButtonBancoClick(Sender: TObject);
begin
  with TOpenDialog.Create(Self) do
  try
    Filter := 'Arquivo de Banco de Dados (*.fdb;*.gdb;*.spi)|*.fdb;*.gdb;*.spi';
    DefaultExt := '*.fdb';
    if Execute then
    begin
      EditArquivoOrigem.Text := FileName;
      DoDesconectarDatabase;
      BitBtnConectar.Click;
    end;
  finally
    Free;
  end;
end;

procedure TFormDataCopy.ClientDataSetTabelasAfterScroll(DataSet: TDataSet);
var
  lItens: string;
begin
  if (not ClientDataSetTabelas.ControlsDisabled) then
  begin
    lItens := Trim( StringReplace(ClientDataSetTabelasLISTACAMPODATA.AsString, ';', sLineBreak, [rfReplaceAll]) );
    TcxComboBoxProperties(cxGrid1DBTableView1CAMPODATA.Properties).Items.Text := sLineBreak + lItens;
  end;
end;

procedure TFormDataCopy.ClientDataSetTabelasNewRecord(DataSet: TDataSet);
begin
  ClientDataSetTabelasPESO.AsInteger := DataSet.RecordCount + 1;
end;

function TFormDataCopy.CriaNovoBanco(const NomeBanco : TFileName): Boolean;
var
  lMetadataName: string;
begin
  FStatus.Moderator := 1;
  FStatus.FormCaption := 'Criando Novo Banco de Dados';
  FStatus.MessageCaption := 'Extraindo Metadata do banco original...';
  FStatus.MaxValue := 100;
  FStatus.Execute;
  try
    lMetadataName := GetWindowsDir + ExtractFileName(NomeBanco) + '.meta';

    AddLog( FStatus.MessageCaption );

    IBOExtract1.ExtractObject(eoDatabase);
    IBOExtract1.Items.SaveToFile( lMetadataName );

    FStatus.MessageCaption := 'Criando Novo Banco de Dados...';
    AddLog( FStatus.MessageCaption );

    IBODatabaseDestino.Server := 'LOCALHOST';
    IBODatabaseDestino.PageSize := 8192;
    IBODatabaseDestino.Username := 'SYSDBA';
    IBODatabaseDestino.Password := 'masterkey';
    IBODatabaseDestino.DatabaseName := NomeBanco;
    IBODatabaseDestino.Connect(True);

    IB_ScriptNovo.SQL.LoadFromFile( lMetadataName );
    IB_ScriptNovo.Execute;

    Result := True;
  finally
    FStatus.FreeExecute;
    DeleteFile(lMetadataName);
  end;
end;

procedure TFormDataCopy.cxGrid1DBTableView1CUSTOMSQLPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
var
  lNewSQL: string;
begin
  if TFormSQL.Execute(ClientDataSetTabelasCUSTOMSQL.AsString, lNewSQL) then
  begin
    if not (ClientDataSetTabelas.State in dsEditModes) then
      ClientDataSetTabelas.Edit;
    ClientDataSetTabelasCUSTOMSQL.AsString := lNewSQL;
    ClientDataSetTabelas.Post;
  end;
end;

procedure TFormDataCopy.cxGrid1DBTableView1IMPORTHeaderClick(Sender: TObject);
var
  lRecNo : Integer;
begin
  try
    lRecNo := ClientDataSetTabelas.RecNo;
    ClientDataSetTabelas.DisableControls;
    ClientDataSetTabelas.First;
    while not ClientDataSetTabelas.Eof do
    begin
      ClientDataSetTabelas.Edit;
      ClientDataSetTabelasIMPORT.AsBoolean := cxGrid1DBTableView1IMPORT.Tag = 1;
      ClientDataSetTabelas.Next;
    end;
    ClientDataSetTabelas.RecNo := lRecNo;
    cxGrid1DBTableView1IMPORT.Tag := IfThen(cxGrid1DBTableView1IMPORT.Tag = 1, 0, 1);
  finally
    ClientDataSetTabelas.EnableControls;
  end;
end;

procedure TFormDataCopy.DoConnectDatabase;
var
  lConfig : TConfig;
  lTable: TTableInfo;
begin
  FStatus.Moderator := 1;
  if AnsiUpperCase(EditDestino.Text) = AnsiUpperCase(ExtractFileName(EditArquivoOrigem.Text)) then
    MostraErro('O arquivo de ORIGEM é o mesmo de DESTINO.', True);

  lConfig := TConfig.Create(nil);
  try
    with IBODatabaseOrigem do
    begin
      Server := 'LOCALHOST';
      DatabaseName := EditArquivoOrigem.Text;
      Username := 'SYSDBA';
      Password := 'masterkey';
      Connect;
    end;

    if IsSaveEnabled and (FileExists(FConfigFileName)) then
      lConfig.LoadFromFile(FConfigFileName);

    if IBODatabaseOrigem.Connected then
    begin
      ClientDataSetTabelas.EmptyDataSet;
      ClientDataSetTabelas.DisableControls;
      with IBOQueryTabelas do
      begin
        Open;
        FStatus.FormCaption := 'Verificando Tabelas';
        FStatus.MaxValue := RecordCount;
        FStatus.Execute;
        First;
        while not Eof do
        begin
          FStatus.StepIt(RecNo, 'Verificando ' + FieldByName('Tabela').AsString);

          ClientDataSetTabelas.Append;
          ClientDataSetTabelasNOME.AsString := FieldByName('Tabela').AsString;
          ClientDataSetTabelasREGISTROS.AsInteger := GetTableRecordCount( FieldByName('Tabela').AsString );
          ClientDataSetTabelasLISTACAMPODATA.AsString := GetTableDateList( FieldByName('Tabela').AsString );
          ClientDataSetTabelasSTATUS.AsInteger := 0;
          ClientDataSetTabelasIMPORT.AsBoolean := True;
          ClientDataSetTabelasPAGINACAO.AsInteger := 0;
          ClientDataSetTabelasCAMPODATA.AsString := '';

          lTable := lConfig.TableList.GetTable(FieldByName('Tabela').AsString);
          if Assigned(lTable) then
          begin
            ClientDataSetTabelasCAMPODATA.AsString := lTable.DateField;
            ClientDataSetTabelasIMPORT.AsBoolean := lTable.Generate;
            ClientDataSetTabelasCUSTOMSQL.AsString := lTable.CustomWhere;
            ClientDataSetTabelasPAGINACAO.AsInteger := lTable.PageLimit;
          end;
          ClientDataSetTabelas.Post;
          Next;
        end;
      end;

      ClientDataSetTabelas.First;
      ClientDataSetTabelas.EnableControls;
    end;

    with IBOQueryPesos do
    begin
      if not Prepared then
        Prepare;
      Open;
    end;
  finally
    if IBODatabaseOrigem.Connected then
      BitBtnConectar.Caption := C_DESCONECTAR;

    BitBtnIniciar.Enabled := ClientDataSetTabelas.RecordCount > 0;
    FStatus.FreeExecute;
    FreeAndNil(lConfig);
  end;
end;

procedure TFormDataCopy.DoDesconectarDatabase;
begin
  if IBODatabaseOrigem.Connected then
    IBODatabaseOrigem.Disconnect;

  if not IBODatabaseOrigem.Connected then
  begin
    ClientDataSetTabelas.EmptyDataSet;
    BitBtnIniciar.Enabled := False;
    BitBtnConectar.Caption := C_CONECTAR;
  end;
end;

procedure TFormDataCopy.DoEnableFormForImport(const Value: Boolean);
begin
  EditArquivoOrigem.Enabled := Value;
  ButtonBanco.Enabled := Value;
  EditDestino.Enabled := Value;
  cxDateEditFiltro.Enabled := Value;
  CheckBoxGenerators.Enabled := Value;
  CheckBoxNovo.Enabled := Value;
  CheckBoxDesativaTriggers.Enabled := Value;
  CheckBoxDependencia.Enabled := Value;
end;

procedure TFormDataCopy.DoLoadConfig;
var
  lRecNo : Integer;
  lTable : TTableInfo;
begin
  if not FileExists(FConfigFileName) then
    MostraErro('Arquivo não encontrado.' + sLineBreak + '[' + FConfigFileName + ']', True);

  Self.Caption := 'DataCopy - ' + FConfigFileName;

  if Assigned(FConfig) then
    FreeAndNil(FConfig);

  FConfig := TConfig.Create(nil);
  try
    FConfig.LoadFromFile(FConfigFileName);

    EditDestino.Text := FConfig.BancoDestino;

    // se esta conectado, desconecta para refazer a lista do banco de origem
    if BitBtnConectar.Caption = C_DESCONECTAR then
      BitBtnConectar.Click;

    EditArquivoOrigem.Text := FConfig.BancoOrigem;
    EditArquivoOrigem.Hint := FConfig.BancoOrigem;

    CheckBoxDesativaTriggers.Checked := FConfig.DesativarTriggers;
    CheckBoxIgnorar.Checked := FConfig.IgnorarErros;
    CheckBoxNovo.Checked := FConfig.CriarNovoBanco;
    CheckBoxDependencia.Checked := FConfig.ForcarDependencia;
    cxDateEditFiltro.Date := FConfig.DataFiltro;
    CheckBoxGenerators.Checked := FConfig.Generators;

    lRecNo := ClientDataSetTabelas.RecNo;
    ClientDataSetTabelas.DisableControls;
    ClientDataSetTabelas.First;
    while not ClientDataSetTabelas.Eof do
    begin
      lTable := FConfig.TableList.GetTable(ClientDataSetTabelasNOME.AsString);
      if Assigned(lTable) then
      begin
        ClientDataSetTabelas.Edit;
        ClientDataSetTabelasIMPORT.AsBoolean := lTable.Generate;
        ClientDataSetTabelasCAMPODATA.AsString := lTable.DateField;
        ClientDataSetTabelasCUSTOMSQL.AsString := lTable.CustomWhere;
        ClientDataSetTabelasPAGINACAO.AsInteger := lTable.PageLimit;
        ClientDataSetTabelas.Post;
      end;
      ClientDataSetTabelas.Next;
    end;
    ClientDataSetTabelas.RecNo := lRecNo;
  finally
    ClientDataSetTabelas.EnableControls;
  end;

  if BitBtnConectar.Enabled and (FileExists(EditArquivoOrigem.Text)) then
    BitBtnConectar.Click;
end;

function TFormDataCopy.DoReconnect: Boolean;
type
  TExecState = (esNormal, esMinimized, esMaximized, esHidden);
const
  C_SHOW_COMMANDS: array[TExecState] of Integer = (SW_SHOWNORMAL, SW_MINIMIZE, SW_SHOWMAXIMIZED, SW_HIDE);

  function FileExecuteWait(const FileName, Params, StartDir: string; InitialState: TExecState): Integer;
  var
    Info: TShellExecuteInfo;
    ExitCode: DWORD;
  begin
    FillChar(Info, SizeOf(Info), 0);
    Info.cbSize := SizeOf(TShellExecuteInfo);
    with Info do
    begin
      fMask := SEE_MASK_NOCLOSEPROCESS;
      Wnd := Application.Handle;
      lpFile := PChar(FileName);
      lpParameters := PChar(Params);
      lpDirectory := PChar(StartDir);
      nShow := C_SHOW_COMMANDS[InitialState];
    end;
    if ShellExecuteEx(@Info) then
    begin
      repeat
        Application.ProcessMessages;
        GetExitCodeProcess(Info.hProcess, ExitCode);
      until (ExitCode <> STILL_ACTIVE) or Application.Terminated;
      Result := ExitCode;
    end
    else
      Result := -1;
  end;

begin
  // para o servico
  try
    BitBtnReconectar.Enabled := False;

    FStatus.MessageCaption := 'Parando serviço do FireBird...';
    AddLog('Parando serviço do FireBird...');
    FileExecuteWait('sc', 'stop FirebirdGuardianDefaultInstance', '', esNormal);
    FileExecuteWait('sc', 'stop FirebirdServerDefaultInstance', '', esNormal);
    Sleep(2000);
    // inicia o servico
    FStatus.MessageCaption := 'Iniciando serviço do FireBird...';
    AddLog('Iniciando serviço do FireBird...');
    FileExecuteWait('sc', 'start FirebirdServerDefaultInstance', '', esNormal);
    FileExecuteWait('sc', 'start FirebirdGuardianDefaultInstance', '', esNormal);
    Sleep(5000);

    AddLog('FireBird Reiniciado com sucesso!');
  finally
    BitBtnReconectar.Enabled := True;
  end;
  Result := True;
end;

procedure TFormDataCopy.DoSaveConfig(const AFileName : string);
var
  lRecNo : Integer;
  lTable : TTableInfo;
  lIsUpdate : Boolean;
begin
  lIsUpdate := not EditArquivoOrigem.Enabled;

  if not Assigned(FConfig) then
  begin
    FConfig := TConfig.Create(nil);
    FConfig.FormatType := bfJSON;
  end;

  try
    FConfig.IncludeClassName := False;
    FConfig.BancoDestino := EditDestino.Text;
    FConfig.BancoOrigem := EditArquivoOrigem.Text;
    FConfig.DesativarTriggers := CheckBoxDesativaTriggers.Checked;
    FConfig.ForcarDependencia := CheckBoxDependencia.Checked;
    FConfig.IgnorarErros := CheckBoxIgnorar.Checked;
    FConfig.DataFiltro := cxDateEditFiltro.Date;
    FConfig.Generators := CheckBoxGenerators.Checked;

    if lIsUpdate then
    begin
      FConfig.CriarNovoBanco := False;
      lTable := FConfig.TableList.GetTable(ClientDataSetTabelasNOME.AsString);
      if Assigned(lTable) then
        lTable.Generate := False;
    end else
    begin
      FConfig.CriarNovoBanco := CheckBoxNovo.Checked;

      lRecNo := ClientDataSetTabelas.RecNo;
      ClientDataSetTabelas.DisableControls;
      ClientDataSetTabelas.First;
      while not ClientDataSetTabelas.Eof do
      begin
        lTable := FConfig.TableList.NewTable(ClientDataSetTabelasNOME.AsString);
        if Assigned(lTable) then
        begin
          lTable.Generate := ClientDataSetTabelasIMPORT.AsBoolean;
          lTable.DateField := ClientDataSetTabelasCAMPODATA.AsString;
          lTable.CustomWhere := ClientDataSetTabelasCUSTOMSQL.AsString;
          lTable.PageLimit := ClientDataSetTabelasPAGINACAO.AsInteger;
        end;
        ClientDataSetTabelas.Next;
      end;
      ClientDataSetTabelas.RecNo := lRecNo;
    end;

    FConfig.SaveToFile(AFileName);
  finally
    ClientDataSetTabelas.EnableControls;
  end;
end;

procedure TFormDataCopy.DoSetGenerators;
var
  lSQL: string;
begin
  if not IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.StartTransaction;

  AddLog('Ajustando Generators...');
  FStatus.Moderator := 1;
  FStatus.FormCaption := 'Generators';
  FStatus.MessageCaption := 'Aguarde... Ajustando Generators...';
  FStatus.Execute;

  with IBOQueryGeneratorsOrigem do
  try
    if not Prepared then
      Prepare;
    Open;
    First;
    FStatus.MaxValue := RecordCount;

    while not Eof do
    begin
      FStatus.StepIt(FStatus.ProgressBarStatus.Position + 1, 'Ajustando ' + FieldByName('genname').AsString);
      try
        lSQL := 'set generator ' + Trim(FieldByName('genname').AsString) + ' to ' + Trim(FieldByName('genval').AsString);
        AddLog(lSQL, False, False);
        try
          IB_DSQLDestino.ExecuteDDL(lSQL);
          if IB_TransactionDestino.InTransaction then
            IB_TransactionDestino.CommitRetaining;
        except on E:Exception do
          begin
            AddLog('Erro ajustando generator: ' + E.Message );
          end;
        end;
      finally
        Next;
      end;
    end;
  finally
    Close;
  end;

  if IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.Commit;

  FStatus.Reset;
  // verifica se os valores foram gravados corretamente
  IBOQueryGeneratorsDestino.Open;

  IBOQueryGeneratorsOrigem.Open;
  IBOQueryGeneratorsOrigem.First;

  FStatus.Moderator := 1;
  FStatus.FormCaption := 'Validando Generators';
  FStatus.MessageCaption := 'Aguarde... Validando Generators...';
  FStatus.MaxValue := IBOQueryGeneratorsOrigem.RecordCount;
  FStatus.Execute;
  while not IBOQueryGeneratorsOrigem.Eof do
  begin
    FStatus.StepIt(FStatus.ProgressBarStatus.Position + 1, 'Validando ' + IBOQueryGeneratorsOrigem.FieldByName('genname').AsString);
    if IBOQueryGeneratorsDestino.Locate('genname', IBOQueryGeneratorsOrigem.FieldByName('genname').AsString, []) then
    begin
      if IBOQueryGeneratorsOrigem.FieldByName('genval').AsFloat <> IBOQueryGeneratorsDestino.FieldByName('genval').AsFloat then
        AddLog(Format('ERRO: GEN %s, Expected Value: %d, Current Value: %d', [
          Trim(IBOQueryGeneratorsOrigem.FieldByName('genname').AsString),
          IBOQueryGeneratorsOrigem.FieldByName('genval').AsInteger,
          IBOQueryGeneratorsDestino.FieldByName('genval').AsInteger ])) ;
    end else
      AddLog('ERRO: GEN ' + IBOQueryGeneratorsOrigem.FieldByName('genname').AsString + ' não foi localizado no destino.');

    IBOQueryGeneratorsOrigem.Next;
  end;

  FStatus.FreeExecute;
end;

procedure TFormDataCopy.EditArquivoOrigemEnter(Sender: TObject);
begin
  EditArquivoOrigem.Hint := EditArquivoOrigem.Text;
end;

procedure TFormDataCopy.EditArquivoOrigemExit(Sender: TObject);
begin
  if EditArquivoOrigem.Hint <> EditArquivoOrigem.Text then
  begin
    DoDesconectarDatabase;
    BitBtnConectar.Click;
  end;
end;

function TFormDataCopy.FileSize(const aFilename: String): Int64;
var
  lInfo: TWin32FileAttributeData;
  lOk: Boolean;
begin
  Result := -1;

  lOk := GetFileAttributesEx(PWideChar(aFileName), GetFileExInfoStandard, @lInfo);

  if not lOk then
    Exit;

  Result := lInfo.nFileSizeLow or (lInfo.nFileSizeHigh shl 32);
end;

procedure TFormDataCopy.FormCreate(Sender: TObject);
begin
  FConfig := nil;

  FStatus := TFormStatus.Create(Self);
  FStatus.Parent := Self;
//  RafaArquivoEditOrigem.InitialDir := ExtractFilePath(Application.ExeName);
  ClientDataSetTabelas.CreateDataSet;
  ClientDataSetTabelas.Open;

  cdsCampos.CreateDataSet;

  cxDateEditFiltro.Properties.MaxDate := Date;
  cxDateEditFiltro.Date := IncMonth(Date, -12);

  LoadVersion;
end;

procedure TFormDataCopy.FormDestroy(Sender: TObject);
begin
  if Assigned(FStatus) then
  begin
    FStatus.FreeExecute;
    FreeAndNil(FStatus);
  end;

  if Assigned(FConfig) then
    FreeAndNil(FConfig);
end;

procedure TFormDataCopy.FormShow(Sender: TObject);
begin
  EditArquivoOrigem.Text := '';
  PostMessage(Handle, WM_AFTER_SHOW, 0, 0);
end;

function TFormDataCopy.GetInsertSQL(const NomeTabela : string): string;
var aCampos, aParans, AuxCampo, AuxParam : string;
    aTab : Integer;
begin
  aTab := 1;
  AuxCampo := StringOfChar(' ', aTab + 2);
  AuxParam := StringOfChar(' ', aTab + 2);
  aCampos := '';
  aParans := '';
  with cdsCampos do
  begin
    First;
    while not Eof do
    begin
      AuxCampo := AuxCampo + FieldByName('Nome').AsString + ', ';
      AuxParam := AuxParam + ':i' + FieldByName('Nome').AsString + ', ';
      if Length(AuxCampo) + aTab > 80 then
      begin
        aCampos := aCampos + AuxCampo + sLineBreak + StringOfChar(' ', aTab + 2);
        AuxCampo := '';
      end;
      if Length(AuxParam) + aTab > 80 then
      begin
        aParans := aParans + AuxParam + sLineBreak + StringOfChar(' ', aTab + 2);
        AuxParam := '';
      end;
      Next;
    end;
  end;
  aCampos := Trim('.' + aCampos + AuxCampo);
  aParans := Trim('.' + aParans + AuxParam);
  Delete(aCampos, Length(aCampos), 3);
  Delete(aCampos, 1, 1);
  Delete(aParans, Length(aParans), 3);
  Delete(aParans, 1, 1);
  Result := Format('insert into %s (' + sLineBreak + '%s)'
                 + sLineBreak + '%*.*svalues(' + sLineBreak + '%s)',
            [NomeTabela, aCampos, aTab, aTab, ' ', aParans]);
end;

function TFormDataCopy.GetInsertUpdateSQL(const NomeTabela : string): string;
var aCampos, aParans, AuxCampo, AuxParam : string;
    aTab : Integer;
    aMatching, AuxMatching : string;
begin
  aTab := 1;
  AuxCampo := StringOfChar(' ', aTab + 2);
  AuxParam := StringOfChar(' ', aTab + 2);
  AuxMatching := '';

  aCampos := '';
  aParans := '';
  aMatching := '';

  with cdsCampos do
  begin
    First;
    while not Eof do
    begin
      AuxCampo := AuxCampo + FieldByName('Nome').AsString + ', ';
      AuxParam := AuxParam + ':i' + FieldByName('Nome').AsString + ', ';
      if Length(AuxCampo) + aTab > 80 then
      begin
        aCampos := aCampos + AuxCampo + sLineBreak + StringOfChar(' ', aTab + 2);
        AuxCampo := '';
      end;
      if Length(AuxParam) + aTab > 80 then
      begin
        aParans := aParans + AuxParam + sLineBreak + StringOfChar(' ', aTab + 2);
        AuxParam := '';
      end;
      if FieldByName('PrimaryKey').AsBoolean then
      begin
        AuxMatching := AuxMatching + FieldByName('Nome').AsString + ', ';
        if Length(AuxMatching) + aTab > 80 then
        begin
          aMatching := aMatching + AuxMatching + sLineBreak + StringOfChar(' ', aTab + 2);
          AuxMatching := '';
        end;
      end;

      Next;
    end;
  end;
  aCampos := Trim('.' + aCampos + AuxCampo);
  aParans := Trim('.' + aParans + AuxParam);
  aMatching := Trim('.' + aMatching + AuxMatching);

  Delete(aCampos, Length(aCampos), 3);
  Delete(aCampos, 1, 1);
  Delete(aParans, Length(aParans), 3);
  Delete(aParans, 1, 1);
  Delete(aMatching, Length(aMatching), 3);
  Delete(aMatching, 1, 1);

  Result := Format('update or insert into %s (' + sLineBreak + '%s)'
                 + sLineBreak + '%*.*svalues(' + sLineBreak + '%s)'
                 + sLineBreak + ' matching (%s)',
            [NomeTabela, aCampos, aTab, aTab, ' ', aParans, aMatching]);
end;

function TFormDataCopy.GetTableDateList(const ANomeTabela: string): string;
begin
  Result := '';
  try
    with IBOQueryCampos do
    try
      Close;
      if not Prepared then
        Prepare;
      Params.ParamByName('Nome_Tabela').AsString := ANomeTabela;
      Params.ParamByName('FilterDate').AsInteger := 1;
      Open;
      First;
      while not Eof do
      begin
        Result := Result + Trim(FieldByName('Coluna').AsString) + ';';
        Next;
      end;
    finally
      Close;
    end;
  except
    on E:Exception do
    begin
      Result := '.: ERRO :.';
      ReConnect(IBOQueryTabelas, E.Message);
    end;
  end;
end;

function TFormDataCopy.GetTableRecordCount(const ANomeTabela: string): Integer;
begin
  Result := -1;
  try
    if not IBOQueryRegistros.Active then
    begin
      IBOQueryRegistros.Prepare;
      IBOQueryRegistros.DisableControls;
      IBOQueryRegistros.Open;
    end;

    if IBOQueryRegistros.Locate('Table_Name', ANomeTabela, [loCaseInsensitive]) then
      Result := IBOQueryRegistros.FieldByName('Qtd').AsInteger;

    if Result = -1 then
      Result := GetTableRecordCountEx(ANomeTabela);
  except
    on E:Exception do
    begin
      if Result <= 0 then
        Result := -1;
      ReConnect(IBOQueryTabelas, E.Message);
    end;
  end;
end;

function TFormDataCopy.GetTableRecordCountEx(const ANomeTabela: string): Integer;
begin
  Result := 0;
  try
    with TIBOQuery.Create(nil) do
    try
      IB_Connection := IBOQueryRegistros.IB_Connection;
      IB_Transaction := IBOQueryRegistros.IB_Transaction;
      SQL.Add('select count(1) as Qtd from ' + ANomeTabela);
      if not Prepared then
        Prepare;
      Open;
      Result := FieldByName('Qtd').AsInteger;
    finally
      Close;
      UnPrepare;
      Free;
    end;
  except
    on E:Exception do
    begin
      if Result = 0 then
        Result := -1;
      ReConnect(IBOQueryTabelas, E.Message);
    end;
  end;
end;

function TFormDataCopy.GetWindowsDir: string;
var
  tempFolder: array[0..MAX_PATH] of Char;
begin
  GetTempPath(MAX_PATH, @tempFolder);
  Result := IncludeTrailingPathDelimiter( StrPas(tempFolder) );
end;

procedure TFormDataCopy.IB_ScriptNovoStatement(Sender: TIB_Script;
  var Statement: string; var SkipIt: Boolean);
var
  i: Integer;
begin
  if FStatus.ProgressBarStatus.Position = 100 then
    FStatus.Reset;

  FStatus.StepIt(FStatus.ProgressBarStatus.Position + 1, StringReplace(Statement, sLineBreak, ' ', [rfReplaceAll]));

  Application.ProcessMessages;

  SkipIt := False;
  for i := 0 to High(ConstSkipCreate) do
  begin
    SkipIt := Pos(ConstSkipCreate[i], Statement) > 0;
    if SkipIt then
      Break;
  end;
end;

function TFormDataCopy.ImportaTabela(const ANomeTabela: string; const ATotalRegistros: Integer;
  const AIsDependencia : Boolean): Boolean;
var
  TemChave : Boolean;
  lNomeCampo: String;
  Ultimo : string;
  ProcessoOriginal : string;
  lRecNo: Integer;
  lAddWhere: Boolean;
  lCustomSQL: string;
  lTotal: LongInt;
  lQryProc: TIBOQuery;
  lPreparedOk: Boolean;

  function DeveImportarTabela : Boolean;
  begin
    Result := ClientDataSetTabelasIMPORT.AsBoolean
           or (AIsDependencia and CheckBoxDependencia.Checked);
  end;

  procedure DoNext;
  begin
    try
      lQryProc.Next;
    except
      lQryProc.Close;
      if not lQryProc.IB_Transaction.InTransaction then
        lQryProc.IB_Transaction.StartTransaction;

      lQryProc.Open;
      lQryProc.RecNo := lRecNo + 1;
    end;
  end;

begin
  // antes de importar verifica as dependencias
  Result := False;
  FStatus.Moderator := 1;
  FStatus.FormCaption := 'Importando';
  FStatus.MessageCaption := 'Aguarde... Preparando Tabela ' + ANomeTabela;
  FStatus.Execute;

  ProcessoOriginal := ANomeTabela;

  // não importa
  if not DeveImportarTabela then
  begin
    MensIgnorarTabela;
    ClientDataSetTabelas.Edit;
    ClientDataSetTabelasSTATUS.AsInteger := 4;
    ClientDataSetTabelas.Post;
    Result := True;
    Exit;
  end;

  // verifica dependencias
  lQryProc := TIBOQuery.Create(Self);
  with lQryProc do
  try
    Close;
    SQL.Text := IBOQueryPesos.SQL.Text;
    IB_Connection := IBOQueryPesos.IB_Connection;
    IB_Transaction := IBOQueryPesos.IB_Transaction;
    if not Prepared then
      Prepare;
    {$IFDEF TESTE}
    OutputDebugString(PChar(NomeTabela));
    {$ENDIF}
    Params.ParamByName('Nome').AsString := ANomeTabela;
    Open;
    if not IsEmpty then
    begin
      First;
      while not Eof do
      begin
        lRecNo := RecNo;
        lNomeCampo := FieldByName('TABPRINCIPAL').AsString;
        {$IFDEF TESTE}
        OutputDebugString(PChar(NomeTabela + '->' + lNomeCampo ));
        {$ENDIF}
        ClientDataSetTabelas.Locate('NOME', lNomeCampo, []);
        if ClientDataSetTabelasSTATUS.AsInteger <> 0 then
          DoNext
        else
        begin
          Ultimo := lNomeCampo;
          if ImportaTabela(lNomeCampo, ClientDataSetTabelasREGISTROS.AsInteger, True) then  // False ? True para forçar as dependencias
          begin // Importa certo
            ClientDataSetTabelas.Locate('NOME', Ultimo, []);
            ClientDataSetTabelas.Edit;
            if ClientDataSetTabelas.FieldByName('STATUS').AsInteger <> 4 then
              ClientDataSetTabelas.FieldByName('STATUS').AsInteger := 2;
//          ClientDataSetTabelas.FieldByName('IMPORT').AsBoolean := False;
            ClientDataSetTabelas.Post;
          end
          else
          begin // Erro
            ClientDataSetTabelas.Locate('NOME', Ultimo, []);
            ClientDataSetTabelas.Edit;
            if ClientDataSetTabelas.FieldByName('STATUS').AsInteger <> 4 then
              ClientDataSetTabelas.FieldByName('STATUS').AsInteger := 3;
//          ClientDataSetTabelas.FieldByName('IMPORT').AsBoolean := False;
            ClientDataSetTabelas.Post;
          end;
          Inc(FTabelas);
          AtualizaStatus(stTabela);
          DoSaveConfig(FTmpConfigFileName);
          DoNext;
        end;
      end;
    end;
  finally
    Close;
    UnPrepare;
    Free;
  end;

  ClientDataSetTabelas.Locate('NOME', ProcessoOriginal, []);

  if not IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.StartTransaction;

  with ClientDataSetTabelas do
  begin
    Edit;
    FieldByName('STATUS').AsInteger := 1;
    Post;
  end;

  TemChave := False;
  // configura o progresso
  FStatus.FormCaption := 'Importando';
  FStatus.MessageCaption := 'Aguarde... Preparando Dados para ' + ANomeTabela;
  FStatus.MaxValue := ATotalRegistros;
  FStatus.Moderator := 7;
  FStatus.Execute;
  // ajusta a tabela de ORIGEM
  lAddWhere := False;
  with IB_QueryDadosOrigem do
  begin
    Close;
    SQL.Clear;
    // no SQL personalizado, pode ter todo o {Select} ou somente a consulta por {Data}
    lCustomSQL := Trim( ClientDataSetTabelasCUSTOMSQL.AsString );
    if Pos(ConstTagSELECT, lCustomSQL) > 0 then
      SQL.Add(StringReplace(lCustomSQL, ConstTagSELECT, 'SELECT', [rfReplaceAll, rfIgnoreCase]))
    else begin
      SQL.Add('SELECT * FROM ' + ANomeTabela);
      if FAplicarFiltroData and (ClientDataSetTabelasCAMPODATA.AsString <> '') then
      begin
        lAddWhere := True;
        SQL.Add( Format('WHERE CAST(%s as date) >= %s', [
          ClientDataSetTabelasCAMPODATA.AsString,
          ConstTagDATA] ) );
      end;
    end;

    // somente where adicional
    if (lCustomSQL <> EmptyStr) and (Pos(ConstTagSELECT, lCustomSQL) <= 0) then
    begin
      SQL.Add( Ifthen(lAddWhere, ' and ', ' where ')
             + lCustomSQL );
    end;

    SQL.Text := StringReplace(SQL.Text, ConstTagDATA, QuotedStr(FormatDateTime('dd.mm.yyyy', cxDateEditFiltro.Date)), [rfIgnoreCase, rfReplaceAll]);

    AddLog(ANomeTabela);
    AddLog(IB_QueryDadosOrigem.SQL.Text, False, False);

    try
      if not Prepared then
        Prepare;
      AddLog(StatementPlan, True, True);
      Open;
    except
      on E:Exception do
      begin
        AddLog('ERRO abrindo tabela ' + ANomeTabela + ': ' + E.Message, True, False);
        ReConnect(IB_QueryDadosOrigem, E.Message);
      end;
    end;
  end;

  { prepara a chave da tabela }
  with IBOQueryChave do
  try
    Close;
    if not Prepared then
      Prepare;
    Params.ParamByName('Tabela').AsString := ANomeTabela;
    Open;
  except
    on E:Exception do
    begin
      AddLog('ERRO obtendo chaves para tabela ' + ANomeTabela + ': ' + E.Message, True, False);
      ReConnect(IBOQueryChave, E.Message);
    end;
  end;

  { prepara os campos da tabela }
  if not cdsCampos.Active then
    cdsCampos.Open;

  with IBOQueryCampos do
  begin
    Close;
    try
      if not Prepared then
        Prepare;
      Params.ParamByName('Nome_Tabela').AsString := ANomeTabela;
      Params.ParamByName('FilterDate').AsInteger := 0;
      Open;
    except
      on E:Exception do
      begin
        AddLog('ERRO verificando chaves para ' + ANomeTabela + ': ' + E.Message);
        ReConnect(IBOQueryCampos, E.Message);
        TemChave := False;
      end;
    end;
    cdsCampos.EmptyDataSet;

    if Active then
    begin
      First;
      while not Eof do
      begin
        cdsCampos.Append;
        cdsCamposNome.AsString := FieldByName('Coluna').AsString;
        // Verifica se o campo é uma chave
        cdsCamposPrimaryKey.AsBoolean := IBOQueryChave.Locate('Rdb$Field_Name', cdsCamposNome.AsString, [loCaseInsensitive]);
        if cdsCamposPrimaryKey.AsBoolean and (not TemChave) then
          TemChave := True;
        cdsCampos.Post;
        Next;
      end;
      Close;
      UnPrepare;
    end;
  end;
  IBOQueryChave.Close;
  IBOQueryChave.UnPrepare;

  try
    with IB_DSQLDestino do
    begin
      if Prepared then
        Unprepare;
      SQL.Clear;
      if TemChave then
        SQL.Add( GetInsertUpdateSQL(ANomeTabela) )
      else
        SQL.Add( GetInsertSQL(ANomeTabela) );
      if not Prepared then
        Prepare;
    end;
    lPreparedOk := True;
  except
    on E:Exception do
    begin
      AddLog('Erro abrindo tabela de destino. ' + E.Message);
      lPreparedOk := False;
    end;
  end;

  if lPreparedOk then
  begin
    lTotal := ATotalRegistros;
    if IB_QueryDadosOrigem.Active then
    begin
      IB_QueryDadosOrigem.First;

      if (ClientDataSetTabelasCAMPODATA.AsString <> '') or (lCustomSQL <> '') or (lTotal = 0) then
      begin
        try
          lTotal := IB_QueryDadosOrigem.RecordCount;
        except
          on E:Exception do
          begin
            AddLog('Erro verificando total de registros. ' + E.Message);
            ReConnect(IB_QueryDadosOrigem, E.Message);
            lTotal := 0;
          end;
        end;
        FStatus.MaxValue := lTotal;
      end;

      Result := ProcessarRegistrosTabela(lTotal, ANomeTabela);
    end else
      Result := False;
  end;

  if IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.Commit;

  if MemoLog.Lines.Count > 0 then
    MemoLog.Lines.SaveToFile(FNomeLog);

  FStatus.FreeExecute;
end;

function TFormDataCopy.IsSaveEnabled: Boolean;
begin
  Result := FConfigFileName <> EmptyStr;
end;

procedure TFormDataCopy.LoadVersion;
var
  Exe: string;
  Size, Handle: DWORD;
  Buffer: TBytes;
  FixedPtr: PVSFixedFileInfo;
begin
  Exe := ParamStr(0);
  Size := GetFileVersionInfoSize(PChar(Exe), Handle);
  if Size = 0 then
    RaiseLastOSError;

  SetLength(Buffer, Size);
  if not GetFileVersionInfo(PChar(Exe), Handle, Size, Buffer) then
    RaiseLastOSError;

  if not VerQueryValue(Buffer, '\', Pointer(FixedPtr), Size) then
    RaiseLastOSError;

  LabelVersao.Caption := Format('%d.%d.%d.%d',
    [LongRec(FixedPtr.dwFileVersionMS).Hi,  //major
     LongRec(FixedPtr.dwFileVersionMS).Lo,  //minor
     LongRec(FixedPtr.dwFileVersionLS).Hi,  //release
     LongRec(FixedPtr.dwFileVersionLS).Lo]) //build
end;

procedure TFormDataCopy.MensIgnorarTabela;
begin
  FStatus.Moderator := 1;
  FStatus.FormCaption := 'Importando';
  FStatus.MessageCaption := 'Aguarde... Ignorando Tabela ' + ClientDataSetTabelas.FieldByName('NOME').AsString;
  AddLog(FStatus.MessageCaption, True, True);
  FStatus.Execute;
  Application.ProcessMessages;
end;

procedure TFormDataCopy.MostraErro(const AMens: string; const AAbort : Boolean);
begin
  Application.MessageBox(PChar(AMens), 'Erro', MB_ICONERROR);

  if AAbort then
  begin
    BitBtnIniciar.Enabled := True;
    BitBtnConectar.Enabled := True;
    Abort;
  end;
end;

function TFormDataCopy.MostraPergunta(const APergunta: string): TModalResult;
begin
  Result := Application.MessageBox(PChar(APergunta), 'Atenção', MB_ICONQUESTION + MB_YESNO);
end;

procedure TFormDataCopy.MostraPerigo(const AMens: string);
begin
  Application.MessageBox(PChar(AMens), 'Atenção', MB_ICONWARNING);
end;

function TFormDataCopy.ProcessarRegistrosTabela(const ATotalRegistros : Int64;
  const ANomeTabela : string): Boolean;
var
  lFazerPaginacao: Boolean;
  lRecNo: Integer;
  i: Integer;
  NomeCampo: string;
  lDic: TFieldIdx;
  lIdxParam: integer;
  lSqlOrigem: string;
  lItensPorPagina: Integer;
  lSkip: Integer;
  lRegOk: Integer;
  lRegErro: Integer;
  lTotalRegistros: Int64;

  procedure TratarErroImportacao(const AError : string);
  var
    lRestartOk : Boolean;
    lBkpPos: Integer;
    lError: string;
  begin
    lRestartOk := False;
    lBkpPos := lRecNo;
    lError := AError;

    while not lRestartOk do
    begin
      AddLog(Format('-- %s [%d]: %s',
        [ANomeTabela, lBkpPos, StringReplace(lError, sLineBreak, ' ', [rfIgnoreCase, rfReplaceAll]) ]
      ), False, False);
      Result := False;
      FErros := FErros + 1;
      Inc(lRegErro);
      AtualizaStatus(stErro);

      if Pos('violation of FOREIGN KEY constraint', lError) > 0 then
      begin
        IB_QueryDadosOrigem.Next;
        lRestartOk := True;
      end else
      begin
        try
          // para nao perder os dados ate o momento
          try
            if IB_DSQLDestino.IB_Transaction.InTransaction then
              IB_DSQLDestino.IB_Transaction.CommitRetaining;
          except
            OutputDebugString('Erro ao realizar commit no tratamento do erro');
          end;

          lBkpPos := lBkpPos + 1;
          ReConnect(IB_QueryDadosOrigem, lError, False);
          Application.ProcessMessages;
          IB_QueryDadosOrigem.RecNo := lBkpPos;

          if not IB_DSQLDestino.IB_Transaction.InTransaction then
            IB_DSQLDestino.IB_Transaction.StartTransaction;

          lRestartOk := True;
        except
          on E:Exception do
          begin
            lError := E.Message;
            AddLog('Tentando Reconectar...');
          end;
        end;
      end;
    end;
    lRecNo := lBkpPos;
  end;

begin
  Result := True;
  lSqlOrigem := EmptyStr;
  lDic := TFieldIdx.Create;
  lTotalRegistros := ATotalRegistros;

  lItensPorPagina := ClientDataSetTabelasPAGINACAO.AsInteger;
  lFazerPaginacao := (lItensPorPagina > 0)
                 and (ATotalRegistros > lItensPorPagina);

  lRecNo := 0;
  lRegOk := 0;
  lRegErro := 0;
  while not IB_QueryDadosOrigem.Eof do
  try
    Inc(lRecNo);

    if ((lRecNo) > lTotalRegistros) and (lTotalRegistros > 0) then
    begin
      lTotalRegistros := GetTableRecordCountEx(ANomeTabela);
      AddLog(Format('IDX ERROR -> Registros Informados [%d] / Registros Existentes [%d]',
        [lRecNo - 1, lTotalRegistros]) );

      if lRecNo > lTotalRegistros then
      begin
        AtualizaStatus(stErro);
        Result := False;
        Break;
      end;
    end;

    if (lRecNo = 1) or (lRecNo mod 17 = 0) then
      FStatus.StepIt(lRecNo, Format('Copiando %s - Registro %d de %d', [ANomeTabela, lRecNo, lTotalRegistros]) );

    if (lRecNo = 1) or (lRecNo mod 100 = 0) then
      Application.ProcessMessages;

    if not IB_DSQLDestino.Prepared then
      IB_DSQLDestino.Prepare;

    for i := 0 to IB_QueryDadosOrigem.FieldCount - 1 do
    begin
      NomeCampo := IB_QueryDadosOrigem.Fields[i].FieldName;
      if NomeCampo <> 'DB_KEY' then
      begin
        if not lDic.TryGetValue('i' + NomeCampo, lIdxParam) then
        begin
          lIdxParam := IB_DSQLDestino.Params.ParamByName('i' + NomeCampo).Index;
          lDic.Add('i' + NomeCampo, lIdxParam);
        end;

        if IB_QueryDadosOrigem.FieldByName(NomeCampo).IsNull then
          IB_DSQLDestino.Params[lIdxParam].Clear
        else
          IB_DSQLDestino.Params[lIdxParam].AsString := IB_QueryDadosOrigem.Fields[i].AsString;
        {$IFDEF TESTE}
        OutputDebugString(PChar(NomeCampo + ': [' + IBOQueryDadosOrigem.Fields[i].AsString + ']'));
        {$ENDIF}
      end;
    end;
    IB_DSQLDestino.Execute;
    Inc(lRegOk);

    if lRecNo mod 5000 = 0 then
      IB_TransactionDestino.Commit; //Retaining;

    FImportados := FImportados + 1;
    if lRecNo mod 17 = 0 then
      AtualizaStatus(stRegistro);

    if lFazerPaginacao then
    begin
      if (lRecNo = 1) or (lSqlOrigem = EmptyStr) then
      begin
        lSqlOrigem := IB_QueryDadosOrigem.SQL.Text;
        Insert(' first %d skip %d ', lSqlOrigem, 7);
      end;

      // a primeira vez nao precisa atualizar o sql, já esta carregado e não foi feito o fetch ainda
      if (lRecNo > 1) and (lRecNo mod lItensPorPagina = 0) then
      begin
        lSkip := lItensPorPagina * (lRecNo div lItensPorPagina);
        IB_QueryDadosOrigem.Close;
        IB_QueryDadosOrigem.Unprepare;
        IB_QueryDadosOrigem.SQL.Text := Format(lSqlOrigem, [
          lItensPorPagina, lSkip]);

        AddLog(Format('Ajustando paginação para %s - First %d Skip %d', [ANomeTabela, lItensPorPagina, lSkip]) );
        FStatus.MessageCaption :=  Format('Ajustando paginação para %s - First %d Skip %d', [ANomeTabela, lItensPorPagina, lSkip]);
        Application.ProcessMessages;

        if not IB_QueryDadosOrigem.Prepared then
          IB_QueryDadosOrigem.Prepare;

        AddLog('/* Paginando... */' + sLineBreak + IB_QueryDadosOrigem.SQL.Text, False, False);
        AddLog(IB_QueryDadosOrigem.StatementPlan, True, True);

        IB_QueryDadosOrigem.Open;
        IB_QueryDadosOrigem.First;
      end else
        IB_QueryDadosOrigem.Next;
    end else
      IB_QueryDadosOrigem.Next;
  except
    on E:Exception do
    begin
      TratarErroImportacao( E.Message );
//      AddLog(Format('-- %s [%d]: %s',
//        [ANomeTabela, lRecNo, StringReplace(E.Message, sLineBreak, ' ', [rfIgnoreCase, rfReplaceAll]) ]
//      ), False, False);
//      Result := False;
//      FErros := FErros + 1;
//      Inc(lRegErro);
//      AtualizaStatus(stErro);
//      IBOQueryDadosOrigem.Next;
    end;
  end;

  IB_QueryDadosOrigem.Close;
  IB_QueryDadosOrigem.UnPrepare;

  AddLog('Tabela finalizada, ' + IntToStr(lRegErro + lRegOk) + ' processados. '
       + ' Ok = ' + IntToStr(lRegOk)
       + IfThen(lRegErro > 0, ' | Erros = ' + IntToStr(lRegErro) ) );

  FreeAndNil( lDic );

  if IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.Commit;
end;

function TFormDataCopy.ReConnect(const AQry: TComponent; const AException: string; const ARestorePos : Boolean) : Boolean;
var
  lRecNo : Integer;
  lStateOrigem: Boolean;
  lStateDestino: Boolean;
  lMensOriginal: string;

  procedure ActiveCloseDataBases(const Value : Boolean);
  begin
    if Value then
    begin
      if lStateDestino then
        IBODatabaseDestino.Connect;

      if lStateOrigem then
        IBODatabaseOrigem.Connect;
    end else
    begin
      if lStateDestino then
        IBODatabaseDestino.Disconnect;

      if lStateOrigem then
        IBODatabaseOrigem.Disconnect;
    end;
  end;

begin
  Result := False;
  lRecNo := -1;
  if ARestorePos then
  begin
    if AQry is TIBOQuery then
      lRecNo := TIBOQuery(AQry).RecNo
    else
      lRecNo := TIB_Query(AQry).RecNo;
  end;

  if AQry is TIBOQuery then
  begin
    lStateDestino := TIBOQuery(AQry).IB_Connection = IBODatabaseDestino;
    lStateOrigem  := TIBOQuery(AQry).IB_Connection = IBODatabaseOrigem;
    TIBOQuery(AQry).Close;
  end else
  begin
    lStateDestino := TIB_Query(AQry).IB_Connection = IBODatabaseDestino;
    lStateOrigem  := TIB_Query(AQry).IB_Connection = IBODatabaseOrigem;
    TIB_Query(AQry).Close;
  end;

  ActiveCloseDataBases(False);

  if (Pos('INTERNAL FIREBIRD CONSISTENCY', UpperCase(AException)) > 0)
  or (Pos('CONNECTION LOST TO DATABASE', UpperCase(AException)) > 0)
  or (Pos('CONTINUE AFTER BUGCHECK', UpperCase(AException)) > 0) then
  begin
    Result := True;

    lMensOriginal := FStatus.MessageCaption;
    lStateDestino := True;
    lStateOrigem  := True;
    ActiveCloseDataBases(False);
    DoReconnect;
    FStatus.MessageCaption := lMensOriginal;
  end;

  ActiveCloseDataBases(True);

  try
    if AQry is TIBOQuery then
    begin
      TIBOQuery(AQry).Open;
      if ARestorePos then
        TIBOQuery(AQry).RecNo := lRecNo;
    end else
    begin
      TIB_Query(AQry).Open;
      if ARestorePos then
        TIB_Query(AQry).RecNo := lRecNo;
    end;
  except
    on E:Exception do
    begin
      AddLog('Erro restaurando conexão. ' + E.Message);
    end;
  end;
end;

procedure TFormDataCopy.ReOpenApp;
var
  lParamatros: string;
  i: Integer;
begin
  lParamatros := '';
  for i := 1 to ParamCount do
    lParamatros := lParamatros + ' ' + ParamStr(i);

  TAdminPriv.RunAsAdmin(Application.ExeName, lParamatros);
  Close;
end;

procedure TFormDataCopy.ResetTableStatus;
var
  lRecNo: Integer;
  lTotal: Integer;
begin
  ClientDataSetTabelas.DisableControls;
  lRecNo := 0;
  lTotal := ClientDataSetTabelas.RecordCount;
  while not ClientDataSetTabelas.Eof do
  begin
    Inc(lRecNo);
    FStatus.StepIt(ClientDataSetTabelas.RecNo, Format('Preparando Registros... %d de %d...', [lRecNo, lTotal]) );

    ClientDataSetTabelas.Edit;
    ClientDataSetTabelasSTATUS.AsInteger := 0;
    ClientDataSetTabelas.Post;
    ClientDataSetTabelas.Next
  end;
  ClientDataSetTabelas.First;
  ClientDataSetTabelas.EnableControls;
end;

procedure TFormDataCopy.SetTriggersOnOff(const AValue: Boolean);
var
  lCmd : string;
  lQtd: integer;
  lName: string;
begin
  FStatus.Reset;
  FStatus.FormCaption := IfThen(Avalue, 'Ativando', 'Desativando') + ' Triggers';
  FStatus.Moderator := 1;
  FStatus.MessageCaption := '';
  FStatus.Execute;

  AddLog('Iniciando ' + FStatus.FormCaption);

  if not IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.StartTransaction;

  with IBOQueryTriggers do
  begin
    Close;
    Open;
    Last;
    First;
    FStatus.MaxValue := RecordCount;
    lQtd := 0;
    while not Eof do
    begin
      lName := FieldByName('rdb$trigger_name').AsString;
      Inc(lQtd);
      FStatus.StepIt(lQtd, 'Alterando ' + lName);
      if AValue then
        lCmd :=  ' active'
      else
        lCmd :=  ' inactive';
      IB_DSQLDestino.ExecuteDDL('alter trigger ' + lName + lCmd);
      Next;
    end;
  end;

  if IB_TransactionDestino.InTransaction then
    IB_TransactionDestino.Commit;

  AddLog('Finalizando ' + FStatus.FormCaption);

  FStatus.FreeExecute;
end;

procedure TFormDataCopy.WmAfterShow(var Dummy: Boolean);
var
  i: Integer;
begin
  if not TAdminPriv.IsAdmin then
  begin
    Application.MessageBox('Em pleno século XXI e você ainda executa programas sem ser como Administrador. Na na na, aperte Ok para corrigir isso.',
      'Aviso', MB_ICONWARNING + MB_OK);

    ReOpenApp;
  end else
  begin
    for i := 1 to ParamCount do
    begin
      if FileExists(ParamStr(i)) then
      begin
        FConfigFileName := ParamStr(i);
        DoLoadConfig;
        Break;
      end;
    end;
  end;
end;

end.