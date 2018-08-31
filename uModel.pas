unit uModel;

interface

uses
  uSPITCPBase, XMLIntf, XMLDoc, SysUtils, Dialogs;

type
  TTableInfo = class(TBase)
  private
    FTableName: string;
    FGenerate: Boolean;
    FDateField: string;
    FCustomSQL: string;
    FPageLimit: Integer;
    function IsBase64File : Boolean;
  protected
    procedure DoSaveToNode; override;
    procedure DoLoadFromNode(const ANode : IXMLNode); override;
    procedure Finalize; override;
    procedure Initialize; override;
  public
    property TableName : string read FTableName write FTableName;
    property Generate : Boolean read FGenerate write FGenerate;
    property DateField : string read FDateField write FDateField;
    property CustomWhere : string read FCustomSQL write FCustomSQL;
    property PageLimit : Integer read FPageLimit write FPageLimit;

    procedure Resetar; override;
    procedure FromOther(const AOther : TBase); override;
  end;

  TTableInfoList = class(TBaseList)
  protected
    function GetItem(Index: Integer): TTableInfo;
    function GetItemClass : TBaseClass; override;
  public
    property Items[Index: Integer]: TTableInfo read GetItem; default;

    function NewTable(const ATableName : string) : TTableInfo;
    function GetTable(const ATableName : string) : TTableInfo;
  end;

  TConfig = class(TBase)
  private
    FBancoOrigem: string;
    FBancoDestino: string;
    FIgnorarErros: Boolean;
    FCriarNovoBanco: Boolean;
    FDesativarTriggers: Boolean;
    FTableList: TTableInfoList;
    FDataFiltro: TDateTime;
    FGenerators: Boolean;
    FForcarDependencia: Boolean;
    FLoadAsBase64: Boolean;
    function IsBase64File(const AFileName : string) : Boolean;
    function GetHandleAsBase64: Boolean;
  protected
    procedure DoSaveToNode; override;
    procedure DoLoadFromNode(const ANode : IXMLNode); override;
    procedure Initialize; override;
    procedure Finalize; override;
  public
    property BancoOrigem : string read FBancoOrigem write FBancoOrigem;
    property BancoDestino : string read FBancoDestino write FBancoDestino;
    property IgnorarErros : Boolean read FIgnorarErros write FIgnorarErros;
    property CriarNovoBanco : Boolean read FCriarNovoBanco write FCriarNovoBanco;
    property DesativarTriggers : Boolean read FDesativarTriggers write FDesativarTriggers;
    property DataFiltro : TDateTime read FDataFiltro write FDataFiltro;
    property Generators : Boolean read FGenerators write FGenerators;
    property ForcarDependencia : Boolean read FForcarDependencia write FForcarDependencia;
    property TableList : TTableInfoList read FTableList write FTableList;

    property HandleAsBase64 : Boolean read GetHandleAsBase64;

    procedure LoadFromFile(const AFile : string; const AFormat : TBaseFormato = bfUnknown); override;
    procedure SaveToFile(const AFile : string); override;
  end;

implementation

{ TConfig }

procedure TConfig.DoLoadFromNode(const ANode: IXMLNode);
begin
  inherited;
  FromNode('BancoDestino', FBancoDestino);
  FromNode('BancoOrigem', FBancoOrigem);
  FromNode('IgnorarErros', FIgnorarErros);
  FromNode('CriarNovoBanco', FCriarNovoBanco);
  FromNode('DesativarTriggers', FDesativarTriggers);
  FromNode('DataFiltro', FDataFiltro);
  FromNode('Generators', FGenerators);
  FromNode('ForcarDependencia', FForcarDependencia);
  FromNode('TableList', FTableList);
end;

procedure TConfig.DoSaveToNode;
begin
  inherited;
  ToNode('BancoDestino', FBancoDestino);
  ToNode('BancoOrigem', FBancoOrigem);
  ToNode('IgnorarErros', FIgnorarErros);
  ToNode('CriarNovoBanco', FCriarNovoBanco);
  ToNode('DesativarTriggers', FDesativarTriggers);
  ToNode('DataFiltro', FDataFiltro);
  ToNode('Generators', FGenerators);
  ToNode('ForcarDependencia', FForcarDependencia);
  ToNode('TableList', FTableList);
end;

procedure TConfig.Finalize;
begin
  inherited;
  FreeAndNil(FTableList);
end;

function TConfig.GetHandleAsBase64: Boolean;
begin
  Result := FLoadAsBase64;
end;

procedure TConfig.Initialize;
begin
  FLoadAsBase64 := True;
  FTableList := TTableInfoList.Create(Self);
  inherited;
end;

function TConfig.IsBase64File(const AFileName: string): Boolean;
begin
  Result := LowerCase(ExtractFileExt(AFileName)) = '.xdef';
end;

procedure TConfig.LoadFromFile(const AFile: string;
  const AFormat: TBaseFormato);
begin
  FLoadAsBase64 := IsBase64File(AFile);
  inherited LoadFromFile(AFile, AFormat);
end;

procedure TConfig.SaveToFile(const AFile: string);
begin
  FLoadAsBase64 := IsBase64File(AFile);
  inherited SaveToFile(AFile);
end;

{ TTableInfoList }

function TTableInfoList.GetItem(Index: Integer): TTableInfo;
begin
  Result := TTableInfo(inherited GetItem(Index));
end;

function TTableInfoList.GetItemClass: TBaseClass;
begin
  Result := TTableInfo;
end;

function TTableInfoList.GetTable(const ATableName: string): TTableInfo;
var
  i: Integer;
begin
  Result := nil;

  for i := 0 to Count - 1 do
    if Trim(Items[i].TableName) = Trim(ATableName) then
    begin
      Result := Items[i];
      Break;
    end;
end;

function TTableInfoList.NewTable(const ATableName: string): TTableInfo;
begin
  Result := GetTable(ATableName);

  if Result = nil then
  begin
    Result := TTableInfo.Create(Self);
    Result.TableName := ATableName;
    Self.Add(Result);
  end;
end;

{ TTableInfo }

procedure TTableInfo.DoLoadFromNode(const ANode: IXMLNode);
begin
  inherited;
  FromNode('TableName', FTableName);
  FromNode('Generate', FGenerate);
  FromNode('DateField', FDateField);
  FromNode('PageLimit', FPageLimit);
  FromNode('CustomSQL', FCustomSQL, IsBase64File);
end;

procedure TTableInfo.DoSaveToNode;
begin
  inherited;
  ToNode('TableName', FTableName);
  ToNode('Generate', FGenerate);
  ToNode('DateField', FDateField);
  ToNode('PageLimit', FPageLimit);
  ToNode('CustomSQL', FCustomSQL, IsBase64File);
end;

procedure TTableInfo.Finalize;
begin
  inherited;
  FTableName := '';
end;

procedure TTableInfo.FromOther(const AOther: TBase);
var
  lTbInfo : TTableInfo absolute AOther;
begin
  inherited;
  FTableName := lTbInfo.TableName;
  FGenerate := lTbInfo.Generate;
  FDateField := lTbInfo.DateField;
  FCustomSQL := lTbInfo.CustomWhere;
  FPageLimit := lTbInfo.PageLimit;
end;

procedure TTableInfo.Initialize;
begin
  inherited;
end;

function TTableInfo.IsBase64File : Boolean;
begin
  Result := False;
  // TTableInfoList -> TConfig
  if Assigned(Owner) and Assigned(Owner.Owner) then
  begin
    if Owner.Owner is TConfig then
      Result := TConfig(Owner.Owner).HandleAsBase64
  end;
end;

procedure TTableInfo.Resetar;
begin
  inherited;
  FTableName := '';
  FGenerate := False;
  FDateField := '';
  FCustomSQL := '';
  FPageLimit := 0;
end;

initialization
  TRegistroBase.RegistroBase.Registrar(TTableInfo);
  TRegistroBase.RegistroBase.Registrar(TTableInfoList);

end.