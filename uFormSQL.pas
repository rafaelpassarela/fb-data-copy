unit uFormSQL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ComCtrls;

type
  TFormSQL = class(TForm)
    PanelBotoes: TPanel;
    LabelAviso: TLabel;
    ButtonCancel: TButton;
    ButtonOk: TButton;
    RichEditSQL: TRichEdit;
  private
    { Private declarations }
  public
    { Public declarations }
    class function Execute(const AInputSQL : string; out AOutputSQL : string) : Boolean;
  end;

var
  FormSQL: TFormSQL;

implementation

{$R *.dfm}

{ TForm1 }

class function TFormSQL.Execute(const AInputSQL: string;
  out AOutputSQL: string): Boolean;
begin
  if not Assigned(FormSQL) then
    FormSQL := TFormSQL.Create(Application);

  try
    FormSQL.RichEditSQL.Clear;
    FormSQL.RichEditSQL.Lines.Add(AInputSQL);

    Result := FormSQL.ShowModal = mrOk;
    if Result then
      AOutputSQL := FormSQL.RichEditSQL.Text;
  finally
    FreeAndNil( FormSQL );
  end;
end;

end.
