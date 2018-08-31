unit uFormStatus;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ExtCtrls, StdCtrls;

type
  TFormStatus = class(TForm)
    PanelHeader: TPanel;
    ProgressBarStatus: TProgressBar;
    LabelStatus: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
    FModerator: Cardinal;
    function GetFormCaption: string;
    procedure SetFormCaption(const Value: string);
    function GetMaxValue: Integer;
    procedure SetMaxValue(const Value: Integer);
    function GetMessageCaption: string;
    procedure SetMessageCaption(const Value: string);
  public
    { Public declarations }

    procedure Execute;
    procedure FreeExecute;
    procedure StepIt(const APos : Integer; const AStatus : string);
    procedure Reset;

    property Moderator : Cardinal read FModerator write FModerator;
    property FormCaption : string read GetFormCaption write SetFormCaption;
    property MaxValue : Integer read GetMaxValue write SetMaxValue;
    property MessageCaption : string read GetMessageCaption write SetMessageCaption;
  end;

var
  FormStatus: TFormStatus;

implementation

{$R *.dfm}

{ TFormStatus }

procedure TFormStatus.Execute;
begin
  Reset;
  Self.Show;
  Self.BringToFront;

  Application.ProcessMessages;
end;

procedure TFormStatus.FormShow(Sender: TObject);
begin
  if Assigned(Self.Parent) and (Self.Parent = Application.MainForm) then
  begin
    SetBounds(
      (Self.Parent.Width - Self.Width) div 2,
      (Self.Parent.Height - Self.Height) div 2,
      Self.Width,
      Self.Height);
  end;

  Application.ProcessMessages;
end;

procedure TFormStatus.FreeExecute;
begin
  Self.Close;
end;

function TFormStatus.GetFormCaption: string;
begin
  Result := PanelHeader.Caption;
end;

function TFormStatus.GetMaxValue: Integer;
begin
  Result := ProgressBarStatus.Max;
end;

function TFormStatus.GetMessageCaption: string;
begin
  Result := LabelStatus.Caption;
end;

procedure TFormStatus.Reset;
begin
  ProgressBarStatus.Position := 0;
//  LabelStatus.Caption := '';
end;

procedure TFormStatus.SetFormCaption(const Value: string);
begin
  PanelHeader.Caption := Value;
end;

procedure TFormStatus.SetMaxValue(const Value: Integer);
begin
  ProgressBarStatus.Max := Value;
end;

procedure TFormStatus.SetMessageCaption(const Value: string);
begin
  LabelStatus.Caption := Value;
end;

procedure TFormStatus.StepIt(const APos: Integer; const AStatus: string);
begin
  ProgressBarStatus.Position := APos;
  LabelStatus.Caption := AStatus;
//  if (FModerator = 0) or (APos mod FModerator = 0) or (APos = ProgressBarStatus.Max) then
  Self.Refresh;
end;

end.
