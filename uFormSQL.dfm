object FormSQL: TFormSQL
  Left = 0
  Top = 0
  Caption = 'Comando SQL'
  ClientHeight = 427
  ClientWidth = 596
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object LabelAviso: TLabel
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 590
    Height = 65
    Align = alTop
    Caption = 
      'Informe um comando SQL personalizado para a consulta. Para subst' +
      'ituir todo o select, utilize a tag {SELECT}. Para utilizar a dat' +
      'a na pesquisa, use a tag {DATA}.'#13#10'Ex.:'#13#10'    {SELECT} a.* from Ta' +
      'bleA a join TableB b on (a.ID = b.ID)'#13#10'    where b.MinDate >= {D' +
      'ATA}'
    WordWrap = True
    ExplicitWidth = 587
  end
  object PanelBotoes: TPanel
    Left = 0
    Top = 386
    Width = 596
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object ButtonCancel: TButton
      AlignWithMargins = True
      Left = 430
      Top = 3
      Width = 75
      Height = 35
      Margins.Right = 10
      Align = alRight
      Cancel = True
      Caption = '&Cancelar'
      ModalResult = 2
      TabOrder = 0
    end
    object ButtonOk: TButton
      AlignWithMargins = True
      Left = 518
      Top = 3
      Width = 75
      Height = 35
      Align = alRight
      Caption = '&Ok'
      ModalResult = 1
      TabOrder = 1
    end
  end
  object RichEditSQL: TRichEdit
    AlignWithMargins = True
    Left = 3
    Top = 74
    Width = 590
    Height = 309
    Align = alClient
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Courier New'
    Font.Style = []
    Lines.Strings = (
      'RichEditSQL')
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 1
    WordWrap = False
  end
end
