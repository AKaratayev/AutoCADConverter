object fmMain: TfmMain
  Left = 106
  Top = 346
  Width = 800
  Height = 600
  Caption = #1050#1086#1085#1074#1077#1088#1090#1077#1088' AutoCAD'
  Color = clBtnFace
  Constraints.MinHeight = 600
  Constraints.MinWidth = 800
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pnBtns: TPanel
    Left = 0
    Top = 521
    Width = 784
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object btAutoCAD: TButton
      Left = 24
      Top = 8
      Width = 75
      Height = 25
      Caption = 'AutoCAD'
      TabOrder = 0
      OnClick = btAutoCADClick
    end
  end
  object stgStatistics: TStringGrid
    Left = 0
    Top = 401
    Width = 784
    Height = 120
    Align = alBottom
    DefaultColWidth = 120
    DefaultRowHeight = 18
    RowCount = 2
    TabOrder = 1
  end
  object PageControl: TPageControl
    Left = 0
    Top = 0
    Width = 784
    Height = 401
    ActivePage = tsGraph
    Align = alClient
    TabOrder = 2
    object tsGraph: TTabSheet
      Caption = #1055#1088#1086#1088#1080#1089#1086#1074#1082#1072
      object PaintBox: TPaintBox
        Left = 0
        Top = 0
        Width = 776
        Height = 373
        Align = alClient
        OnPaint = PaintBoxPaint
      end
    end
    object tsResults: TTabSheet
      Caption = #1054#1087#1080#1089#1072#1085#1080#1077
      ImageIndex = 1
      object Memo: TMemo
        Left = 0
        Top = 0
        Width = 776
        Height = 373
        Align = alClient
        TabOrder = 0
      end
    end
  end
  object OpenDialog: TOpenDialog
    DefaultExt = '.dxf'
    Filter = #1060#1072#1081#1083#1099' AutoCAD (*.dxf, *.dwg)|*.dxf;*.dwg'
    Options = [ofHideReadOnly, ofPathMustExist, ofFileMustExist, ofEnableSizing]
    Left = 40
    Top = 48
  end
end
