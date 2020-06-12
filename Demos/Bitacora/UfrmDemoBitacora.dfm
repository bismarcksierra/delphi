object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 326
  ClientWidth = 439
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object btnIniciarProceso: TBitBtn
    Left = 32
    Top = 16
    Width = 145
    Height = 65
    Caption = 'Iniciar proceso'
    TabOrder = 0
    OnClick = btnIniciarProcesoClick
  end
  object btnProcesoNormal: TBitBtn
    Left = 32
    Top = 87
    Width = 145
    Height = 65
    Caption = 'Proceso normal'
    TabOrder = 1
    OnClick = btnProcesoNormalClick
  end
  object btnError: TBitBtn
    Left = 32
    Top = 158
    Width = 145
    Height = 65
    Caption = 'Error'
    TabOrder = 2
    OnClick = btnErrorClick
  end
  object btnDetalle: TBitBtn
    Left = 32
    Top = 229
    Width = 145
    Height = 65
    Caption = 'Detalle'
    TabOrder = 3
    OnClick = btnDetalleClick
  end
  object cmbModoBitacora: TComboBox
    Left = 200
    Top = 16
    Width = 145
    Height = 22
    Style = csOwnerDrawFixed
    TabOrder = 4
    OnChange = cmbModoBitacoraChange
  end
end
