object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Demonstration of Commune_APIUtilities'
  ClientHeight = 411
  ClientWidth = 852
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
  object Memo1: TMemo
    Left = 232
    Top = 51
    Width = 577
    Height = 321
    Lines.Strings = (
      'Memo1')
    TabOrder = 0
  end
  object Button_Enumerate: TButton
    Left = 24
    Top = 22
    Width = 177
    Height = 25
    Caption = 'Enumerate All Drives'
    TabOrder = 1
    OnClick = Button_EnumerateClick
  end
  object Button_Close: TButton
    Left = 734
    Top = 378
    Width = 75
    Height = 25
    Caption = 'CLOSE'
    TabOrder = 2
    OnClick = Button_CloseClick
  end
  object ComboBox1: TComboBox
    Left = 232
    Top = 24
    Width = 577
    Height = 21
    TabOrder = 3
    Text = 'ComboBox1'
  end
end
