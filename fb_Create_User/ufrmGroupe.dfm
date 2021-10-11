object frmGroupe: TfrmGroupe
  Left = 414
  Top = 243
  Width = 409
  Height = 534
  Caption = 'Liste des groupes'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 120
  TextHeight = 17
  object lstGroupe: TListBox
    Left = 0
    Top = 0
    Width = 514
    Height = 598
    ItemHeight = 17
    MultiSelect = True
    Sorted = True
    TabOrder = 0
  end
  object cmdOK: TButton
    Left = 209
    Top = 607
    Width = 98
    Height = 32
    Caption = 'Ajouter'
    ModalResult = 1
    TabOrder = 1
  end
  object pnlInfo: TPanel
    Left = 0
    Top = 0
    Width = 514
    Height = 54
    Align = alTop
    BevelOuter = bvLowered
    Caption = 'Recherche des groupes en cours ..'
    Color = clBlack
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clLime
    Font.Height = -15
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 2
    Visible = False
  end
end
