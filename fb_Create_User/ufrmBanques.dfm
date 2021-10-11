object frmBanque: TfrmBanque
  Left = 272
  Top = 96
  Width = 449
  Height = 477
  Caption = 'Liste des banques'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 120
  TextHeight = 17
  object lstBanques: TListBox
    Left = 10
    Top = 10
    Width = 546
    Height = 452
    ItemHeight = 17
    MultiSelect = True
    TabOrder = 0
    OnClick = lstBanquesClick
  end
  object pnlGetBanques: TPanel
    Left = 10
    Top = 471
    Width = 546
    Height = 53
    BevelOuter = bvNone
    TabOrder = 1
    object cmdAjouter: TButton
      Left = 230
      Top = 10
      Width = 98
      Height = 33
      Caption = 'Ajouter'
      ModalResult = 1
      TabOrder = 0
    end
  end
  object pnlEditBanques: TPanel
    Left = 10
    Top = 471
    Width = 546
    Height = 95
    TabOrder = 2
    object Label1: TLabel
      Left = 10
      Top = 10
      Width = 38
      Height = 17
      Caption = 'Nom :'
    end
    object Label2: TLabel
      Left = 10
      Top = 52
      Width = 55
      Height = 17
      Caption = 'Chemin :'
    end
    object edtNomBanques: TEdit
      Left = 73
      Top = 10
      Width = 336
      Height = 21
      TabOrder = 0
    end
    object edtCheminLdap: TEdit
      Left = 73
      Top = 52
      Width = 462
      Height = 21
      TabOrder = 1
    end
    object cmdAjouterBanque: TButton
      Left = 418
      Top = 10
      Width = 33
      Height = 33
      Caption = '+'
      TabOrder = 2
      OnClick = cmdAjouterBanqueClick
    end
    object cmdEnregistrer: TButton
      Left = 460
      Top = 10
      Width = 33
      Height = 33
      Caption = 'S'
      TabOrder = 3
      OnClick = cmdEnregistrerClick
    end
    object cmdSupprimer: TButton
      Left = 502
      Top = 10
      Width = 33
      Height = 33
      Caption = '-'
      TabOrder = 4
      OnClick = cmdSupprimerClick
    end
  end
  object fichierBanques: TXMLDocument
    Options = [doNodeAutoCreate, doNodeAutoIndent, doAttrNull, doAutoPrefix, doNamespaceDecl]
    Left = 24
    Top = 232
    DOMVendorDesc = 'MSXML'
  end
end
