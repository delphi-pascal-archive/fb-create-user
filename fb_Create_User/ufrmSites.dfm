object frmSites: TfrmSites
  Left = 400
  Top = 86
  Width = 647
  Height = 697
  Caption = 'Liste des Entites'
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
  object Label1: TLabel
    Left = 21
    Top = 21
    Width = 54
    Height = 17
    Caption = 'Entites : '
  end
  object Label2: TLabel
    Left = 418
    Top = 31
    Width = 81
    Height = 17
    Caption = 'Nom Entite : '
  end
  object Label4: TLabel
    Left = 31
    Top = 230
    Width = 37
    Height = 17
    Caption = 'Sites :'
  end
  object Label5: TLabel
    Left = 435
    Top = 241
    Width = 64
    Height = 17
    Caption = 'Nom Site :'
  end
  object Label6: TLabel
    Left = 443
    Top = 272
    Width = 54
    Height = 17
    Caption = 'OU Site :'
  end
  object Label7: TLabel
    Left = 467
    Top = 303
    Width = 30
    Height = 17
    Caption = 'Ville :'
  end
  object Label8: TLabel
    Left = 429
    Top = 335
    Width = 69
    Height = 17
    Caption = 'Serveur P :'
  end
  object Label9: TLabel
    Left = 10
    Top = 450
    Width = 106
    Height = 17
    Caption = 'Groupes societe :'
  end
  object Label10: TLabel
    Left = 408
    Top = 450
    Width = 118
    Height = 17
    Caption = 'Groupes exterieur :'
  end
  object Label11: TLabel
    Left = 424
    Top = 63
    Width = 73
    Height = 17
    Caption = 'Profile TSE :'
  end
  object Label12: TLabel
    Left = 408
    Top = 94
    Width = 93
    Height = 17
    Caption = 'Nom Complet :'
  end
  object Label13: TLabel
    Left = 430
    Top = 126
    Width = 67
    Height = 17
    Caption = 'Password :'
  end
  object Label14: TLabel
    Left = 420
    Top = 157
    Width = 79
    Height = 17
    Caption = 'Login script :'
  end
  object Label15: TLabel
    Left = 441
    Top = 366
    Width = 56
    Height = 17
    Caption = 'Partage :'
  end
  object Label16: TLabel
    Left = 443
    Top = 429
    Width = 55
    Height = 17
    Caption = 'Lecteur :'
  end
  object Label17: TLabel
    Left = 424
    Top = 398
    Width = 72
    Height = 17
    Caption = 'Repertoire :'
  end
  object Label3: TLabel
    Left = 10
    Top = 659
    Width = 167
    Height = 17
    Caption = 'Banque exchange Societe :'
  end
  object Label18: TLabel
    Left = 408
    Top = 659
    Width = 177
    Height = 17
    Caption = 'Banque exchange exterieur :'
  end
  object cmdAjouterEntite: TButton
    Left = 84
    Top = 167
    Width = 98
    Height = 33
    Caption = 'Creer Nouvelle'
    TabOrder = 0
    OnClick = cmdAjouterEntiteClick
  end
  object cmdSupprimerEntite: TButton
    Left = 293
    Top = 167
    Width = 98
    Height = 33
    Caption = 'Supprimer'
    TabOrder = 1
    OnClick = cmdSupprimerEntiteClick
  end
  object lstEntite: TListBox
    Left = 84
    Top = 21
    Width = 315
    Height = 137
    ItemHeight = 17
    TabOrder = 2
    OnClick = lstEntiteClick
  end
  object edtNomEntite: TEdit
    Left = 502
    Top = 21
    Width = 305
    Height = 21
    TabOrder = 3
  end
  object lstSite: TListBox
    Left = 84
    Top = 230
    Width = 315
    Height = 137
    ItemHeight = 17
    TabOrder = 4
    OnClick = lstSiteClick
  end
  object cmdEnregistrerEntite: TButton
    Left = 188
    Top = 166
    Width = 98
    Height = 34
    Caption = 'Enregistrer'
    TabOrder = 5
    OnClick = cmdEnregistrerEntiteClick
  end
  object edtNomSite: TEdit
    Left = 502
    Top = 230
    Width = 305
    Height = 21
    TabOrder = 6
  end
  object edtOUSite: TEdit
    Left = 502
    Top = 262
    Width = 305
    Height = 21
    TabOrder = 7
  end
  object edtCity: TEdit
    Left = 502
    Top = 293
    Width = 305
    Height = 21
    TabOrder = 8
  end
  object edtSrvP: TEdit
    Left = 502
    Top = 324
    Width = 305
    Height = 21
    TabOrder = 9
  end
  object lstGroupeGroupe: TListBox
    Left = 84
    Top = 471
    Width = 315
    Height = 137
    ItemHeight = 17
    TabOrder = 10
  end
  object lstGroupeExterieur: TListBox
    Left = 492
    Top = 471
    Width = 315
    Height = 137
    ItemHeight = 17
    TabOrder = 11
  end
  object cmdAjouterSite: TButton
    Left = 84
    Top = 377
    Width = 98
    Height = 32
    Caption = 'Creer Nouveau'
    TabOrder = 12
    OnClick = cmdAjouterSiteClick
  end
  object cmdSupprimerSite: TButton
    Left = 293
    Top = 377
    Width = 98
    Height = 32
    Caption = 'Supprimer'
    TabOrder = 13
    OnClick = cmdSupprimerSiteClick
  end
  object cmdEnregistrerSite: TButton
    Left = 188
    Top = 375
    Width = 98
    Height = 34
    Caption = 'Enregistrer'
    TabOrder = 14
    OnClick = cmdEnregistrerSiteClick
  end
  object edtProfileTse: TEdit
    Left = 502
    Top = 52
    Width = 305
    Height = 21
    TabOrder = 15
  end
  object edtNomComplet: TEdit
    Left = 502
    Top = 84
    Width = 305
    Height = 21
    TabOrder = 16
  end
  object edtPassword: TEdit
    Left = 502
    Top = 115
    Width = 253
    Height = 21
    TabOrder = 17
  end
  object cbxNeverExpire: TCheckBox
    Left = 764
    Top = 115
    Width = 43
    Height = 22
    Hint = 'Password never expire'
    Caption = 'NE'
    TabOrder = 18
  end
  object cmdAjouterGrpGroupe: TButton
    Left = 84
    Top = 617
    Width = 32
    Height = 33
    Caption = '+'
    TabOrder = 19
    OnClick = cmdAouterGrproupeClick
  end
  object cmdAjouterGRPExterieur: TButton
    Left = 502
    Top = 616
    Width = 33
    Height = 34
    Caption = '+'
    TabOrder = 20
    OnClick = cmdAjouterGRPExterieurClick
  end
  object cmdSupprimerGRPGroupe: TButton
    Left = 126
    Top = 616
    Width = 32
    Height = 34
    Caption = '-'
    TabOrder = 21
    OnClick = cmdSupprimerGRPGroupeClick
  end
  object cmdSupprimerGRPExterieur: TButton
    Left = 544
    Top = 616
    Width = 33
    Height = 34
    Caption = '-'
    TabOrder = 22
    OnClick = cmdSupprimerGRPExterieurClick
  end
  object edtLoginScript: TEdit
    Left = 502
    Top = 146
    Width = 305
    Height = 21
    TabOrder = 23
  end
  object edtSRVPartage: TEdit
    Left = 502
    Top = 356
    Width = 305
    Height = 21
    TabOrder = 24
  end
  object edtLecteurP: TEdit
    Left = 502
    Top = 418
    Width = 43
    Height = 21
    TabOrder = 25
  end
  object edtsrvRepertoire: TEdit
    Left = 502
    Top = 387
    Width = 305
    Height = 21
    TabOrder = 26
  end
  object lstBanqueExchangeGroupe: TListBox
    Left = 10
    Top = 680
    Width = 389
    Height = 137
    ItemHeight = 17
    TabOrder = 27
  end
  object lstBanqueExchangeExterieur: TListBox
    Left = 408
    Top = 680
    Width = 399
    Height = 137
    ItemHeight = 17
    TabOrder = 28
  end
  object cmdSupprimerBanqueGroupe: TButton
    Left = 52
    Top = 825
    Width = 33
    Height = 34
    Caption = '-'
    TabOrder = 29
    OnClick = cmdSupprimerBanqueGroupeClick
  end
  object cmdAjouteBanqueGroupe: TButton
    Left = 10
    Top = 826
    Width = 33
    Height = 33
    Caption = '+'
    TabOrder = 30
    OnClick = cmdAjouteBanqueGroupeClick
  end
  object cmdSupprimerbanqueExterieur: TButton
    Left = 450
    Top = 825
    Width = 33
    Height = 34
    Caption = '-'
    TabOrder = 31
    OnClick = cmdSupprimerbanqueExterieurClick
  end
  object cmdAjouterBanqueExterieur: TButton
    Left = 408
    Top = 826
    Width = 33
    Height = 33
    Caption = '+'
    TabOrder = 32
    OnClick = cmdAjouterBanqueExterieurClick
  end
  object fichierSites: TXMLDocument
    Options = [doNodeAutoCreate, doNodeAutoIndent, doAttrNull, doAutoPrefix, doNamespaceDecl]
    Left = 128
    Top = 472
    DOMVendorDesc = 'MSXML'
  end
end
