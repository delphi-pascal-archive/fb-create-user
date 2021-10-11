object frmUser: TfrmUser
  Left = 363
  Top = 94
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  ClientHeight = 667
  ClientWidth = 794
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -14
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 120
  TextHeight = 17
  object GrpOptionel: TGroupBox
    Left = 366
    Top = 293
    Width = 420
    Height = 127
    Caption = 'Informations Optionels'
    TabOrder = 0
    object Label6: TLabel
      Left = 59
      Top = 31
      Width = 46
      Height = 17
      Caption = 'N'#176' Tel :'
    end
    object Label8: TLabel
      Left = 54
      Top = 63
      Width = 51
      Height = 17
      Caption = 'N'#176' Fax :'
    end
    object Label10: TLabel
      Left = 14
      Top = 94
      Width = 90
      Height = 17
      Caption = 'Departement :'
    end
    object edtTel: TEdit
      Left = 115
      Top = 21
      Width = 158
      Height = 21
      TabOrder = 0
    end
    object edtFax: TEdit
      Left = 115
      Top = 52
      Width = 158
      Height = 21
      TabOrder = 1
    end
    object edtDepartement: TEdit
      Left = 115
      Top = 84
      Width = 200
      Height = 21
      TabOrder = 2
    end
  end
  object grpObligaatoire: TGroupBox
    Left = 10
    Top = 10
    Width = 347
    Height = 609
    Caption = 'Informations Obligatoires'
    TabOrder = 1
    object Label1: TLabel
      Left = 67
      Top = 31
      Width = 38
      Height = 17
      Caption = 'Nom :'
    end
    object Label2: TLabel
      Left = 47
      Top = 63
      Width = 57
      Height = 17
      Caption = 'Prenom :'
    end
    object Label3: TLabel
      Left = 71
      Top = 94
      Width = 32
      Height = 17
      Caption = 'UID :'
    end
    object Label4: TLabel
      Left = 61
      Top = 126
      Width = 42
      Height = 17
      Caption = 'Login :'
    end
    object Label7: TLabel
      Left = 71
      Top = 220
      Width = 31
      Height = 17
      Caption = 'Site :'
    end
    object Label9: TLabel
      Left = 5
      Top = 188
      Width = 94
      Height = 17
      Caption = 'Appartenance :'
    end
    object Label11: TLabel
      Left = 10
      Top = 356
      Width = 60
      Height = 17
      Caption = 'Groupes :'
    end
    object Label5: TLabel
      Left = 54
      Top = 157
      Width = 47
      Height = 17
      Caption = '@ Mail :'
    end
    object Label18: TLabel
      Left = 18
      Top = 251
      Width = 84
      Height = 17
      Caption = 'Nom societe :'
    end
    object edtNom: TEdit
      Left = 115
      Top = 21
      Width = 221
      Height = 21
      TabOrder = 0
      OnChange = edtChange
    end
    object edtPrenom: TEdit
      Left = 115
      Top = 52
      Width = 221
      Height = 21
      TabOrder = 1
      OnChange = edtChange
    end
    object edtUID: TEdit
      Left = 115
      Top = 84
      Width = 221
      Height = 21
      TabOrder = 2
    end
    object edtLogin: TEdit
      Left = 115
      Top = 115
      Width = 221
      Height = 21
      TabOrder = 3
      OnChange = edtLoginChange
    end
    object cbxLocalisation: TComboBox
      Left = 115
      Top = 209
      Width = 221
      Height = 25
      ItemHeight = 17
      TabOrder = 6
      OnClick = edtChange
    end
    object cbxApartenance: TComboBox
      Left = 115
      Top = 178
      Width = 221
      Height = 25
      ItemHeight = 17
      TabOrder = 5
      OnClick = cbxApartenanceClick
    end
    object lstGroupe: TListBox
      Left = 10
      Top = 377
      Width = 326
      Height = 179
      ItemHeight = 17
      TabOrder = 9
    end
    object cmdAddGrp: TBitBtn
      Left = 10
      Top = 565
      Width = 99
      Height = 33
      Caption = 'Ajouter'
      Default = True
      TabOrder = 10
      OnClick = cmdAddGrpClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        3333333333333333333333330000333333333333333333333333F33333333333
        00003333344333333333333333388F3333333333000033334224333333333333
        338338F3333333330000333422224333333333333833338F3333333300003342
        222224333333333383333338F3333333000034222A22224333333338F338F333
        8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
        33333338F83338F338F33333000033A33333A222433333338333338F338F3333
        0000333333333A222433333333333338F338F33300003333333333A222433333
        333333338F338F33000033333333333A222433333333333338F338F300003333
        33333333A222433333333333338F338F00003333333333333A22433333333333
        3338F38F000033333333333333A223333333333333338F830000333333333333
        333A333333333333333338330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object cmdSupprimer: TBitBtn
      Left = 241
      Top = 565
      Width = 98
      Height = 33
      Cancel = True
      Caption = 'Supprimer'
      TabOrder = 11
      OnClick = cmdSupprimerClick
      Glyph.Data = {
        DE010000424DDE01000000000000760000002800000024000000120000000100
        0400000000006801000000000000000000001000000000000000000000000000
        80000080000000808000800000008000800080800000C0C0C000808080000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333000033338833333333333333333F333333333333
        0000333911833333983333333388F333333F3333000033391118333911833333
        38F38F333F88F33300003339111183911118333338F338F3F8338F3300003333
        911118111118333338F3338F833338F3000033333911111111833333338F3338
        3333F8330000333333911111183333333338F333333F83330000333333311111
        8333333333338F3333383333000033333339111183333333333338F333833333
        00003333339111118333333333333833338F3333000033333911181118333333
        33338333338F333300003333911183911183333333383338F338F33300003333
        9118333911183333338F33838F338F33000033333913333391113333338FF833
        38F338F300003333333333333919333333388333338FFF830000333333333333
        3333333333333333333888330000333333333333333333333333333333333333
        0000}
      NumGlyphs = 2
    end
    object edtMail: TEdit
      Left = 115
      Top = 146
      Width = 221
      Height = 21
      TabOrder = 4
    end
    object edtSociete: TEdit
      Left = 115
      Top = 241
      Width = 221
      Height = 21
      TabOrder = 7
      OnChange = edtChange
    end
    object grbDate: TGroupBox
      Left = 105
      Top = 282
      Width = 231
      Height = 65
      TabOrder = 8
      object Label13: TLabel
        Left = 9
        Top = 37
        Width = 82
        Height = 17
        Caption = 'Date depart :'
      end
      object dtDateDepart: TDateTimePicker
        Left = 105
        Top = 26
        Width = 116
        Height = 25
        Date = 39738.527067291670000000
        Time = 39738.527067291670000000
        TabOrder = 1
        OnChange = edtChange
      end
      object cbxDuree: TCheckBox
        Left = 10
        Top = 0
        Width = 138
        Height = 22
        Caption = 'Duree determinee'
        TabOrder = 0
        OnClick = cbxDureeClick
      end
    end
  end
  object grbAutogenere: TGroupBox
    Left = 366
    Top = 10
    Width = 420
    Height = 274
    Caption = 'Champ auto'
    TabOrder = 2
    object Label14: TLabel
      Left = 27
      Top = 126
      Width = 73
      Height = 17
      Caption = 'Profile TSE :'
    end
    object Label12: TLabel
      Left = 31
      Top = 94
      Width = 71
      Height = 17
      Caption = 'PassWord :'
    end
    object Label15: TLabel
      Left = 71
      Top = 157
      Width = 30
      Height = 17
      Caption = 'Ville :'
    end
    object Label16: TLabel
      Left = 12
      Top = 31
      Width = 93
      Height = 17
      Caption = 'Nom Complet :'
    end
    object Label17: TLabel
      Left = 25
      Top = 63
      Width = 77
      Height = 17
      Caption = 'Description :'
    end
    object Label19: TLabel
      Left = 25
      Top = 188
      Width = 78
      Height = 17
      Caption = 'Compagnie :'
    end
    object Label20: TLabel
      Left = 17
      Top = 220
      Width = 87
      Height = 17
      Caption = 'Logon Script :'
    end
    object edtProfileTse: TEdit
      Left = 105
      Top = 115
      Width = 304
      Height = 21
      TabOrder = 3
    end
    object edtPassword: TEdit
      Left = 105
      Top = 84
      Width = 189
      Height = 21
      TabOrder = 2
    end
    object edtVille: TEdit
      Left = 105
      Top = 146
      Width = 221
      Height = 21
      TabOrder = 4
    end
    object edtNomComplet: TEdit
      Left = 105
      Top = 21
      Width = 304
      Height = 21
      TabOrder = 0
    end
    object edtDescription: TEdit
      Left = 105
      Top = 52
      Width = 304
      Height = 21
      TabOrder = 1
    end
    object edtCompagnie: TEdit
      Left = 105
      Top = 178
      Width = 221
      Height = 21
      TabOrder = 5
    end
    object edtLogonScript: TEdit
      Left = 105
      Top = 209
      Width = 221
      Height = 21
      TabOrder = 6
    end
    object cbxP: TCheckBox
      Left = 105
      Top = 241
      Width = 106
      Height = 22
      Caption = 'Creation P'
      Checked = True
      State = cbChecked
      TabOrder = 7
    end
    object cbxNeverExpire: TCheckBox
      Left = 303
      Top = 84
      Width = 114
      Height = 22
      Caption = 'Never Expire'
      TabOrder = 8
    end
  end
  object cmdCreer: TBitBtn
    Left = 209
    Top = 628
    Width = 148
    Height = 32
    Caption = 'OK'
    Default = True
    TabOrder = 3
    OnClick = cmdCreerClick
    Glyph.Data = {
      DE010000424DDE01000000000000760000002800000024000000120000000100
      0400000000006801000000000000000000001000000000000000000000000000
      80000080000000808000800000008000800080800000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      3333333333333333333333330000333333333333333333333333F33333333333
      00003333344333333333333333388F3333333333000033334224333333333333
      338338F3333333330000333422224333333333333833338F3333333300003342
      222224333333333383333338F3333333000034222A22224333333338F338F333
      8F33333300003222A3A2224333333338F3838F338F33333300003A2A333A2224
      33333338F83338F338F33333000033A33333A222433333338333338F338F3333
      0000333333333A222433333333333338F338F33300003333333333A222433333
      333333338F338F33000033333333333A222433333333333338F338F300003333
      33333333A222433333333333338F338F00003333333333333A22433333333333
      3338F38F000033333333333333A223333333333333338F830000333333333333
      333A333333333333333338330000333333333333333333333333333333333333
      0000}
    NumGlyphs = 2
  end
  object cmdAnnuler: TBitBtn
    Left = 366
    Top = 628
    Width = 148
    Height = 32
    Caption = 'Annuler'
    Default = True
    ModalResult = 2
    TabOrder = 4
    Glyph.Data = {
      DE010000424DDE01000000000000760000002800000024000000120000000100
      0400000000006801000000000000000000001000000000000000000000000000
      80000080000000808000800000008000800080800000C0C0C000808080000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
      333333333333333333333333000033338833333333333333333F333333333333
      0000333911833333983333333388F333333F3333000033391118333911833333
      38F38F333F88F33300003339111183911118333338F338F3F8338F3300003333
      911118111118333338F3338F833338F3000033333911111111833333338F3338
      3333F8330000333333911111183333333338F333333F83330000333333311111
      8333333333338F3333383333000033333339111183333333333338F333833333
      00003333339111118333333333333833338F3333000033333911181118333333
      33338333338F333300003333911183911183333333383338F338F33300003333
      9118333911183333338F33838F338F33000033333913333391113333338FF833
      38F338F300003333333333333919333333388333338FFF830000333333333333
      3333333333333333333888330000333333333333333333333333333333333333
      0000}
    NumGlyphs = 2
  end
  object GroupBox1: TGroupBox
    Left = 366
    Top = 429
    Width = 420
    Height = 74
    Caption = 'Banque Messagerie'
    TabOrder = 5
    object cbxBanque: TComboBox
      Left = 10
      Top = 31
      Width = 399
      Height = 25
      ItemHeight = 17
      TabOrder = 0
    end
  end
end
