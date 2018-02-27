object frmLoginSetup: TfrmLoginSetup
  Left = 404
  Top = 275
  HorzScrollBar.Visible = False
  VertScrollBar.Visible = False
  BorderStyle = bsToolWindow
  Caption = 'Login Setup'
  ClientHeight = 237
  ClientWidth = 239
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnCreate = FormCreate
  DesignSize = (
    239
    237)
  PixelsPerInch = 106
  TextHeight = 15
  object GroupBox1: TGroupBox
    Left = 7
    Top = 8
    Width = 219
    Height = 195
    Caption = 'Login Details'
    TabOrder = 0
    object Label1: TLabel
      Left = 17
      Top = 138
      Width = 53
      Height = 15
      Caption = 'Username'
    end
    object Label2: TLabel
      Left = 17
      Top = 166
      Width = 50
      Height = 15
      Caption = 'Password'
    end
    object edUserName: TEdit
      Left = 79
      Top = 135
      Width = 121
      Height = 21
      AutoSize = False
      TabOrder = 1
    end
    object edPassword: TEdit
      Left = 79
      Top = 163
      Width = 121
      Height = 21
      AutoSize = False
      PasswordChar = '*'
      TabOrder = 2
    end
    object Database: TGroupBox
      Left = 6
      Top = 22
      Width = 207
      Height = 105
      Caption = 'Database Details'
      TabOrder = 0
      object Label3: TLabel
        Left = 5
        Top = 25
        Width = 32
        Height = 15
        Caption = 'Server'
      end
      object Label4: TLabel
        Left = 5
        Top = 51
        Width = 48
        Height = 15
        Caption = 'Database'
      end
      object Label5: TLabel
        Left = 5
        Top = 80
        Width = 22
        Height = 15
        Caption = 'Port'
      end
      object edServerName: TEdit
        Left = 57
        Top = 22
        Width = 146
        Height = 21
        AutoSize = False
        TabOrder = 0
      end
      object edDatabase: TEdit
        Left = 57
        Top = 48
        Width = 121
        Height = 23
        TabOrder = 1
      end
      object edPort: TEdit
        Left = 57
        Top = 77
        Width = 57
        Height = 23
        TabOrder = 2
        Text = '1521'
      end
    end
  end
  object Button1: TButton
    Left = 151
    Top = 208
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = 'Save'
    TabOrder = 1
    OnClick = Button1Click
  end
end
