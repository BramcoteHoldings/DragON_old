object DataModule2: TDataModule2
  OldCreateOrder = False
  Left = 366
  Top = 404
  Height = 232
  Width = 409
  object orsAxiom: TOraSession
    ConnectPrompt = False
    Options.NeverConnect = True
    Username = 'axiom'
    Password = 'axiom'
    Server = 'AXIOMNW'
    Left = 29
    Top = 9
  end
  object OraQuery1: TOraQuery
    Session = orsAxiom
    Left = 37
    Top = 70
  end
  object OraDataSource1: TOraDataSource
    Left = 83
    Top = 71
  end
end
