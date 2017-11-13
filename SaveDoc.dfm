object dmSaveDoc: TdmSaveDoc
  OldCreateOrder = False
  Height = 405
  Width = 370
  object uniInsight: TUniConnection
    ProviderName = 'Oracle'
    SpecificOptions.Strings = (
      'Oracle.Direct=True'
      'Oracle.IPVersion=ivIPBoth')
    Options.ConvertEOL = True
    Options.DisconnectedMode = True
    Username = 'axiom'
    Server = 'orion:1521:marketng'
    LoginPrompt = False
    OnError = uniInsightError
    Left = 25
    Top = 8
    EncryptedPassword = '9EFF87FF96FF90FF92FF'
  end
  object qryEmps: TUniQuery
    Connection = uniInsight
    Left = 181
    Top = 4
  end
  object qryMatters: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'select * from matter'
      'where closed = 0 and entity = :P_Entity')
    Left = 32
    Top = 65
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'P_Entity'
        Value = nil
      end>
  end
  object dsMatters: TUniDataSource
    DataSet = qryMatters
    Left = 136
    Top = 256
  end
  object qryGetSeq: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'select DOC_DOCID.nextval as nextdoc from dual')
    Left = 95
    Top = 5
  end
  object qryMatterAttachments: TUniQuery
    UpdatingTable = 'DOC'
    Connection = uniInsight
    SQL.Strings = (
      'SELECT'
      '  DOC.DOCUMENT,'
      '  DOC.IMAGEINDEX,'
      '  DOC.FILE_EXTENSION,'
      '  DOC.DOC_NAME,'
      '  DOC.SEARCH,'
      '  DOC.DOC_CODE,'
      '  DOC.JURIS,'
      '  DOC.D_CREATE,'
      '  DOC.AUTH1,'
      '  DOC.D_MODIF,'
      '  DOC.AUTH2,'
      '  DOC.PATH,'
      '  DOC.DESCR,'
      '  DOC.FILEID,'
      '  DOC.DOCID,'
      '  DOC.NPRECCATEGORY,'
      '  DOC.NMATTER,'
      '  DOC.PRECEDENT_DETAILS,'
      '  DOC.NPRECCLASSIFICATION,'
      '  DOC.KEYWORDS,'
      '  DOC.URL,'
      '  DOC.DISPLAY_PATH,'
      '  DOC.EMAIL_SENT_TO,'
      '  DOC.PARENTDOCID,'
      '  DOC.D_CREATE,'
      '  DOC.ROWID'
      'FROM'
      '  DOC'
      'where'
      '  DOCID = :DOCID')
    CachedUpdates = True
    OnNewRecord = qryMatterAttachmentsNewRecord
    Left = 45
    Top = 127
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'DOCID'
        Value = nil
      end>
  end
  object qryGetMatter: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'select fileid, nmatter '
      'from'
      'matter'
      'where'
      'fileid = :fileid')
    Left = 137
    Top = 194
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'fileid'
        Value = nil
      end>
  end
  object qryGetEntity: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'SELECT VALUE,INTVALUE'
      'FROM SETTINGS '
      'WHERE EMP = :Emp'
      '  AND OWNER = :Owner'
      '  AND ITEM = :Item')
    Left = 25
    Top = 183
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'Emp'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'Owner'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'Item'
        Value = nil
      end>
  end
  object qryPrecCategory: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'select * from PRECCATEGORY')
    Left = 189
    Top = 63
  end
  object dsPrecCategory: TUniDataSource
    DataSet = qryPrecCategory
    Left = 191
    Top = 111
  end
  object qryTmp: TUniQuery
    Connection = uniInsight
    Left = 289
    Top = 134
  end
  object qrySysFile: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'SELECT * FROM SYSTEMFILE')
    Left = 108
    Top = 78
  end
  object procTemp: TUniStoredProc
    Connection = uniInsight
    Left = 28
    Top = 239
  end
  object qryPrecClassification: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'select * from PRECCLASSIFICATION')
    Left = 266
    Top = 251
  end
  object dsPrecClassification: TUniDataSource
    DataSet = qryPrecClassification
    Left = 271
    Top = 300
  end
  object dsEmployee: TUniDataSource
    DataSet = qryEmployee
    Left = 300
    Top = 59
  end
  object qryEmployee: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'select code, name from employee where active = '#39'Y'#39' order by code')
    Left = 295
    Top = 7
  end
  object qrySaveEmailAttachments: TUniQuery
    Connection = uniInsight
    SQL.Strings = (
      'SELECT'
      '  DOC.DOCUMENT,'
      '  DOC.IMAGEINDEX,'
      '  DOC.FILE_EXTENSION,'
      '  DOC.DOC_NAME,'
      '  DOC.SEARCH,'
      '  DOC.DOC_CODE,'
      '  DOC.JURIS,'
      '  DOC.D_CREATE,'
      '  DOC.AUTH1,'
      '  DOC.D_MODIF,'
      '  DOC.AUTH2,'
      '  DOC.PATH,'
      '  DOC.DESCR,'
      '  DOC.FILEID,'
      '  DOC.DOCID,'
      '  DOC.NPRECCATEGORY,'
      '  DOC.NMATTER,'
      '  DOC.PRECEDENT_DETAILS,'
      '  DOC.NPRECCLASSIFICATION,'
      '  DOC.KEYWORDS,'
      '  DOC.URL,'
      '  DOC.DISPLAY_PATH,'
      '  DOC.PARENTDOCID,'
      '  DOC.IS_ATTACHMENT,'
      '  DOC.ROWID'
      'FROM'
      '  DOC'
      'where'
      '  DOCID = :DOCID')
    CachedUpdates = True
    OnNewRecord = qrySaveEmailAttachmentsNewRecord
    Left = 279
    Top = 192
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'DOCID'
        Value = nil
      end>
  end
end
