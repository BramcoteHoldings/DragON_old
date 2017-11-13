object frmWhatsNew: TfrmWhatsNew
  Left = 509
  Top = 293
  BorderStyle = bsDialog
  Caption = 'Whats new in DragON'
  ClientHeight = 499
  ClientWidth = 570
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 15
  object Memo1: TMemo
    Left = 0
    Top = 0
    Width = 570
    Height = 499
    Align = alClient
    Lines.Strings = (
      'What'#39's new'
      '=========='
      ''
      'Build 40'
      '------------'
      '- Changes to email saving for Office 2016.'
      ''
      'Build 39'
      '------------'
      
        '- If document source and target are the same, the document is no' +
        't moved but a'
      '  record is added to Insight.'
      '  '
      ''
      'Build 38'
      '------------'
      
        '- Now appends document sequence to attachment name to prevent ov' +
        'erwriting of '
      '  documents with same name.'
      ''
      '- Author now defaults to user logged in.'
      ''
      ''
      'Build 37'
      '------------'
      
        '- Now appends document sequence to document name to prevent over' +
        'writing of '
      '  documents with same name.'
      ''
      'Build 36'
      '------------'
      
        '- Ability to change the background image.  To change the backgro' +
        'und image, '
      
        '  save a JPG image named "DragOn.jpg" in the same folder that Dr' +
        'agon '
      
        '  is executed from.  The same can be done for the icon.  Save an' +
        ' icon image named '
      '  "DragOn.ico" in the same folder that Dragon is executed from.'
      ''
      ''
      'Build 35'
      '------------'
      '- Fixed focusing issue when initiating drag drop operation.'
      ''
      ''
      'Version 3.0.1 Build 34'
      '----------------------------------'
      
        '- Change to behaviour when prompted that file already exists mes' +
        'sage comes up.'
      ''
      ''
      'Version 3.0.1 Build 33'
      '-----------------------------------'
      '- Further enhancements to validation when saving files.'
      ''
      ''
      'Version 3.0.1 Build 32'
      '----------------------------------'
      
        '- added checkbx to enable whether to save email attachments as s' +
        'eparate files'
      '  (if any)'
      ''
      ''
      'Version 3.0.1 Build 31'
      '----------------------------------'
      
        '- enhanced validation for saving files.  Will display common err' +
        'ors otherwise will '
      '  display error number.'
      ''
      
        '- ability to save attachments from email.  Will use same profile' +
        ' details as email.  '
      
        '  NOTE: If email has images as part of content,  these will be s' +
        'aved as attachments.'
      '  Original attachments will remain with original email.')
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 0
  end
end
