program DragON;

uses
  Forms,
  main in 'main.pas' {frmMain},
  LoginDetails in 'LoginDetails.pas' {frmLoginSetup},
  MatterSearch in 'MatterSearch.pas' {frmMtrSearch},
  SavedocDetails in 'SavedocDetails.pas' {frmSaveDocDetails},
  SaveDocFunc in 'SaveDocFunc.pas',
  SaveDoc in 'SaveDoc.pas' {dmSaveDoc: TDataModule},
  whatsnew in 'whatsnew.pas' {frmWhatsNew},
  Vcl.Themes,
  Vcl.Styles;

{$R *.res}



begin
  Application.Initialize;
  TStyleManager.TrySetStyle('Iceberg Classico');
  Application.Title := 'DragON';
  Application.CreateForm(TdmSaveDoc, dmSaveDoc);
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
