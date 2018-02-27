unit main;

interface

uses
  DragDrop, DropSource, DropTarget, DragDropFile, ActiveX,
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DropComboTarget, Registry, Menus, JvMenus, jpeg, JvGIF,
  dxGDIPlusClasses;

type
  TfrmMain = class(TForm)
    DropComboTarget: TDropComboTarget;
    Popup: TJvPopupMenu;
    LoginSetup1: TMenuItem;
    N2: TMenuItem;
    Close1: TMenuItem;
    N1: TMenuItem;
    Version11: TMenuItem;
    Whatsnew1: TMenuItem;
    Image1: TImage;
    procedure FormCreate(Sender: TObject);
    procedure DropComboTargetDrop(Sender: TObject;
      ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
    procedure FormDestroy(Sender: TObject);
    procedure LoginSetup1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PopupPopup(Sender: TObject);
    procedure Whatsnew1Click(Sender: TObject);
  private
    { Private declarations }
     tmpdir: string;
     FFileList: TStringList;
     FURL: boolean;

     procedure DoSave;
     procedure CleanUpEmails;
     procedure WMNCHitTest(var Msg: TWMNCHitTest) ; message WM_NCHitTest;

//     procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

uses
  DragDropFormats, ComObj, SaveDoc, SaveDocFunc,
  SavedocDetails, LoginDetails, whatsnew;

{
procedure TfrmMain.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
//  Params.Style := Params.Style AND NOT WS_CAPTION;
  with Params do begin
    ExStyle := ExStyle or WS_EX_TOPMOST or WS_EX_APPWINDOW;
    WndParent := GetDesktopwindow;
  end;
end;   }

procedure TfrmMain.WMNCHitTest(var Msg: TWMNCHitTest) ;
begin
   inherited;
   if Msg.Result = htClient then Msg.Result := htCaption;
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
   regAxiom: TRegistry;
   sRegistryRoot,
   DFilePath,
   DIconPath,
   DebugFile: string;
begin
   DIconPath := ExtractFilePath(Application.EXEName)+ 'dragon.ico';
   if FileExists(DIconPath) then
      Application.Icon.LoadFromFile(DIconPath);

   DFilePath := ExtractFilePath(Application.EXEName)+ 'dragon.jpg';
   if FileExists(DFilePath) then
      Image1.Picture.LoadFromFile(DFilePath);

   sRegistryRoot := 'Software\Colateral\Axiom';
   regAxiom := TRegistry.Create;
   try
      regAxiom.RootKey := HKEY_CURRENT_USER;
      if regAxiom.OpenKey(sRegistryRoot+'\InsightDocSave', False) then
      begin
         try
            if regAxiom.ReadInteger('WinTop') <> 0 then
               Self.Top := regAxiom.ReadInteger('WinTop');
         except
         //
         end;
         try
            if regAxiom.ReadInteger('WinLeft') <> 0 then
               Self.Left := regAxiom.ReadInteger('WinLeft');
         except
         //
         end;
      end;
   finally
      regAxiom.Free;
   end;
   width := 58;
   clientwidth := 58;
   height := 56;
   clientheight := 56;
   tmpdir := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP'));
end;

procedure TfrmMain.DropComboTargetDrop(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
var
  Stream: TStream;
  i: integer;
  Name, TempFile, FileID: string;
begin
   Self.Activate;
   FURL := False;
   // Extract dropped data.
//   MessageDlg('about to test dropped data',mtInformation,[mbOk], 0, mbOk);
   if DropComboTarget.Data.Count > 0 then
      FFileList := TStringList.Create;
   try
//      MessageDlg('total documents being dropped = '+inttostr(DropComboTarget.Data.Count),mtInformation,[mbOk], 0, mbOk);
      for i := 0 to DropComboTarget.Data.Count-1 do
      begin
          Name := DropComboTarget.Data.Names[i];
          if (Name = '') then
              Name := intToStr(i)+'.dat';
          Stream := TFileStream.Create(tmpdir + Name, fmCreate);
          try
              // Copy dropped data to stream (in this case a file stream).
             Stream.CopyFrom(DropComboTarget.Data[i], DropComboTarget.Data[i].Size);
          finally
             Stream.Free;
          end;
          TempFile := tmpdir + Name;
//          MessageDlg('filename = '+TempFile,mtInformation,[mbOk], 0, mbOk);
          FFileList.Add(TempFile);
      end;
   finally
      DoSave();
      CleanUpEmails();
      DropComboTarget.Data.Clear;
      DropComboTarget.Files.Clear;
      FFileList.Free;
      FFileList := nil;
   end;

   if (DropComboTarget.URL <> '') then
   begin
      try
         FFileList := TStringList.Create;
         FFileList.Add(DropComboTarget.URL);
         FURL := True;
         DoSave();
      finally
         FFileList.Free;
         FFileList := nil;
      end;
   end;

  // Copy the rest of the dropped formats.
   if (DropComboTarget.Files.Count > 0) then
   begin
      try
         FFileList := TStringList.Create;
         FFileList.Assign(DropComboTarget.Files);
         DoSave();
      finally
         FFileList.Free;
         FFileList := nil;
      end;
   end;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
var
  regAxiom: TRegistry;
  sRegistryRoot: string;
begin
   sRegistryRoot := 'Software\Colateral\Axiom';
   regAxiom := TRegistry.Create;
   try
      regAxiom.RootKey := HKEY_CURRENT_USER;
      if regAxiom.OpenKey(sRegistryRoot+'\InsightDocSave', True) then
      begin
         regAxiom.WriteInteger('WinTop',Self.Top);
         regAxiom.WriteInteger('WinLeft',Self.Left);
         regAxiom.CloseKey;
      end;
   finally
      regAxiom.Free;
   end;
end;

procedure TfrmMain.LoginSetup1Click(Sender: TObject);
begin
   try
      frmLoginSetup := TfrmLoginSetup.Create(Application);
      if frmLoginSetup.ShowModal = mrOk then
      begin
         try
            GetUserID;
         except
            Application.Terminate;
         end;
      end;
   finally
      frmLoginSetup.Free;
   end;
end;

procedure TfrmMain.Close1Click(Sender: TObject);
begin
   try
      dmSaveDoc.uniInsight.Disconnect;
      dmSaveDoc.Free;
   finally
      Application.Terminate;
   end;
end;

procedure TfrmMain.DoSave;
var
   i: integer;
begin
   if not FormExists(frmSaveDocDetails) then
      frmSaveDocDetails := TfrmSaveDocDetails.Create(Self);
   try
      if FFileList.Count > 0 then
      begin
         for i := 0 to (FFileList.Count - 1) do
         begin
            frmSaveDocDetails.DocName := FFileList.Strings[i];
            if FURL then
               frmSaveDocDetails.URLOnly := True
            else
               frmSaveDocDetails.URLOnly := False;
            frmSaveDocDetails.Repaint;
            frmSaveDocDetails.ShowModal;
         end;
      end;
   except
      frmSaveDocDetails.Release;
   end;
end;

procedure TfrmMain.Image1MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
   try
      if Button = mbLeft then
      begin
         ReleaseCapture;
         SendMessage(frmMain.Handle, WM_SYSCOMMAND, 61458, 0) ;
      end;
   except
   //
   end;
end;

procedure TfrmMain.PopupPopup(Sender: TObject);
begin
   Popup.Items[5].Caption := 'Version: ' + ReportVersion;
end;

procedure TfrmMain.CleanUpEmails;
var
   i: integer;
begin
   try
      if FFileList.Count > 0 then
      begin
         for i := 0 to (FFileList.Count - 1) do
            DeleteFile(FFileList.Strings[i]);
      end;
   except
      //
   end;
end;

procedure TfrmMain.Whatsnew1Click(Sender: TObject);
begin
   try
      frmWhatsNew := TfrmWhatsNew.Create(Application);
      frmWhatsNew.ShowModal;
   finally
      frmWhatsNew.Free;
   end;
end;

end.
