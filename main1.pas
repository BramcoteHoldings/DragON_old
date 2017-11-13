unit main;

interface

uses
  MapiDefs, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DragDrop, DropTarget, DragDropFile, DropHandler,
  DropSource,  ComCtrls, StdCtrls, DragDropText,
  LMDCustomControl, LMDCustomPanel, LMDCustomBevelPanel, LMDSimplePanel,
  JvMenus, LMDControl, Registry, jpeg, ImgList, Menus;

//const
//     NPR_FILEID: TRwMapiNamedProperty = (PropSetID: '{001b04db-360a-424e-ae80-3f1fce8c7458}'; PropID: $8000; PropName: 'NPR_FILEID'; PropType: PT_STRING8; PropKind: MNID_ID);


type
  TVirtualFileDataFormat = class(TCustomDataFormat)
  private
    FContents: AnsiString;
    FFileName: string;
  public
    constructor Create(AOwner: TDragDropComponent); override;
    function Assign(Source: TClipboardFormat): boolean; override;
    function AssignTo(Dest: TClipboardFormat): boolean; override;
    procedure Clear; override;
    function HasData: boolean; override;
    function NeedsData: boolean; override;
    property FileName: string read FFileName write FFileName;
    property Contents: AnsiString read FContents write FContents;
  end;


  TfrmMain = class(TForm)
    ImageList1: TImageList;
    ImageListSmall: TImageList;
    ImageListBig: TImageList;
    DropTextTarget1: TDropTextTarget;
    DataFormatAdapterFile: TDataFormatAdapter;
    DataFormatAdapterURL: TDataFormatAdapter;
    DataFormatAdapterOutlook: TDataFormatAdapter;
    DropEmptyTarget1: TDropEmptyTarget;
    Popup: TJvPopupMenu;
    LoginSetup1: TMenuItem;
    N2: TMenuItem;
    Close1: TMenuItem;
    LMDSimplePanel1: TLMDSimplePanel;
    Image1: TImage;
    ListViewBrowser: TListView;
    procedure FormCreate(Sender: TObject);
    procedure DropEmptyTarget1Drop(Sender: TObject;
      ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
    procedure Image1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure LoginSetup1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }

     FCleanUpList: TStringList;
     FOwnedMessage: TMessage;
     FFileList: TStringList;

     FEditFileName: string;
     FSourceDataFormat: TVirtualFileDataFormat;
     FTargetDataFormat: TVirtualFileDataFormat;
     originalPanelWindowProc : TWndMethod;
     procedure Reset;
     procedure CleanUp;
     procedure ResetView;
     procedure WMNCHitTest(var Msg: TWMNCHitTest) ; message WM_NCHitTest;
     procedure DoSave;
     procedure CreateParams(var Params: TCreateParams); override;
     procedure ReadBodyText(const AMessage: IMessage);
  public
    { Public declarations }
    function GetSender(const AMessage: IMessage): string;
    function GetSubject(const AMessage: IMessage): string;
    function GetFileID(const AMessage: IMessage): string;
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}

uses
  ShellAPI, SavedocDetails, DragDropFormats,
  ShlObj, ComObj, ActiveX, DragDropInternet,
  LoginDetails, SaveDocFunc, SaveDoc, MapiUtil, MapiTags;

type
  TMAPIINIT_0 =
    record
      Version: ULONG;
      Flags: ULONG;
    end;

  PMAPIINIT_0 = ^TMAPIINIT_0;
  TMAPIINIT = TMAPIINIT_0;
  PMAPIINIT = ^TMAPIINIT;

const
  MAPI_INIT_VERSION = 0;
  MAPI_MULTITHREAD_NOTIFICATIONS = $00000001;
  MAPI_NO_COINIT = $00000008;
  NPR_FILEID = $8000;

var
  MapiInit: TMAPIINIT_0 = (Version: MAPI_INIT_VERSION; Flags: 0);

procedure TfrmMain.WMNCHitTest(var Msg: TWMNCHitTest) ;
begin
   inherited;
   if Msg.Result = htClient then Msg.Result := htCaption;
end;

procedure TfrmMain.CleanUp;
var
  i: integer;
begin
  for i := 0 to FCleanUpList.Count-1 do
    try
      DeleteFile(FCleanUpList[i]);
    except
      // Ignore errors - nothing we can do about it anyway.
    end;
  ListViewBrowser.Clear;
  FCleanUpList.Clear;
end;

procedure TfrmMain.Reset;
begin
  ResetView;
end;

procedure TfrmMain.ResetView;
begin
   try
      FFileList.Clear;
   except
     //
   end;
end;


procedure TfrmMain.FormCreate(Sender: TObject);
var
  SHFileInfo: TSHFileInfo;
  regAxiom: TRegistry;
  sRegistryRoot: string;
begin
   try
    // It appears that for for Win XP and later it is OK to let MAPI call
    // coInitialize.
    // V5.1 = WinXP.
//    if ((Win32MajorVersion shl 16) or Win32MinorVersion < $00050001) then
//      MapiInit.Flags := MapiInit.Flags or MAPI_NO_COINIT;

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

      width := 36;
      clientwidth := 36;
      height := 36;
      clientheight := 36;
      OleCheck(MAPIInitialize(@MapiInit));
   except
      on E: Exception do
      ShowMessage(Format('Failed to initialize MAPI: %s', [E.Message]));
   end;

   // FCleanUpList contains a list of temporary files that should be deleted
   // before the application exits.
   FCleanUpList := TStringList.Create;

   // Get the system image list to use for the attachment listview.
   ImageListSmall.Handle := SHGetFileInfo('', 0, SHFileInfo, sizeOf(SHFileInfo),
      SHGFI_SYSICONINDEX or SHGFI_SMALLICON);
   ImageListBig.Handle := SHGetFileInfo('', 0, SHFileInfo, sizeOf(SHFileInfo),
      SHGFI_SYSICONINDEX or SHGFI_LARGEICON);

   FFileList := TStringList.Create();

//  FSourceDataFormat := TVirtualFileDataFormat.Create(DropEmptySource1);
   FTargetDataFormat := TVirtualFileDataFormat.Create(DropEmptyTarget1);
end;

{ TVirtualFileDataFormat }

constructor TVirtualFileDataFormat.Create(AOwner: TDragDropComponent);
begin
  inherited Create(AOwner);

  // Add the "file group descriptor" and "file contents" clipboard formats to
  // the data format's list of compatible formats.
  // Note: This is normally done via TCustomDataFormat.RegisterCompatibleFormat,
  // but since this data format is only used here, it is just as easy for us to
  // add the formats manually.
  CompatibleFormats.Add(TFileContentsClipboardFormat.Create);
  CompatibleFormats.Add(TAnsiFileGroupDescriptorClipboardFormat.Create);
  CompatibleFormats.Add(TUnicodeFileGroupDescriptorClipboardFormat.Create);
end;

function TVirtualFileDataFormat.Assign(Source: TClipboardFormat): boolean;
begin
  Result := True;

  (*
  ** TFileContentsClipboardFormat
  *)
  if (Source is TFileContentsClipboardFormat) then
  begin
    FContents := TFileContentsClipboardFormat(Source).Data
  end else
  (*
  ** TAnsiFileGroupDescriptorClipboardFormat
  *)
  if (Source is TAnsiFileGroupDescriptorClipboardFormat) then
  begin
    if (TAnsiFileGroupDescriptorClipboardFormat(Source).Count > 0) then
      FFileName := TAnsiFileGroupDescriptorClipboardFormat(Source).Filenames[0];
  end else
  (*
  ** TUnicodeFileGroupDescriptorClipboardFormat
  *)
//   if (Source is TOutlookDataFormat <> nil) then

  if (Source is TUnicodeFileGroupDescriptorClipboardFormat) then
  begin
    if (TUnicodeFileGroupDescriptorClipboardFormat(Source).Count > 0) then
      FFileName := TUnicodeFileGroupDescriptorClipboardFormat(Source).Filenames[0];
  end else
  (*
  ** None of the above...
  *)
    Result := inherited Assign(Source);
end;

function TVirtualFileDataFormat.AssignTo(Dest: TClipboardFormat): boolean;
begin
  Result := True;

  (*
  ** TFileContentsClipboardFormat
  *)
  if (Dest is TFileContentsClipboardFormat) then
  begin
    TFileContentsClipboardFormat(Dest).Data := FContents;
  end else
  (*
  ** TAnsiFileGroupDescriptorClipboardFormat
  *)
  if (Dest is TAnsiFileGroupDescriptorClipboardFormat) then
  begin
    TAnsiFileGroupDescriptorClipboardFormat(Dest).Count := 1;
    TAnsiFileGroupDescriptorClipboardFormat(Dest).Filenames[0] := FFileName;
  end else
  (*
  ** TUnicodeFileGroupDescriptorClipboardFormat
  *)
  if (Dest is TUnicodeFileGroupDescriptorClipboardFormat) then
  begin
    TUnicodeFileGroupDescriptorClipboardFormat(Dest).Count := 1;
    TUnicodeFileGroupDescriptorClipboardFormat(Dest).Filenames[0] := FFileName;
  end else
  (*
  ** None of the above...
  *)
    Result := inherited AssignTo(Dest);
end;

procedure TVirtualFileDataFormat.Clear;
begin
  FFileName := '';
  FContents := ''
end;

function TVirtualFileDataFormat.HasData: boolean;
begin
  Result := (FFileName <> '');
end;

function TVirtualFileDataFormat.NeedsData: boolean;
begin
  Result := (FFileName = '') or (FContents = '');
end;


procedure TfrmMain.DropEmptyTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
var
  OutlookDataFormat: TOutlookDataFormat;
  i, x, y: integer;
  Item: TListItem;
  AMessage: IMessage;
  tmpdir: string;
  Outlook, MailItem: OLEVariant;
  outlookSel: Variant;
  FileID, sSubject, TempFile: string;
begin
   if (DataFormatAdapterFile.DataFormat <> nil) and DataFormatAdapterFile.DataFormat.HasData then
   begin
      FFileList := TStringList.Create;
    // ...Extract the dropped data from it.
//    MemoFile.Lines.Assign((DataFormatAdapterFile.DataFormat as TFileDataFormat).Files);
      FFileList.Assign((DataFormatAdapterFile.DataFormat as TFileDataFormat).Files);
   end;

   if (DataFormatAdapterURL.DataFormat <> nil) and DataFormatAdapterURL.DataFormat.HasData then
//    MemoURL.Lines.Text := (DataFormatAdapterURL.DataFormat as TURLDataFormat).URL;
     frmSaveDocDetails.ShowModal;

  // Check if we have a data format and if so...
   if (DataFormatAdapterOutlook.DataFormat <> nil) and DataFormatAdapterOutlook.DataFormat.HasData then
   begin
    // ...Extract the dropped data from it.
      OutlookDataFormat := DataFormatAdapterOutlook.DataFormat as TOutlookDataFormat;
    (*
    ** Reset everything
    *)
      Reset;

      CleanUp;
      if not FormExists(frmSaveDocDetails) then
         frmSaveDocDetails := TfrmSaveDocDetails.Create(Self);

      tmpdir := GetEnvironmentVariable('TMP')+'\';
      // Get all the dropped messages
      for i := 1 to OutlookDataFormat.Messages.Count do
      begin
         // Get an IMessage interface
         if (Supports(OutlookDataFormat.Messages[i], IMessage, AMessage)) then
         begin
            try
               TempFile := '';
{               ReadBodyText(AMessage);
               Item := ListViewBrowser.Items.Add;
               Item.Caption := GetSender(AMessage);
               Item.SubItems.Add(GetSubject(AMessage));
               Item.Data := TMessage.Create(AMessage);
             }
              try
                  Outlook := GetActiveOleObject('Outlook.Application');
               except
                  Outlook := CreateOleObject('Outlook.Application');
               end;
               outlookSel := Outlook.ActiveExplorer.Selection;

               MailItem := Outlook.ActiveExplorer.Selection[i];
               TempFile := tmpdir+ListViewBrowser.Items[i].SubItems.Strings[0]+'.msg';
               MailItem.SaveAs(TempFile);

               try
                  FileID := GetFileID(TMessage(ListViewBrowser.Items[i].Data).Msg);
               except
                  FileID := '';
               end;

               if FileID = '' then
               begin
                  sSubject := trimleft(ListViewBrowser.Items[i].SubItems.Strings[0]);
                  for y := 1 to length(sSubject) do
                  begin
                     if sSubject[y] = '#' then
                     begin
                        for x := y + 1 to length(sSubject) do
                        begin
                           if (sSubject[x] <> ' ') and (sSubject[x] <> ']') then
                              FileID := FileID + sSubject[x];
                        end;
                     end;
                  end;
               end;
//               frmSaveDocDetails.EMessage := TMessage(ListViewBrowser.Items[i].Data).Msg;
               frmSaveDocDetails.DocName := TempFile;  // trimleft(ListViewBrowser.Items[i].SubItems.Strings[0]);
               frmSaveDocDetails.FileID := FileID;
               frmSaveDocDetails.ShowModal;
            finally
               frmSaveDocDetails.Free;
               AMessage := nil;
            end;
         end;
      end;

 //   StatusBar1.SimpleText := Format('%d messages dropped', [ListViewBrowser.Items.Count]);

{    if (ListViewBrowser.Items.Count > 1) then
    begin
      ListViewBrowser.Visible := True;
      SplitterBrowser.Left := Width;
      SplitterBrowser.Visible := True;
    end;
 }
    // If there's only one message, display it without further ado
 //   if (ListViewBrowser.Items.Count = 1) then
 //     ViewMessage(TMessage(ListViewBrowser.Items[0].Data))
//    else
      // Otherwise select and display the first message
//      if (ListViewBrowser.Items.Count > 0) then
//        ListViewBrowser.Items[0].Selected := True;
   end;
   DoSave();
end;

function TfrmMain.GetSender(const AMessage: IMessage): string;
var
  Prop: PSPropValue;
begin
  if (Succeeded(HrGetOneProp(AMessage, PR_SENDER_NAME, Prop))) then
    try
{$ifdef UNICODE}
      { TODO : TSPropValue.Value.lpszW is declared wrong }
      Result := PWideChar(Prop.Value.lpszW);
{$else}
      Result := Prop.Value.lpszA;
{$endif}
    finally
      MAPIFreeBuffer(Prop);
    end
  else
    Result := '';
end;

function TfrmMain.GetSubject(const AMessage: IMessage): string;
var
  Prop: PSPropValue;
begin
   try
      if (Succeeded(HrGetOneProp(AMessage, PR_SUBJECT, Prop))) then
         try
      {$ifdef UNICODE}
      { TODO : TSPropValue.Value.lpszW is declared wrong }
            Result := PWideChar(Prop.Value.lpszW);
      {$else}
            Result := Prop.Value.lpszA;
      {$endif}
         finally
            MAPIFreeBuffer(Prop);
         end
      else
         Result := '';
   except
   //
   end;
end;

function TfrmMain.GetFileID(const AMessage: IMessage): string;
var
  Prop: PSPropValue;
begin
   try
      if (Succeeded(HrGetOneProp(AMessage, NPR_FILEID, Prop))) then
         try
   {$ifdef UNICODE}
         { TODO : TSPropValue.Value.lpszW is declared wrong }
            Result := PWideChar(Prop.Value.lpszW);
   {$else}
            Result := Prop.Value.lpszA;
   {$endif}
         finally
            MAPIFreeBuffer(Prop);
         end
      else
         Result := '';
   except
   //
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

procedure TfrmMain.LoginSetup1Click(Sender: TObject);
begin
   frmLoginSetup := TfrmLoginSetup.Create(Application);
   frmLoginSetup.ShowModal;
   frmLoginSetup.Free;
   try
      GetUserID;
   except
      Application.Terminate;
   end;
end;

procedure TfrmMain.Close1Click(Sender: TObject);
begin
   try
      dmSaveDoc.orsAxiom.Disconnect;
      dmSaveDoc.Free;
   finally
      Application.Terminate;
   end;
end;

procedure TfrmMain.DoSave;
var
   i, y, x: integer;
   FileID, sSubject: string;
   AMessage: IMessage;
   AStream: TFileStream;
begin
   if not FormExists(frmSaveDocDetails) then
      frmSaveDocDetails := TfrmSaveDocDetails.Create(Self);

   if FFileList.Count > 0 then
   begin
      for i := 0 to (FFileList.Count - 1) do
      begin
         frmSaveDocDetails.DocName := FFileList.Strings[i];
         frmSaveDocDetails.Repaint;
         frmSaveDocDetails.ShowModal;
      end;
   end;

{   if (ListViewBrowser.Items.Count > 0) then
   begin
      for i := 0 to (ListViewBrowser.Items.Count -1)do
      begin
         ListViewBrowser.Items[i].Selected := True;
         try
            FileID := GetFileID(TMessage(ListViewBrowser.Items[i].Data).Msg);
         except
            FileID := '';
         end;

         if FileID = '' then
         begin
            sSubject := trimleft(ListViewBrowser.Items[i].SubItems.Strings[0]);
            for y := 1 to length(sSubject) do
            begin
               if sSubject[y] = '#' then
               begin
                  for x := y + 1 to length(sSubject) do
                  begin
                     if (sSubject[x] <> ' ') and (sSubject[x] <> ']') then
                        FileID := FileID + sSubject[x];
                  end;
               end;
            end;
         end;
//         AStream := TFileStream.Create(ListViewBrowser.Items[i].Data, fmOpenRead);
//         AStream.Write(ListViewBrowser.Items[i].Data, Length(ListViewBrowser.Items[i].Data));
         frmSaveDocDetails.EMessage := TMessage(ListViewBrowser.Items[i].Data).Msg;
         frmSaveDocDetails.DocName := trimleft(ListViewBrowser.Items[i].SubItems.Strings[0]);
         frmSaveDocDetails.FileID := FileID;
         frmSaveDocDetails.ShowModal;
      end;
   end; }
end;

procedure TfrmMain.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  Params.ExStyle := Params.ExStyle OR WS_EX_OVERLAPPEDWINDOW;
end;


procedure TfrmMain.FormDestroy(Sender: TObject);
var
  regAxiom: TRegistry;
  sRegistryRoot: string;
begin
   FFileList.Free;
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

procedure TfrmMain.ReadBodyText(const AMessage: IMessage);
var
  Outlook, MailItem: OLEVariant;
  outlookSel: Variant;
  i: Integer;
begin

end;

end.
