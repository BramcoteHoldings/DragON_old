unit SavedocDetails;

interface

uses
  MapiDefs, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer,
  cxEdit, Vcl.Menus, cxClasses, cxShellBrowserDialog, uRwEasyMAPI,
  VCL.uRwMAPISession, cxCheckBox, cxGroupBox, cxRadioGroup, cxMemo,
  cxDropDownEdit, cxLookupEdit, cxDBLookupEdit, cxDBLookupComboBox, cxTextEdit,
  dxStatusBar, Vcl.StdCtrls, cxButtons, cxLabel, cxMaskEdit, cxButtonEdit,
  JvExControls, JvLabel, uRwMapiInterfaces,
  Dialogs, DB;

 
type
  TfrmSaveDocDetails = class(TForm)
    btnEditMatter: TcxButtonEdit;
    lblMatter: TcxLabel;
    btnSave: TcxButton;
    btnClose: TcxButton;
    StatusBar: TdxStatusBar;
    cxLabel1: TcxLabel;
    txtDocName: TcxTextEdit;
    cmbCategory: TcxLookupComboBox;
    cxLabel2: TcxLabel;
    cxLabel4: TcxLabel;
    cmbClassification: TcxLookupComboBox;
    edKeywords: TcxTextEdit;
    memoPrecDetails: TcxMemo;
    cxLabel5: TcxLabel;
    cxLabel6: TcxLabel;
    cmbAuthor: TcxLookupComboBox;
    cxLabel7: TcxLabel;
    rgStorage: TcxRadioGroup;
    cxLabel3: TcxLabel;
    btnTxtDocPath: TcxButtonEdit;
    lblPath: TJvLabel;
    cbEmailAttachSave: TcxCheckBox;
    MAPISession: TRwMAPISession;
    BrowseDlg: TcxShellBrowserDialog;
    procedure btnEditMatterPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure btnCloseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure rgStorageClick(Sender: TObject);
    procedure btnEditMatterPropertiesValidate(Sender: TObject;
      var DisplayValue: Variant; var ErrorText: TCaption;
      var Error: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnTxtDocPathPropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure cmbCategoryPropertiesInitPopup(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure MapiSessionAfterLogon(Sender: TObject);
    procedure MapiSessionBeforeLogoff(Sender: TObject);
  private
    { Private declarations }
    nMatter: integer;
    tmpFileName: string;
    tmpdir: string;
    FFileID: string;
    FPrec_Category: string;
    FEditing: boolean;
    FSavedInDB: string;
    FDocName: string;
    FPrec_Classification: string;
    FDoc_Keywords: string;
    FDoc_Precedent: string;
    FDoc_FileName: string;
    FDoc_Author: string;
    FOldDocName: string;
    FNewDocName: string;
    FMatter: integer;
    AFileID: string;
    FMessage: IMessage;
    FPath: string;
    FURLOnly: boolean;
    FMsgStore: IRwMapiMsgStore;
    FSavedMsg: IRwMapiMessage;
    function SaveDocument(DocSequence: string; ANewDocName: string): boolean;
    procedure CreateParams(var Params: TCreateParams); override;
    function WriteFileToDisk(var ANewDocName: string; AOldDocName: string; ADeleteFile: boolean = False): boolean;
    function WriteFileDetailsToDB(ANewDocName: string; AParentDocID: integer): boolean;
  public
    { Public declarations }
    property DocName: string read FDocName write FDocName;
    property Matter: integer read FMatter write FMatter;
    property OldDocName: string read FOldDocName write FOldDocName;
    property NewDocName: string read FNewDocName write FNewDocName;
    property FileID: string read AFileID write AFileID;
    property EMessage: IMessage read  FMessage write FMessage;
    property URLOnly: boolean read FURLOnly write FURLOnly;
  end;

var
  frmSaveDocDetails: TfrmSaveDocDetails;

function ShowDocSave: Integer; StdCall;

implementation

uses
   MatterSearch
   , savedoc
   , SaveDocFunc
   , uRwSysUtils
   , uRwDateTimeUtils
   , uRwMapiUtils
   , uRwBoxes
   , uRwMapiMessage
   , uRwMapiProps
   , uRwMapiBase;

{$R *.dfm}

function ShowDocSave:integer;
var
   frmSaveDocDetails: TfrmSaveDocDetails;
begin
//   Application.Handle := AHandle;
   frmSaveDocDetails := TfrmSaveDocDetails.Create(Application);
   try
      frmSaveDocDetails.ShowModal;
      Result := frmSaveDocDetails.nMatter;
   finally
      frmSaveDocDetails.Free;
   end;
end;

procedure TfrmSaveDocDetails.btnEditMatterPropertiesButtonClick(Sender: TObject;
  AButtonIndex: Integer);
begin
   try
      frmMtrSearch :=TfrmMtrSearch.Create(Self);
      frmMtrSearch.MakeSql;
      if (frmMtrSearch.ShowModal = mrOK) then
      begin
         btnEditMatter.Text := frmMtrSearch.vMattersFILEID.EditValue;   //  dmSaveDoc.qryMatters.FieldByName('fileid').AsString;
         nMatter := frmMtrSearch.vMattersNMATTER.EditValue;
         FFileID := frmMtrSearch.vMattersFILEID.EditValue;
      end;
   finally
      frmMtrSearch.Release;
      Self.Invalidate;
   end;
end;

procedure TfrmSaveDocDetails.btnCloseClick(Sender: TObject);
begin
//   dmSaveDoc.orsAxiom.Disconnect;
//   dmSaveDoc.Free;
   Close;
end;

procedure TfrmSaveDocDetails.FormCreate(Sender: TObject);
begin
//   Application.CreateForm(TdmSaveDoc, dmSaveDoc);
   try
      GetUserID;
      if (FSavedInDB = 'N') or (FSavedInDB = '')  then
      begin
         rgStorage.ItemIndex := SystemInteger('DFLT_DOC_SAVE_OPTION');
         rgStorage.Enabled := (SystemString('DISABLE_SAVE_MODE') = 'N');
         lblPath.Caption :=  IncludeTrailingPathDelimiter(SystemString('DRAG_DEFAULT_DIRECTORY'));
//         btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY');
         FPath := lblPath.Caption;  //btnTxtDocPath.Text;
      end;

      dmSaveDoc.qryPrecCategory.Open;
      dmSaveDoc.qryPrecClassification.Open;
      dmSaveDoc.qryEmployee.Open;
//      dmSaveDoc.qryMatters.Active := True;
//      dmSavedoc.qryPrecCategory.Open;
      StatusBar.Panels[0].Text := 'Ver: '+ReportVersion + ' (' +DateTimeToStr(FileDateToDateTime(FileAge(Application.ExeName)))+')';
   except
      Application.Terminate;
   end;
end;

procedure TfrmSaveDocDetails.btnSaveClick(Sender: TObject);
var
   DocSequence, AParsedDocName, ATempParsedDocName: string;
   bUsePath: boolean;
   bError: boolean;
   TempName, TempExt, TempDir, TempFile: string;
begin
   bError := False;
   if btnEditMatter.Text = '' then
   begin
//      with Application do
//      begin
//         NormalizeTopMosts;
         MessageBox(Self.Handle,'Please enter a Matter number.','DragON',MB_OK+MB_ICONEXCLAMATION);
         bError := True;
//         RestoreTopMosts;
//         exit;
//      end;
   end
   else
   begin
      dmSaveDoc.qryGetMatter.Close;
      dmSaveDoc.qryGetMatter.ParamByName('FILEID').AsString := string(UpperCase(btnEditMatter.Text));
      dmSaveDoc.qryGetMatter.Open;
      if dmSavedoc.qryGetMatter.Eof then
      begin
         MessageBox(Self.Handle,'Invalid Matter Number','DragON', MB_OK+MB_ICONERROR);
         bError := True;
      end;
   end;
   
   if btnTxtDocPath.Text <> '' then
   begin
//      try
         if cmbAuthor.Text = '' then
         begin
//            with Application do
//            begin
//               NormalizeTopMosts;
               MessageBox(Self.Handle, 'Please enter an Author.','DragON',MB_OK+MB_ICONEXCLAMATION);
               bError := True;
//               RestoreTopMosts;
//               exit;
//            end;
         end;
   end;

   if not bError then
   begin
      try
         if dmSaveDoc.uniInsight.InTransaction then
            dmSaveDoc.uniInsight.Commit;
         dmSaveDoc.uniInsight.StartTransaction;
//         dmSaveDoc.qryMatterAttachments.ParamByName('docid').AsString := dmSaveDoc.DocID;
         dmSaveDoc.qryMatterAttachments.Open;

//         dmSaveDoc.qryMatterAttachments.insert;

         FEditing := False;
         bUsePath := False;
         tmpdir := IncludeTrailingPathDelimiter(GetEnvironmentVariable('TEMP'));

         if btnTxtDocPath.Visible and
            (btnTxtDocPath.Text <> '') then
               bUsePath := True;
//          + '_' + DocSequence

         if btnTxtDocPath.Text <> '' then
            ATempParsedDocName := ParseMacros(lblPath.Caption+btnTxtDocPath.Text,TableInteger('MATTER','FILEID',uppercase(string(btnEditMatter.Text)),'NMATTER'));

         if (UpperCase(DocName) <> UpperCase(ATempParsedDocName)) then
         begin
            if btnTxtDocPath.Enabled then
            begin
               if btnTxtDocPath.Text = '' then
               begin
                  TempName := Copy(ExtractFileName(btnTxtDocPath.Text),1, Length(ExtractFileName(btnTxtDocPath.Text)) - Length(ExtractFileExt(btnTxtDocPath.Text)));
                  TempExt := ExtractFileExt(btnTxtDocPath.Text);
                  NewDocName := TempName+ '_[DOCSEQUENCE]' + TempExt;   //txtDocName.Text
               end
               else
               begin
                  TempName := Copy(ExtractFileName(btnTxtDocPath.Text),1, Length(ExtractFileName(btnTxtDocPath.Text)) - Length(ExtractFileExt(btnTxtDocPath.Text)));
                  TempExt := ExtractFileExt(btnTxtDocPath.Text);
                  NewDocName := lblPath.Caption + TempName+ '_[DOCSEQUENCE]' + TempExt; // btnTxtDocPath.Text;
               end;
            end
            else
            begin
               TempName := Copy(ExtractFileName(DocName),1, Length(ExtractFileName(DocName)) - Length(ExtractFileExt(DocName)));
               TempExt := ExtractFileExt(DocName);
               TempDir := ExtractFileDir(DocName);
               TempFile := TempDir + TempName+ '_[DOCSEQUENCE]' + TempExt;
               NewDocName := TempFile;
            end;
            AParsedDocName := ParseMacros(NewDocName,TableInteger('MATTER','FILEID',uppercase(string(btnEditMatter.Text)),'NMATTER'));
         end
         else
            AParsedDocName := DocName;

//         AParsedDocName := ParseMacros(NewDocName,TableInteger('MATTER','FILEID',uppercase(string(btnEditMatter.Text)),'NMATTER'));

         try
            if SaveDocument(DocSequence, AParsedDocName)then
            begin
               if (rgStorage.ItemIndex = 0) then
                  DeleteFile(tmpFileName);
               if dmSaveDoc.uniInsight.InTransaction then
                  dmSaveDoc.uniInsight.Commit;
               Self.Close;
            end;
         except
            raise;
         end;
      except
         if dmSaveDoc.uniInsight.InTransaction then
            dmSaveDoc.uniInsight.Rollback;
      end;
   end;
//   else
//      ModalResult := mrNone;
end;

function TfrmSaveDocDetails.SaveDocument(DocSequence, ANewDocName: string): boolean;
var
   SavedInDB, DispName,
   ltmpdir,
   FileExt,
   POldName, AExt, ADispName,
   AParsedDocName, VarDocName: string;
   nCat, nClass,
   FileImg, FileSave,
   AError, iParentDocID: integer;
   ADocumentSaved: boolean;
   Outlook, MailItem: OLEVariant;
   AttTable: IRwMapiAttachmentTable;
   Attachment: IRwMapiAttachment;
begin
   SaveDocument := False;
   ADocumentSaved := True;
   VarDocName := '';
   try
      if not URLOnly then
      begin
         case rgStorage.ItemIndex of
            0: begin
                  ltmpdir := ParseMacros(lblPath.Caption + btnTxtDocPath.Text,TableInteger('MATTER','FILEID',FFileID,'NMATTER'));
                  ltmpDir := tmpdir + ExtractFileName(ltmpdir);

                  if ExtractFileExt(ltmpdir) = '' then
                     ltmpdir := ltmpdir + '.doc';

                  NewDocName := ltmpdir;
               end;
            1: begin
                  SavedInDB := 'N';
                  if (UpperCase(ANewDocName) = UpperCase(OldDocName)) then
                     ADocumentSaved := True
                  else
                     ADocumentSaved := WriteFileToDisk(ANewDocName, OldDocName);
               end;
         end;
      end;

      if ADocumentSaved then
      begin
         try
            dmSaveDoc.qryMatterAttachments.insert;

            iParentDocID := StrToInt(dmSaveDoc.DocID);

            dmSaveDoc.qryMatterAttachments.FieldByName('docid').AsString := dmSaveDoc.DocID;
            dmSaveDoc.qryMatterAttachments.FieldByName('parentdocid').AsInteger := iParentDocID;
            dmSaveDoc.qryMatterAttachments.FieldByName('fileid').AsString := btnEditMatter.Text;
            dmSaveDoc.qryMatterAttachments.FieldByName('nmatter').AsInteger := nMatter;
            dmSaveDoc.qryMatterAttachments.FieldByName('auth1').AsString := cmbAuthor.Text;   //  UpperCase(dmSaveDoc.UserID);
//            if not FEditing then

//            if rgStorage.ItemIndex = 0 then
//               dmSaveDoc.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(NewDocName)
//            else

            if URLOnly then
            begin
               dmSaveDoc.qryMatterAttachments.FieldByName('DOC_NAME').AsString := 'Internet shortcut';
               dmSaveDoc.qryMatterAttachments.FieldByName('URL').AsString := DocName;
            end
            else
               dmSaveDoc.qryMatterAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(ANewDocName);

            dmSaveDoc.qryMatterAttachments.FieldByName('DESCR').AsString := txtDocName.Text;
            if (not URLOnly) then
            begin
               dmSaveDoc.qryMatterAttachments.FieldByName('FILE_EXTENSION').AsString := Copy(ExtractFileExt(ANewDocName),2, length(ExtractFileExt(ANewDocName)));
               if rgStorage.ItemIndex = 0 then
               begin
                  TBlobField(dmSaveDoc.qryMatterAttachments.fieldByname('DOCUMENT')).LoadFromFile(ANewDocName);
               end
               else
               begin
                  dmSaveDoc.qryMatterAttachments.FieldByName('PATH').AsString := IndexPath(ANewDocName, 'DOC_SHARE_PATH');  //NewDocName;
                  dmSaveDoc.qryMatterAttachments.FieldByName('DISPLAY_PATH').AsString := ANewDocName;
               end;
            end;

            FileExt := uppercase(dmSaveDoc.qryMatterAttachments.FieldByName('FILE_EXTENSION').AsString);
            if (FileExt = 'DOC') or (FileExt = 'DOCX') then
               FileImg := 2
            else if (FileExt = 'XLS') or (FileExt = 'XLSX') then
               FileImg := 3
            else if (FileExt = 'PDF')  then
               FileImg := 5
            else if (FileExt = 'MSG') then
               FileImg := 4
            else if URLOnly then
               FileImg := 6
            else if (AExt = 'PPT') or (AExt = 'PPTX') then
               FileImg := 8
            else if (AExt = 'ZIP') or (AExt = 'ZIPX') then
               FileImg := 9
            else
               FileImg := 1;

            try
               if URLOnly then
                  dmSaveDoc.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := Now
               else
          //        dmSaveDoc.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := FileDateToDateTime(FileAge(OldDocName));  //  Now;
               begin
               if (UpperCase(FileExt) = 'MSG') then
               begin
                  if (MapiSession.Active = False) then
                  begin
                     try
                        MapiSession.LogonInfo.UseExtendedMapi    := True;
//                        MapiSession.LogonInfo.ProfileName        := TableString('EMPLOYEE', 'CODE',cmbAuthor.Text ,'EMAIL_PROFILE_DEFAULT'); // 'Outlook';
//                        MapiSession.LogonInfo.Password           := '';
//                        MapiSession.LogonInfo.ProfileRequired    := True;
                        MapiSession.LogonInfo.NewSession         := False;
                        MapiSession.LogonInfo.ShowPasswordDialog := False;
//                        MapiSession.LogonInfo.ShowLogonDialog    := True;
                        MapiSession.Logon;
                     finally
                           //
                     end;
                  end;

                  FSavedMsg := FMsgStore.OpenSavedMessage(ANewDocName);
                  dmSaveDoc.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := RwUTCToLocal(FSavedMsg.PropByName(PR_MESSAGE_DELIVERY_TIME).AsDateTime);
                  POldName := OldDocName;
                  dmSaveDoc.qryMatterAttachments.FieldByName('EMAIL_SENT_TO').AsString := FSavedMsg.PropByName(PR_DISPLAY_TO).AsString;
                  if (cbEmailAttachSave.Checked) then
                  begin
                     AttTable := FSavedMsg.Attachments;
//                     AttTable.SetFields(VarArrayOf([PR_ATTACH_NUM]));
                     AttTable.First;
                     while not AttTable.EOF do
                     begin
                        Attachment := FSavedMsg.OpenAttachment(AttTable.FieldByName(PR_ATTACH_NUM).AsInteger, alReadOnly);

                        DispName := Attachment.PropByName(PR_DISPLAY_NAME).AsString;

                        if DispName = '' then
                           DispName := ExtractFileName(Attachment.PropByName(PR_ATTACH_FILENAME).AsString);

                        if DispName = '' then
                           DispName := ExtractFileName(Attachment.PropByName(PR_ATTACH_LONG_FILENAME).AsString);

                        while Pos('/', DispName) > 0 do
                           DispName[Pos('/', DispName)] := '.';

                        while Pos('\', DispName) > 0 do
                           DispName[Pos('\', DispName)] := '.';

                        AExt := ExtractFileExt(DispName);
                        ADispName := Copy (DispName,1, Length(DispName)- Length(AExt));
                        ADispName := ADispName + '_' + '[DOCSEQUENCE]';
                        DispName := ADispName + AExt;

                        OldDocName := tmpdir + DispName;
                        Attachment.SaveToFile(OldDocName);
                        VarDocName := ParseMacros(lblPath.Caption + DispName,TableInteger('MATTER','FILEID',FFileID,'NMATTER'));
                        if (VarDocName <> OldDocName) then
                           ADocumentSaved := WriteFileToDisk(VarDocName, OldDocName, True);
                        WriteFileDetailsToDB(VarDocName, iParentDocID);
                        AttTable.Next;
                     end;
                  end;
                  OldDocName := POldName;
                  if MapiSession.Active then
                     MapiSession.Logoff;
               end
               else
                  dmSaveDoc.qryMatterAttachments.FieldByName('D_CREATE').AsDateTime := FileDateToDateTime(FileAge(ANewDocName));
            end;
            except
                on E: Exception do
                  begin
//                     dmSaveDoc.uniInsight.Rollback;
                     MessageBox(Self.Handle, pchar('Error during saving document to the database: ' + E.Message), 'DragON', MB_OK + MB_ICONERROR);
                     SaveDocument := False;
                  end;
            end;

            dmSaveDoc.qryMatterAttachments.FieldByName('IMAGEINDEX').AsInteger := FileImg;
            if cmbCategory.Text <> '' then
               dmSaveDoc.qryMatterAttachments.FieldByName('NPRECCATEGORY').AsString := cmbCategory.EditValue;
            dmSaveDoc.qryMatterAttachments.FieldByName('precedent_details').AsString := memoPrecDetails.Text;
            dmSaveDoc.qryMatterAttachments.FieldByName('KEYWORDS').AsString := edKeywords.Text;
            if cmbClassification.Text <> '' then
               dmSaveDoc.qryMatterAttachments.FieldByName('NPRECCLASSIFICATION').AsString := cmbClassification.EditValue;

            dmSaveDoc.qryMatterAttachments.Post;
            dmSaveDoc.qryMatterAttachments.ApplyUpdates;
//            dmSaveDoc.uniInsight.Commit;
         except
            on E: Exception do
            begin
 //              dmSaveDoc.uniInsight.Rollback;
               MessageBox(Self.Handle, pchar('Error during saving document to the database.  The document may exist in the file system : ' + E.Message), 'DragON', MB_OK + MB_ICONERROR);
               SaveDocument := False;
            end;
         end;
         SaveDocument := True;
      end;
   except
      on E: Exception do
       begin
          MessageBox(Self.Handle, PChar('Error during saving document (trying to establish active document): ' + E.Message), PChar('DragON'), MB_ICONERROR);
//          MessageDlg('Error during saving document: ' + E.Message, mtError, [mbOK], 0);
          SaveDocument := False;
       end;
   end;
end;

procedure TfrmSaveDocDetails.rgStorageClick(Sender: TObject);
begin
   case rgStorage.ItemIndex of
      0: begin
            btnTxtDocPath.Visible := False;
         end;
      1: begin
            btnTxtDocPath.Visible := True;
         end;
   end;
end;

procedure TfrmSaveDocDetails.btnEditMatterPropertiesValidate(
  Sender: TObject; var DisplayValue: Variant; var ErrorText: TCaption;
  var Error: Boolean);
begin
   if DisplayValue <> '' then
   begin
      dmSaveDoc.qryGetMatter.Close;
      dmSaveDoc.qryGetMatter.ParamByName('FILEID').AsString := string(UpperCase(DisplayValue));
      dmSaveDoc.qryGetMatter.Open;
      if dmSavedoc.qryGetMatter.Eof then
         MessageBox(Self.Handle,'Invalid Matter Number','DragON', MB_OK+MB_ICONERROR)
      else
      begin
         nMatter := dmSaveDoc.qryGetMatter.FieldByName('NMATTER').AsInteger;
         FFileID := string(UpperCase(DisplayValue));
      end;
   end;
end;

procedure TfrmSaveDocDetails.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//   dmSaveDoc.uniInsight.Commit;
//   if dmSaveDoc.orsAxiom.Connected then
//      dmSaveDoc.orsAxiom.Disconnect;
//   Action := caFree;
end;

procedure TfrmSaveDocDetails.btnTxtDocPathPropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
begin
   case AButtonIndex of
      0: begin
            if BrowseDlg.Execute = True then
               lblPath.Caption := BrowseDlg.Path; // btnTxtDocPath.Text := BrowseDlg.SelectedFolder;
         end;
      1: lblPath.Caption := IncludeTrailingPathDelimiter(SystemString('DRAG_DEFAULT_DIRECTORY'));   // btnTxtDocPath.Text := SystemString('DRAG_DEFAULT_DIRECTORY');
   end;
end;

procedure TfrmSaveDocDetails.cmbCategoryPropertiesInitPopup(
  Sender: TObject);
begin
//   dmSavedoc.qryPrecCategory.Close;
//   dmSavedoc.qryPrecCategory.Open;
end;

procedure TfrmSaveDocDetails.FormShow(Sender: TObject);
begin
   if DocName <> '' then
   begin
      cbEmailAttachSave.Enabled := False;
      lblPath.Caption := IncludeTrailingPathDelimiter(FPath);
      btnTxtDocPath.Text :=  ExtractFileName(DocName);
      FOldDocName := DocName;
      if URLOnly then
      begin
         btnTxtDocPath.Enabled := False;
         txtDocName.Text := DocName;
      end
      else
      begin
         txtDocName.Text := ExtractFileName(DocName);
         if (UpperCase(ExtractFileExt(DocName)) = '.MSG' ) then
         begin
           cbEmailAttachSave.Enabled := True;
           cbEmailAttachSave.Checked := (SystemString('email_separate_attachments') = 'Y');
         end;
      end;
   end;
   cmbAuthor.EditValue := TableString('EMPLOYEE','USER_NAME',UpperCase(dmSaveDoc.uniInsight.Username), 'CODE');
   frmSaveDocDetails.BringToFront;
//   if (not frmSaveDocDetails.Active) then
   frmSaveDocDetails.Activate;
   SetForegroundWindow(frmSaveDocDetails.Handle);
   btnEditMatter.SetFocus;
end;

procedure TfrmSaveDocDetails.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  with Params do begin
    ExStyle := ExStyle or WS_EX_TOPMOST;
    WndParent := GetDesktopwindow;
  end;
end;
 
procedure TfrmSaveDocDetails.FormDestroy(Sender: TObject);
begin
//   if dmSavedoc.qryPrecCategory.Active then
//      dmSavedoc.qryPrecCategory.close;
end;

procedure TfrmSaveDocDetails.MapiSessionAfterLogon(Sender: TObject);
begin
   FMsgStore := MapiSession.OpenDefaultMsgStore(alBestAccess);
end;

procedure TfrmSaveDocDetails.MapiSessionBeforeLogoff(Sender: TObject);
begin
   FreeAndNil(FSavedMsg);
   FreeAndNil(FMsgStore);
end;

function TfrmSaveDocDetails.WriteFileToDisk(var ANewDocName: string; AOldDocName: string; ADeleteFile: boolean): boolean;
var
   NewDocName{, AParsedDocName}: string;
   ADocumentSaved: boolean;
   AError: integer;
begin
   ADocumentSaved := True;
//   AParsedDocName := ParseMacros(ANewDocName,TableInteger('MATTER','FILEID',string(btnEditMatter.Text),'NMATTER'));

//   NewDocName := AParsedDocName;
//   ANewDocName := NewDocName;

   if (ANewDocName <> '') then
   begin
      if not DirectoryExists(ExtractFileDir(ANewDocName)) then
         ForceDirectories(ExtractFileDir(ANewDocName));
   end;

   if Assigned(TMessage(EMessage)) then
   begin
      MessageDlg('DragON Error',mtInformation,[mbOk], 0 );
   end
   else
   try
      if not CopyFile(PChar(AOldDocName) ,pchar(ANewDocName), true) then
      begin
         AError := GetLastError;
         case AError of
            80: begin
                   if MessageBox(Self.Handle, pchar('File [' + ANewDocName + '] already exists. Do you want to overwrite it?' ), 'DragON', MB_YESNO + MB_ICONQUESTION) = IDYES then
                      ADocumentSaved := CopyFile(PChar(OldDocName) ,pchar(ANewDocName), false)
                   else
                      ADocumentSaved := False;
                end;
            82: begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The directory or file could not be created.'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
            5:  begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  Access denied.'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
            39,112: begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The disk is full!'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
            111:begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The filename is to long!'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
            53 :begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The network path was not found!'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
            3:  begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The system cannot find the path specified!'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
            2:  begin
                  MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The system cannot find the file specified!'), 'DragON', MB_OK + MB_ICONERROR);
                  ADocumentSaved := False;
                end;
         else
            MessageBox(Self.Handle, pchar('There was an error during the saving of the document.  The document was not saved. Error: ' + IntTostr(AError)), 'DragON', MB_OK + MB_ICONERROR);
            ADocumentSaved := False;
         end;
      end;
      if (ADeleteFile and ADocumentSaved) then
         DeleteFile(AOldDocName);
   except
      on E: Exception do
      begin
         MessageBox(Self.Handle, pchar('There was an Error during the saving of the document.  The document was not saved: ' + E.Message), 'DragON', MB_OK + MB_ICONERROR);
         ADocumentSaved := False;
      end;
   end;
   Result := ADocumentSaved;
end;

function TfrmSaveDocDetails.WriteFileDetailsToDB(ANewDocName: string; AParentDocID: integer): boolean;
var
   FileExt: string;
   FileImg: integer;
begin
   if dmSaveDoc.uniInsight.InTransaction then
      dmSaveDoc.uniInsight.Commit;
   dmSaveDoc.uniInsight.StartTransaction;
   try
      if dmSaveDoc.qrySaveEmailAttachments.State = dsInactive then
      begin
         dmSaveDoc.qrySaveEmailAttachments.ParamByName('docid').AsString := dmSaveDoc.AttDocID;
         dmSaveDoc.qrySaveEmailAttachments.Open;
      end;

      if dmSaveDoc.qrySaveEmailAttachments.State = dsBrowse then
         dmSaveDoc.qrySaveEmailAttachments.Insert;

      dmSaveDoc.qrySaveEmailAttachments.FieldByName('docid').AsString := dmSaveDoc.AttDocID;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('fileid').AsString := btnEditMatter.Text;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('nmatter').AsInteger := nMatter;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('auth1').AsString := cmbAuthor.Text; //  UpperCase(dmSaveDoc.UserID);

      dmSaveDoc.qrySaveEmailAttachments.FieldByName('DOC_NAME').AsString := ExtractFileName(ANewDocName);

      dmSaveDoc.qrySaveEmailAttachments.FieldByName('DESCR').AsString := txtDocName.Text;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('FILE_EXTENSION').AsString := Copy(ExtractFileExt(ANewDocName),2, length(ExtractFileExt(ANewDocName)));

      dmSaveDoc.qrySaveEmailAttachments.FieldByName('PATH').AsString := IndexPath(NewDocName, 'DOC_SHARE_PATH');  //NewDocName;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('DISPLAY_PATH').AsString := ANewDocName;

      dmSaveDoc.qrySaveEmailAttachments.FieldByName('PARENTDOCID').AsInteger := AParentDocID;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('IS_ATTACHMENT').AsString := 'Y';

      FileExt := uppercase(dmSaveDoc.qrySaveEmailAttachments.FieldByName('FILE_EXTENSION').AsString);
      if (FileExt = 'DOC') or (FileExt = 'DOCX') then
         FileImg := 2
      else if (FileExt = 'XLS') or (FileExt = 'XLSX') then
         FileImg := 3
      else if (FileExt = 'PDF')  then
         FileImg := 5
      else if (FileExt = 'MSG') then
         FileImg := 4
      else if URLOnly then
         FileImg := 6
      else
         FileImg := 1;

      try
         dmSaveDoc.qrySaveEmailAttachments.FieldByName('D_CREATE').AsDateTime := FileDateToDateTime(FileAge(ANewDocName));
      except
       //
      end;

      dmSaveDoc.qrySaveEmailAttachments.FieldByName('IMAGEINDEX').AsInteger := FileImg;
      if (cmbCategory.Text <> '') then
         dmSaveDoc.qrySaveEmailAttachments.FieldByName('NPRECCATEGORY').AsString := cmbCategory.EditValue;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('precedent_details').AsString := memoPrecDetails.Text;
      dmSaveDoc.qrySaveEmailAttachments.FieldByName('KEYWORDS').AsString := edKeywords.Text;
      if (cmbClassification.Text <> '') then
         dmSaveDoc.qrySaveEmailAttachments.FieldByName('NPRECCLASSIFICATION').AsString := cmbClassification.EditValue;

      dmSaveDoc.qrySaveEmailAttachments.ApplyUpdates;
      dmSaveDoc.uniInsight.Commit;
   except
      dmSaveDoc.qrySaveEmailAttachments.RestoreUpdates; //<restore update result for applied records>
      dmSaveDoc.uniInsight.Rollback; //<on failure, undo the changes>
      raise; //<raise the exception to prevent a call to CommitUpdates!>
   end;
   dmSaveDoc.qrySaveEmailAttachments.CommitUpdates;
end;


end.
