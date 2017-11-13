unit Unit1;

interface

uses
  DragDrop, DropSource, DropTarget, DragDropFile, ActiveX,
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, DropComboTarget, Registry, Menus, JvMenus;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    DropComboTarget1: TDropComboTarget;
    Popup: TJvPopupMenu;
    LoginSetup1: TMenuItem;
    N2: TMenuItem;
    Close1: TMenuItem;
    procedure FormCreate(Sender: TObject);
{    procedure DropEmptyTarget1Drop(Sender: TObject;
      ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
    procedure DropEmptyTarget1Enter(Sender: TObject;
      ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);   }
    procedure DropComboTarget1Drop(Sender: TObject;
      ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
    procedure FormDestroy(Sender: TObject);
    procedure LoginSetup1Click(Sender: TObject);
    procedure Close1Click(Sender: TObject);
  private
    { Private declarations }
     FCaption: string;
     tmpdir: string;
     FFileList: TStringList;
     procedure DoSave;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

uses
  DragDropFormats, ComObj, LoginDetails, SaveDoc, SaveDocFunc,
  SavedocDetails;

procedure TForm1.FormCreate(Sender: TObject);
var
   regAxiom: TRegistry;
   sRegistryRoot: string;
begin
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
   tmpdir := GetEnvironmentVariable('TMP')+'\';
end;

{procedure TForm1.DropEmptyTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
type
   PLargeint = ^Largeint;
var
   OldCount: integer;
   Buffer: AnsiString;
   p: PAnsiChar;
   i: integer;
   Stream: IStream;
   StatStg: TStatStg;
   Total, BufferSize, Chunk, Size: longInt;
   FirstChunk: boolean;
   FileStream: TFileStream;
   fName: string;
const
  MaxBufferSize = 100*1024; // 32Kb
begin
  // Transfer the file names and contents from the data format.
  if (TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames.Count > 0) then
  begin
    try
      // Note: Since we can actually drag and drop from and onto ourself,
      // Another, and more common, approach would be to reject or disable drops
      // onto ourself while we are performing drag/drop operations.
      for i := 0 to TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames.Count-1 do
      begin
        FCaption := TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileNames[i];

        // Get data stream from source.
        Stream := TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat).FileContentsClipboardFormat.GetStream(i);
        if (Stream <> nil) then
        begin
          // Read data from stream.
          Stream.Stat(StatStg, STATFLAG_NONAME);
          Total := StatStg.cbSize;

          // Assume that stream is at EOF, so set it to BOF.
          // See comment in TCustomSimpleClipboardFormat.DoSetData (in
          // DragDropFormats.pas) for an explanation of this.
          Stream.Seek(0, STREAM_SEEK_SET, PLargeint(nil)^);

          // If a really big hunk of data has been dropped on us we display a
          // small part of it since there isn't much point in trying to display
          // it all in the limited space we have available.
          // Additionally, it would be *really* bad for performce if we tried to
          // allocate a buffer that is too big and read sequentially into it. Tests have
          // shown that allocating a 10Mb buffer and trying to read data into it
          // in 1Kb chunks takes several minutes, while the same data can be
          // read into a 32Kb buffer in 1Kb chunks in seconds. The Windows
          // explorer uses a 1 Mb buffer, but that's too big for this demo.
          BufferSize := Total;
          if (BufferSize > MaxBufferSize) then
            BufferSize := MaxBufferSize;

          SetLength(Buffer, BufferSize);
          p := PAnsiChar(Buffer);
          Chunk := BufferSize;
          FirstChunk := True;
          while (Total > 0) do
          begin
            Stream.Read(p, Chunk, @Size);
            if (Size = 0) then
              break;

            inc(p, Size);
            dec(Total, Size);
            dec(Chunk, Size);

            if (Chunk = 0) or (Total = 0) then
            begin
              p := PAnsiChar(Buffer);
              // Lets write the buffer to disk
              fName := tmpdir + FCaption;
              FileStream := TFileStream.Create(fName, fmCreate);
              FileStream.WriteBuffer(p^, BufferSize-Chunk);
              Chunk := BufferSize;
              FirstChunk := False;
              FileStream.Free;
            end;
          end;

        end;
      end;
    finally

    end;
  end;
end;   }

{procedure TForm1.DropEmptyTarget1Enter(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
begin
  // Reject the drop unless the source supports *both* the FileContents and
  // FileGroupDescriptor formats in the storage medium we require (IStream).
  // Normally a drop is accepted if just one of our formats is supported.
   with TVirtualFileStreamDataFormat(DataFormatAdapterTarget.DataFormat) do
   begin
      if not(FileContentsClipboardFormat.HasValidFormats(DropEmptyTarget1.DataObject) and
         (AnsiFileGroupDescriptorClipboardFormat.HasValidFormats(DropEmptyTarget1.DataObject) or
         UnicodeFileGroupDescriptorClipboardFormat.HasValidFormats(DropEmptyTarget1.DataObject))) then
         Effect := DROPEFFECT_NONE;
   end;
end;    }

procedure TForm1.DropComboTarget1Drop(Sender: TObject;
  ShiftState: TShiftState; APoint: TPoint; var Effect: Integer);
var
  Stream: TStream;
  i: integer;
  Name: string;
begin
  // Extract and display dropped data.
  for i := 0 to DropComboTarget1.Data.Count-1 do
  begin
    Name := DropComboTarget1.Data.Names[i];
    if (Name = '') then
      Name := intToStr(i)+'.dat';
    Stream := TFileStream.Create(tmpdir {ExtractFilePath(Application.ExeName)} + Name, fmCreate);

    try
//        Caption := Name;
      // Copy dropped data to stream (in this case a file stream).
      Stream.CopyFrom(DropComboTarget1.Data[i], DropComboTarget1.Data[i].Size);
    finally
      Stream.Free;
    end;
  end;

  // Copy the rest of the dropped formats.
   if DropComboTarget1.Files.Count > 0 then
   begin
      try
         FFileList := TStringList.Create;
         FFileList.Assign(DropComboTarget1.Files);
         DoSave();
      finally
         FFileList.Free;
      end;
   end;
{
   ListBoxFiles.Items.Assign(DropComboTarget1.Files);
   ListBoxMaps.Items.Assign(DropComboTarget1.FileMaps);
   EditURLURL.Text := DropComboTarget1.URL;
   EditURLTitle.Text := DropComboTarget1.Title;
   ImageBitmap.Picture.Assign(DropComboTarget1.Bitmap);
   ImageMetaFile.Picture.Assign(DropComboTarget1.MetaFile);
   MemoText.Lines.Text := DropComboTarget1.Text;
}

end;


procedure TForm1.FormDestroy(Sender: TObject);
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

procedure TForm1.LoginSetup1Click(Sender: TObject);
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

procedure TForm1.Close1Click(Sender: TObject);
begin
   try
      dmSaveDoc.orsAxiom.Disconnect;
      dmSaveDoc.Free;
   finally
      Application.Terminate;
   end;
end;

procedure TForm1.DoSave;
var
   i: integer;
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
end;

end.
