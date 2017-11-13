unit SaveDocFunc;

interface

uses
   Forms, Registry, INIFiles, SysUtils, Variants, Windows, Ora, OracleUniProvider, Uni, DBAccess, MemDS;

const
  C_MACRO_USERHOME      = '[USERHOME]';
  C_MACRO_NMATTER       = '[NMATTER]';
  C_MACRO_FILEID        = '[FILEID]';
  C_MACRO_TEMPDIR       = '[TEMPDIR]';
  C_MACRO_TEMPFILE      = '[TEMPFILE]';
  C_MACRO_DATE          = '[DATE]';
  C_MACRO_TIME          = '[TIME]';
  C_MACRO_DATETIME      = '[DATETIME]';
  C_MACRO_CLIENTID      = '[CLIENTID]';
  C_MACRO_AUTHOR        = '[AUTHOR]';
  C_MACRO_USERINITIALS  = '[USERINITIALS]';
  C_MACRO_DOCSEQUENCE   = '[DOCSEQUENCE]';
  C_MACRO_DOCID         = '[DOCID]';



   procedure GetUserID;
   function GetSeqNumber(sSequence: string): Integer;
   function TableString(Table, LookupField, LookupValue, ReturnField: string): string; overload;
   function TableInteger(Table, LookupField, LookupValue, ReturnField: string): integer; overload;
   function ParseMacros(AFileName: String; ANMatter: Integer): String;
   function SystemString(sField: string): string;
   function SystemInteger(sField: string): integer;
   function ProcString(Proc: string; LookupValue: integer): string;
   function ReportVersion: string;
   function FormExists(frmInput: TForm): boolean;
   function TokenizePath(var s,w:string): boolean;
   function IndexPath(PathText, PathLoc: string): string;


implementation

uses
 SaveDoc;

var // for macros..
  GTempPath,
  GHomePath: String;

procedure GetUserID;
var
  regAxiom: TRegistry;
  sRegistryRoot: string;
  LoginUser: string;
begin
   // Find out what registry key we are using
{   try
      iniAxiom := TINIFile.Create(ExtractFilePath(Application.EXEName) + '..\Axiom.INI');
   except
      // do nothing
   end;
   sRegistryRoot := iniAxiom.ReadString('Main', 'RegistryRoot', 'Software\Colateral\Axiom');
   iniAxiom.Free;             }

   sRegistryRoot := 'Software\Colateral\Axiom';

   regAxiom := TRegistry.Create;
   try
      regAxiom.RootKey := HKEY_CURRENT_USER;
      if regAxiom.OpenKey(sRegistryRoot+'\InsightDocSave', False) then
      begin
         if regAxiom.ReadString('Password') <> '' then
         begin
            try
               if dmSaveDoc.uniInsight.Connected then
                  dmSaveDoc.uniInsight.Disconnect;
               if regAxiom.ReadString('Net') = 'Y' then
                  dmSaveDoc.uniInsight.SpecificOptions.Values['Direct'] := 'true'
               else
                  dmSaveDoc.uniInsight.SpecificOptions.Values['Direct'] := 'false';
               dmSaveDoc.uniInsight.Server := regAxiom.ReadString('Server Name');
               dmSaveDoc.uniInsight.Username := regAxiom.ReadString('User Name');
               dmSaveDoc.uniInsight.Password := regAxiom.ReadString('Password');
               dmSaveDoc.uniInsight.Connect;
            except
               Application.MessageBox('Could not connect to Insight database','DragON');
               Application.Terminate;
            end;
         end
         else
         begin
            Application.MessageBox('Could not connect to Insight database','DragON');
            Application.Terminate;
         end;
         regAxiom.CloseKey;
         regAxiom.RootKey := HKEY_CURRENT_USER;
         if regAxiom.OpenKey(sRegistryRoot+'\InsightDocSave', False) then
         begin
            LoginUser := regAxiom.ReadString('User Name');
            with dmSaveDoc.qryEmps do
            begin
               Close;
               SQL.Text := 'SELECT CODE FROM EMPLOYEE WHERE UPPER(USER_NAME) = upper(''' + LoginUser + ''') AND ACTIVE = ''Y''';
               Prepare;
               Open;
               // Make sure that the UserID is valid
               if IsEmpty then
               begin
                  Application.MessageBox('User not valid.','DragON');
                  Application.Terminate;
               end
               else
               dmSaveDoc.UserID := FieldByName('CODE').AsString;
               Close;
            end;
            dmSaveDoc.qryGetEntity.Close();
            dmSaveDoc.qryGetEntity.ParamByName('Emp').AsString := uppercase(dmSaveDoc.UserID);
            dmSaveDoc.qryGetEntity.ParamByName('Owner').AsString := 'Desktop';
            dmSaveDoc.qryGetEntity.ParamByName('Item').AsString := 'Entity';
            dmSaveDoc.qryGetEntity.Open();
            dmSaveDoc.Entity := dmSaveDoc.qryGetEntity.FieldByName('value').AsString;
         end;
      end else
      begin
         Application.MessageBox('Could not connect to Insight database','DragON');
         Application.Terminate;
      end;
   finally
      regAxiom.Free;
   end;
end;

function GetSeqNumber(sSequence: string): Integer;
begin
  with dmSaveDoc.qryTmp do
  begin
    Close;
    SQL.Clear;
    SQL.Add('SELECT ' + sSequence + '.currval');
    SQL.Add('FROM DUAL');
    ExecSQL;
    Result := Fields[0].AsInteger;
    Close;
  end;
end;

function ParseMacros(AFileName: String; ANMatter: Integer): String;
var
  LBfr: Array[0..MAX_PATH] of Char;
begin
  if(GHomePath = '') then
    GHomePath := GetEnvironmentVariable('HOMEDRIVE') + GetEnvironmentVariable('HOMEPATH');

  if(GTempPath = '') then
  begin
    GetTempPath(MAX_PATH,Lbfr);
    GTempPath := String(LBfr);
  end;

  Result := AFileName;

  Result := StringReplace(Result,C_MACRO_USERHOME,GHomePath,[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_TEMPDIR,GTempPath,[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_NMATTER,IntToStr(ANMatter),[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_FILEID, TableString('MATTER','NMATTER',IntToStr(ANMatter),'FILEID'),[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_CLIENTID, TableString('MATTER','NMATTER',IntToStr(ANMatter),'CLIENTID'),[rfReplaceAll, rfIgnoreCase]);

  Result := StringReplace(Result,C_MACRO_DATE,FormatDateTime('dd-mm-yyyy',Now()) ,[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_TIME,FormatDateTime('hh-nn-ss',Now()),[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_DATETIME,FormatDateTime('dd-mm-yyyy-hh-nn-ss',Now()),[rfReplaceAll, rfIgnoreCase]);

  Result := StringReplace(Result,C_MACRO_AUTHOR, TableString('MATTER','NMATTER',IntToStr(ANMatter),'AUTHOR'),[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_DOCSEQUENCE, ProcString('getDocSequence',ANMatter),[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_USERINITIALS, dmSaveDoc.UserID ,[rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result,C_MACRO_DOCID, dmSaveDoc.DocID,[rfReplaceAll, rfIgnoreCase]);


  if(Pos(C_MACRO_TEMPFILE,Result) > 0) then
  begin
    GetTempFileName(PChar(GTempPath),'axm',0,LBfr);
    Result := StringReplace(Result,C_MACRO_TEMPFILE,String(LBfr),[rfReplaceAll, rfIgnoreCase]);
  end;
end;

function TableString(Table, LookupField, LookupValue, ReturnField: string): string; overload;
var
  qryLookup: TUniQuery;
begin
  qryLookup := TUniQuery.Create(nil);
  qryLookup.Connection := dmSaveDoc.uniInsight;
  with qryLookup do
  begin
    SQL.Text := 'SELECT ' + ReturnField + ' FROM ' + Table + ' WHERE ' + LookupField + ' = :' + LookupField;
    Params[0].AsString := LookupValue;
    Open;
    Result := Fields[0].AsString;
    Close;
  end;
  qryLookup.Free;
end;

function TableInteger(Table, LookupField, LookupValue, ReturnField: string): integer; overload;
var
  qryLookup: TUniQuery;
begin
  qryLookup := TUniQuery.Create(nil);
  qryLookup.Connection := dmSaveDoc.uniInsight;
  with qryLookup do
  begin
    SQL.Text := 'SELECT ' + ReturnField + ' FROM ' + Table + ' WHERE ' + LookupField + ' = :' + LookupField;
    Params[0].AsString := LookupValue;
    Open;
    Result := Fields[0].AsInteger;
    Close;
  end;
  qryLookup.Free;
end;

function SystemString(sField: string): string;
begin
   with dmSaveDoc.qrySysfile do
   begin
      SQL.Text := 'SELECT ' + sField + ' FROM SYSTEMFILE';
      try
         Open;
         SystemString := FieldByName(sField).AsString;
         Close;
      except
      //
      end;
   end;
end;

function SystemInteger(sField: string): integer;
begin
   SystemInteger := 0;
   with dmSaveDoc.qrySysfile do
   begin
      SQL.Text := 'SELECT ' + sField + ' FROM SYSTEMFILE';
      try
         Open;
         SystemInteger := FieldByName(sField).AsInteger;
         Close;
      except
      //
      end;
   end;
end;

function ProcString(Proc: string; LookupValue: integer): string;
begin
   Result := IntToStr(dmSaveDoc.uniInsight.ExecProc('getDocSequence',[lookupValue]));
{  with dmSaveDoc.procTemp do
  begin
    storedProcName := Proc;
    Params.Add.DisplayName := 'matterNo';
    ParamByName('matterNo').AsInteger := LookupValue;
    ExecProc;
    Result := ParamByName('tmpVar').AsString;
  end;    }
end;

function ReportVersion: string;
var
  wVersionMajor, wVersionMinor, wVersionRelease, wVersionBuild : Word;
  VerInfoSize:  DWORD;
  VerInfo:      Pointer;
  VerValueSize: DWORD;
  VerValue:     PVSFixedFileInfo;
  Dummy:        DWORD;
begin

  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  GetMem(VerInfo, VerInfoSize);
  GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo);
  VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
  with VerValue^ do begin
    wVersionMajor := dwFileVersionMS shr 16;
    wVersionMinor := dwFileVersionMS and $FFFF;
    wVersionRelease := dwFileVersionLS shr 16;
    wVersionBuild := dwFileVersionLS and $FFFF;
  end;
  FreeMem(VerInfo, VerInfoSize);

  ReportVersion := IntToStr(wVersionMajor) + '.' + IntToStr(wVersionMinor) + '.' + IntToStr(wVersionRelease) + '.' + IntToStr(wVersionBuild);
end;

function FormExists(frmInput : TForm):boolean;
var
  iCount : integer;
  bResult : boolean;
begin
  bResult := false;
  for iCount := 0 to (Application.ComponentCount - 1) do
    if Application.Components[iCount] is TForm then
      if Application.Components[iCount] = frmInput then
        bResult:=true;
  FormExists := bResult;
end;

function IndexPath(PathText, PathLoc: string): string;
var
   iWords, i: integer;
   NewPath, sWord, sNewline, AUNCPath: string;
   LImportFile: array of string;
begin
   AUNCPath := ExpandUNCFileName(PathText);
   NewPath := SystemString(PathLoc);
   if NewPath <> '' then
   begin
      iWords := 0;
      SetLength(LImportFile,length(PathText));
      sNewline := copy(PathText,3,length(PathText));
      while TokenizePath(sNewline ,sWord) do
      begin
         LImportFile[iWords] := sWord;
         inc(iWords);
      end;

      for i := 0 to (length(LImportFile) - 1) do
      begin
         if LImportFile[i] <> '' then
            NewPath := NewPath + '/' + LImportFile[i];
      end;
      Result := NewPath;
   end
   else
      Result := AUNCPath;  //PathText;
end;

function TokenizePath(var s,w:string):boolean;
{Note that this a "destructive" getword.
  The first word of the input string s is returned in w and
  the word is deleted from the input string}
const
  delims:set of char = ['/','\'];
var
  i:integer;
begin
  w:='';
  if length(s)>0 then
  begin
    i:=1;
    while (i<length(s))  and (s[i] in delims) do inc(i);
    delete(s,1,i-1);
    i:=1;
    while (i<=length(s)) and (not (s[i] in delims)) do inc(i);
    w:=copy(s,1,i-1);
    delete(s,1,i);
  end;
  result := (length(w) >0);
end;

end.
