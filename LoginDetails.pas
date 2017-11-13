unit LoginDetails;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Registry;

type
  TfrmLoginSetup = class(TForm)
    GroupBox1: TGroupBox;
    edUserName: TEdit;
    edPassword: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Database: TGroupBox;
    Label3: TLabel;
    edServerName: TEdit;
    edDatabase: TEdit;
    Label4: TLabel;
    Label5: TLabel;
    edPort: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmLoginSetup: TfrmLoginSetup;

implementation

uses SaveDoc;

{$R *.dfm}

procedure TfrmLoginSetup.Button1Click(Sender: TObject);
var
  regAxiom: TRegistry;
  sRegistryRoot: string;
  bSuccess: boolean;
begin
   bSuccess := False;
   sRegistryRoot := 'Software\Colateral\Axiom';
   regAxiom := TRegistry.Create;
   try
      regAxiom.RootKey := HKEY_CURRENT_USER;
      if regAxiom.OpenKey(sRegistryRoot+'\InsightDocSave', True) then
      begin
         regAxiom.WriteString('Net','Y');
         regAxiom.WriteString('Server Name',edServerName.Text+':'+edPort.Text+':'+edDatabase.Text);
         regAxiom.WriteString('User Name',edUserName.Text);
         regAxiom.WriteString('Password',edPassword.Text);
         regAxiom.CloseKey;

         dmSaveDoc.uniInsight.Server := edServerName.Text+':'+edPort.Text+':'+edDatabase.Text;
         dmSaveDoc.uniInsight.Username := edUserName.Text;
         dmSaveDoc.uniInsight.Password := edPassword.Text;
         try
            dmSaveDoc.uniInsight.Connect;
            dmSaveDoc.uniInsight.Disconnect;
            bSuccess := True;
//            MessageBox(Self.Handle,'Connection Test Succesfull.','DragON',MB_OK+MB_ICONINFORMATION);
         except
            bSuccess := False;
         end;
      end;
   finally
      regAxiom.Free;
   end;
   if bSuccess then
      Close;
end;

procedure TfrmLoginSetup.FormCreate(Sender: TObject);
var
   LRegAxiom: TRegistry;
   LoginStr, s: string;
begin
   LregAxiom := TRegistry.Create;
   try
      LregAxiom.RootKey := HKEY_CURRENT_USER;
      LregAxiom.OpenKey('Software\Colateral\Axiom\InsightDocSave', False);

      s := Copy(LregAxiom.ReadString('Server Name'),1,Pos(':',LregAxiom.ReadString('Server Name'))-1);
      LoginStr := Copy(LregAxiom.ReadString('Server Name'),Pos(':',LregAxiom.ReadString('Server Name'))+1, Length(LregAxiom.ReadString('Server Name')) - Pos(':',LregAxiom.ReadString('Server Name')) );
      if s <> '' then
         edServerName.Text := s;

      s := Copy(LoginStr,1,Pos(':',LoginStr)-1);
      LoginStr := Copy(LoginStr,Pos(':',LoginStr)+1, Length(LoginStr));
      if s <> '' then
         edPort.Text := s;

      s := LoginStr;
      if s <> '' then
         edDatabase.Text := s;

      edUserName.Text := LregAxiom.ReadString('User Name');
      edPassword.Text := LregAxiom.ReadString('Password');
   finally
     LregAxiom.Free;
   end;
end;

end.
