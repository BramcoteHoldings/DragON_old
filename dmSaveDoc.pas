unit dmSaveDoc;

interface

uses
  SysUtils, Classes, DB, DBAccess, Ora, MemDS;

type
  TDataModule2 = class(TDataModule)
    orsAxiom: TOraSession;
    OraQuery1: TOraQuery;
    OraDataSource1: TOraDataSource;
  private
    { Private declarations }
    FUserID : string;
  public
    { Public declarations }
    property UserID : string read FUserID write FUserID;
  end;

var
  DataModule2: TDataModule2;

implementation

{$R *.dfm}

end.
