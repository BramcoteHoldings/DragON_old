unit SaveDoc;

interface

uses
  MapiDefs, SysUtils, Classes, DB, OraSmart, Ora, Dialogs,
  MapiUtil, MapiTags, Windows, DBAccess, Uni, MemDS;

type
  TMessage = class(TObject)
  private
    FMessage: IMessage;
    FAttachments: TInterfaceList;
    FAttachmentsLoaded: boolean;
    function GetAttachments: TInterfaceList;
  public
    constructor Create(const AMessage: IMessage);
    destructor Destroy; override;
    property Msg: IMessage read FMessage;
    property Attachments: TInterfaceList read GetAttachments;
  end;

  TdmSaveDoc = class(TDataModule)
    uniInsight: TUniConnection;
    qryEmps: TUniQuery;
    qryMatters: TUniQuery;
    dsMatters: TUniDataSource;
    qryGetSeq: TUniQuery;
    qryMatterAttachments: TUniQuery;
    qryGetMatter: TUniQuery;
    qryGetEntity: TUniQuery;
    qryPrecCategory: TUniQuery;
    dsPrecCategory: TUniDataSource;
    qryTmp: TUniQuery;
    qrySysFile: TUniQuery;
    procTemp: TUniStoredProc;
    qryPrecClassification: TUniQuery;
    dsPrecClassification: TUniDataSource;
    dsEmployee: TUniDataSource;
    qryEmployee: TUniQuery;
    qrySaveEmailAttachments: TUniQuery;
    procedure qryMatterAttachmentsNewRecord(DataSet: TDataSet);
    procedure uniInsightError(Sender: TObject; E: EDAError;
      var Fail: Boolean);
    procedure qrySaveEmailAttachmentsNewRecord(DataSet: TDataSet);
  private
    { Private declarations }
    FUserID : string;
    FEntity : string;
    FDocID   : string;
    FAttDocID: string;
  public
    { Public declarations }
    property UserID : string read FUserID write FUserID;
    property Entity : string read FEntity write FEntity;
    property DocID  : string read FDocID write FDocID;
    property AttDocID  : string read FAttDocID write FAttDocID;
  end;

var
  dmSaveDoc: TdmSaveDoc;

implementation

{$R *.dfm}

constructor TMessage.Create(const AMessage: IMessage);
begin
  FMessage := AMessage;
  FAttachments := TInterfaceList.Create;
end;

destructor TMessage.Destroy;
begin
  FAttachments.Free;
  FMessage := nil;
  inherited Destroy;
end;

{$RANGECHECKS OFF}
function TMessage.GetAttachments: TInterfaceList;
const
  AttachmentTags: packed record
    Values: ULONG;
    PropTags: array[0..0] of ULONG;
  end = (Values: 1; PropTags: (PR_ATTACH_NUM));

var
  Table: IMAPITable;
  Rows: PSRowSet;
  i: integer;
  Attachment: IAttach;
begin
  if (not FAttachmentsLoaded) then
  begin
    FAttachmentsLoaded := True;
    (*
    ** Get list of attachment interfaces from message
    **
    ** Note: This will only succeed the first time it is called for an IMessage.
    ** The reason is probably that it is illegal (according to MSDN) to call
    ** IMessage.OpenAttach more than once for a given attachment. However, it
    ** might also be a bug in my code, but, whatever the reason, the solution is
    ** beyond the scope of this demo.
    ** Let me know if you find a solution.
    *)
    if (Succeeded(FMessage.GetAttachmentTable(0, Table))) then
    begin
      if (Succeeded(HrQueryAllRows(Table, PSPropTagArray(@AttachmentTags), nil, nil, 0, Rows))) then
        try
          for i := 0 to integer(Rows.cRows)-1 do
          begin
            // Get one attachment at a time
            if (Rows.aRow[i].lpProps[0].ulPropTag and PROP_TYPE_MASK <> PT_ERROR) and
              (Succeeded(FMessage.OpenAttach(Rows.aRow[i].lpProps[0].Value.l, IAttach, 0, Attachment))) then
              FAttachments.Add(Attachment);
          end;

        finally
          FreePRows(Rows);
        end;
      Table := nil;
    end;
  end;
  Result := FAttachments;
end;
{$RANGECHECKS ON}

procedure TdmSaveDoc.qryMatterAttachmentsNewRecord(DataSet: TDataSet);
begin
   dmSaveDoc.qryGetSeq.ExecSQL;
   DocID := dmSaveDoc.qryGetSeq.FieldByName('nextdoc').AsString;
   dmSaveDoc.qryMatterAttachments.FieldByName('docid').AsString := DocID;
end;

procedure TdmSaveDoc.uniInsightError(Sender: TObject; E: EDAError;
  var Fail: Boolean);
begin
   case E.ErrorCode of
      1005: Fail := False;
   else
      MessageDlg('Insight Database Error:'#13#10 + e.Message, mtError, [mbOK], 0);
   end;
end;

procedure TdmSaveDoc.qrySaveEmailAttachmentsNewRecord(DataSet: TDataSet);
begin
   dmSaveDoc.qryGetSeq.ExecSQL;
   AttDocID := dmSaveDoc.qryGetSeq.FieldByName('nextdoc').AsString;
   dmSaveDoc.qrySaveEmailAttachments.FieldByName('docid').AsString := AttDocID;
end;

end.
