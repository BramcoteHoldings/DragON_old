unit MatterSearch;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData,
  cxDataStorage, cxEdit, cxDBData, cxMaskEdit, cxLookAndFeelPainters,
  StdCtrls, cxButtons, cxContainer, cxTextEdit, LMDCustomControl,
  LMDCustomPanel, LMDCustomBevelPanel, LMDSimplePanel, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses,
  cxControls, cxGridCustomView, cxGrid, DBAccess, Menus, cxCheckBox, DB,
  ExtCtrls, LMDControl, cxLookAndFeels, dxCore, cxNavigator;

type
  TfrmMtrSearch = class(TForm)
    dbgrMatters: TcxGrid;
    vMatters: TcxGridDBTableView;
    vMattersPARENT: TcxGridDBColumn;
    vMattersTITLE: TcxGridDBColumn;
    vMattersFILEID: TcxGridDBColumn;
    vMattersSHORTDESCR: TcxGridDBColumn;
    vMattersLONGDESCR: TcxGridDBColumn;
    vMattersNMATTER: TcxGridDBColumn;
    vMattersPARTNER: TcxGridDBColumn;
    vMattersAUTHOR: TcxGridDBColumn;
    vMattersTYPE: TcxGridDBColumn;
    vMattersCLIENTID: TcxGridDBColumn;
    vMattersARCHIVENUM: TcxGridDBColumn;
    vMattersSUBTYPE: TcxGridDBColumn;
    vMattersSTATUS: TcxGridDBColumn;
    vMattersJURISDICTION: TcxGridDBColumn;
    vMattersMATTERSTATUS2: TcxGridDBColumn;
    dbgrMattersLevel1: TcxGridLevel;
    LMDSimplePanel1: TLMDSimplePanel;
    Label8: TLabel;
    tbClientSearch: TcxTextEdit;
    tbFileSearch: TcxTextEdit;
    Label31: TLabel;
    btnOk: TcxButton;
    bnCancel: TcxButton;
    cbShowRecentlyAccessed: TcxCheckBox;
    tmrSearch: TTimer;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure vMattersDblClick(Sender: TObject);
    procedure cbShowRecentlyAccessedClick(Sender: TObject);
    procedure vMattersColumnHeaderClick(Sender: TcxGridTableView;
      AColumn: TcxGridColumn);
    procedure tbClientSearchPropertiesChange(Sender: TObject);
    procedure tmrSearchTimer(Sender: TObject);    
    procedure EnableTimer(Sender: TObject);
  private
    { Private declarations }
    sOrderBy: string;
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
    procedure MakeSql(bSearch: boolean = False);
  end;

var
  frmMtrSearch: TfrmMtrSearch;

implementation

{$R *.dfm}

uses
   OraCall, SaveDoc;

procedure TfrmMtrSearch.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
//   dmSaveDoc.qryMatters.Active := False;
end;

procedure TfrmMtrSearch.vMattersDblClick(Sender: TObject);
begin
   btnOk.Click;
end;

procedure TfrmMtrSearch.MakeSql(bSearch: boolean);
var
   sSQL,sTables, sWhereClause, sAND: string;
begin
   try
      dmSaveDoc.qryMatters.Close;
      dmSaveDoc.qryMatters.SQL.Clear;
       sAND := ' AND ';
      sSQL := 'select * ';
      sTables := 'from matter ';
      sWhereClause := ' where closed = 0 AND entity = :P_Entity ';
      if bSearch then
      begin
         if tbClientSearch.Text <> '' then
         begin
            sWhereClause := sWhereClause + sAND + ' UPPER(MATTER.TITLE) LIKE ' + QuotedStr('%' + Uppercase(tbClientSearch.Text) + '%');
            if cbShowRecentlyAccessed.Checked then
            begin
               sWhereClause := sWhereClause + ' AND O.AUTHOR = :P_Author AND O.TYPE = :P_Type AND O.CODE = MATTER.FILEID ';
               sTables := sTables + ', OPENLIST O ';
            end;

            sAND := ' AND ';
         end;
         if tbFileSearch.Text <> '' then
         begin
            sWhereClause := sWhereClause + sAND + 'MATTER.FILEID LIKE ' + QuotedStr(tbFileSearch.Text + '%');
            sAND := ' AND ';
         end;
      end
      else
      begin
         if cbShowRecentlyAccessed.Checked then
         begin
            sWhereClause := sWhereClause + ' AND upper(O.AUTHOR) = upper(:P_Author) AND O.TYPE = :P_Type AND O.CODE = MATTER.FILEID ';
            sTables := sTables + ', OPENLIST O ';
         end;
      end;
      
      dmSaveDoc.qryMatters.SQL.Text := sSQL + sTables + sWhereClause + sOrderBy;
      dmSaveDoc.qryMatters.Prepare;
      if cbShowRecentlyAccessed.Checked then
      begin
         dmSaveDoc.qryMatters.ParamByName('P_TYPE').AsString := 'MATTER';
         dmSaveDoc.qryMatters.ParamByName('P_Author').AsString := dmSaveDoc.UserID;
      end;
      dmSaveDoc.qryMatters.ParamByName('P_Entity').AsString := dmSaveDoc.Entity;
      dmSaveDoc.qryMatters.Open;
   Except
//      Application.MessageBox('An error occurred.','Axiom',);
//      Application.Free;
   end;
end;

procedure TfrmMtrSearch.cbShowRecentlyAccessedClick(Sender: TObject);
begin
   MakeSQL();
end;

procedure TfrmMtrSearch.vMattersColumnHeaderClick(Sender: TcxGridTableView;
  AColumn: TcxGridColumn);
begin
   sOrderBy := ' ORDER BY ';

   sOrderBy := sOrderBy + TcxGridDBColumn(AColumn).DataBinding.FieldName;

   if  AColumn.SortOrder = soNone then
   begin
      sOrderBy := sOrderBy + ' ASC';
      AColumn.SortOrder := soAscending;
   end
   else if AColumn.SortOrder = soAscending then
   begin
      sOrderBy := sOrderBy + ' ASC';
   end
   else
   begin
      sOrderBy := sOrderBy + ' DESC';
   end;

   MakeSql();
end;

procedure TfrmMtrSearch.tbClientSearchPropertiesChange(Sender: TObject);
begin
   EnableTimer(Sender);
end;

procedure TfrmMtrSearch.EnableTimer(Sender: TObject);
begin
   tmrSearch.Enabled := true;
end;

procedure TfrmMtrSearch.tmrSearchTimer(Sender: TObject);
begin
   tmrSearch.Enabled := false;
   if ((tbFileSearch.Text = '') and (tbClientSearch.Text = '')) then
      MakeSQL()
   else
      MakeSQL(True);
end;

procedure TfrmMtrSearch.CreateParams(var Params: TCreateParams);
begin
  inherited CreateParams(Params);
  with Params do begin
    ExStyle := ExStyle or WS_EX_TOPMOST;
    WndParent := GetDesktopwindow;
  end;
end;

end.
