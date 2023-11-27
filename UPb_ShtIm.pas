unit UPb_ShtIm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ComCtrls, Buttons, OleCtrls, ExtCtrls, StdCtrls, Mask, IniFiles,
  Grids, DBGrids, DBCtrls, AxCtrls, DBClient,
  VCL.FlexCel.Core, FlexCel.XlsAdapter,
  System.Generics.Collections, Vcl.Tabs,
  System.ImageList, Vcl.ImgList, Vcl.VirtualImageList, SVGIconVirtualImageList;
type
  TfmSheetImport = class(TForm)
    Panel1: TPanel;
    sbSheet: TStatusBar;
    bbOpenSheet: TBitBtn;
    OpenDialogSprdSheet: TOpenDialog;
    gbDefineFields: TGroupBox;
    bbCancel: TBitBtn;
    gbDefineRows: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    eFromRow: TEdit;
    eToRow: TEdit;
    bbImport: TBitBtn;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    eSampleNo: TEdit;
    eShowData: TEdit;
    e238204: TEdit;
    Label10: TLabel;
    sbFindLastRow: TSpeedButton;
    Panel2: TPanel;
    Panel3: TPanel;
    gbDefineTabSheet: TGroupBox;
    cbSheetName: TComboBox;
    Label21: TLabel;
    Label22: TLabel;
    Label12: TLabel;
    e232238: TEdit;
    Splitter1: TSplitter;
    Panel4: TPanel;
    Label1: TLabel;
    Label4: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label11: TLabel;
    Label14: TLabel;
    e206204: TEdit;
    e207204: TEdit;
    e208204: TEdit;
    e207206: TEdit;
    e208206: TEdit;
    TabControl: TTabControl;
    SheetData: TStringGrid;
    Tabs: TTabSet;
    VirtualImageList1: TVirtualImageList;
    SVGIconVirtualImageList1: TSVGIconVirtualImageList;
    procedure bbOpenSheetClick(Sender: TObject);
    procedure bbImportClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bbCancelClick(Sender: TObject);
    procedure sbFindLastRowClick(Sender: TObject);
    procedure cbSheetNameChange(Sender: TObject);
    procedure TabControlChange(Sender: TObject);
    procedure SheetDataSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    Xls : TXlsFile;
    function ConvertCol2Int(AnyString : string) : integer;
    procedure FillTabs;
    procedure ClearGrid;
    procedure FillGrid(const Formatted: boolean);
    function GetStringFromCell(iRow,iCol : integer) : string;
  public
    { Public declarations }
  end;

var
  fmSheetImport: TfmSheetImport;

implementation

uses allsorts, GDW_varb, UPb_varb, UPb_dm;

{$R *.DFM}

var
  iRec, iRecCount      : integer;

procedure TfmSheetImport.bbOpenSheetClick(Sender: TObject);
var
  tmpStr    : string[3];
  i : integer;
begin
  cbSheetname.Items.Clear;
  //TabControl.Tabs.Clear;
  //cbSheetname.Items.Clear;
  OpenDialogSprdSheet.InitialDir := DataPath;
  OpenDialogSprdSheet.FileName := '';
  if OpenDialogSprdSheet.Execute then
  begin
    if (Trim(OpenDialogSprdSheet.FileName) <> '') then
    begin
      DataPath := ExtractFilePath(OpenDialogSprdSheet.FileName);
      //Open the Excel file.
      if Xls = nil then Xls := TXlsFile.Create(false);
      xls.Open(OpenDialogSprdSheet.FileName);
      FillTabs;
      Tabs.TabIndex := Xls.ActiveSheet - 1;
      cbSheetName.ItemIndex := Xls.ActiveSheet - 1;
      FillGrid(true);

      {
      FlexCelImport1.OpenFile(OpenDialogSprdSheet.FileName);
      for i := 1 to FlexCelImport1.SheetCount do
      begin
        FlexCelImport1.ActiveSheet:=i;
        TabControl.Tabs.Add(FlexCelImport1.ActiveSheetName);
        cbSheetname.Items.Add(FlexCelImport1.ActiveSheetName);
      end;
      FlexCelImport1.ActiveSheet:=1;
      TabControl.TabIndex:=FlexCelImport1.ActiveSheet-1;
      cbSheetName.ItemIndex := FlexCelImport1.ActiveSheet-1;
      Data.LoadSheet;
      Data.Zoom := 70;
      }
      bbImport.Visible := true;
      gbDefineFields.Visible := true;
      TabControl.Visible := true;
      gbDefineRows.Visible := true;
      //cbSheetname.Text := cbSheetname.Items.Strings[0];
      //cbSheetname.ItemIndex := 0;
      {
      try
        sbFindLastRowClick(Sender);
      except
      end;
      }
    end else
    begin
      bbCancelClick(Sender);
    end;
  end;
end;


procedure TfmSheetImport.FillTabs;
var
  s: integer;
begin
  Tabs.Tabs.Clear;
  cbSheetname.Items.Clear;
  for s := 1 to Xls.SheetCount do
  begin
    Tabs.Tabs.Add(Xls.GetSheetName(s));
    cbSheetname.Items.Add(Xls.GetSheetName(s));
  end;
end;

procedure TfmSheetImport.ClearGrid;
var
  r: integer;
begin
  for r := 1 to SheetData.RowCount do SheetData.Rows[r].Clear;
end;

procedure TfmSheetImport.FillGrid(const Formatted: boolean);
var
  r, c, cIndex: Integer;
  v: TCellValue;
begin
  if Xls = nil then exit;

  if (Tabs.TabIndex + 1 <= Xls.SheetCount) and (Tabs.TabIndex >= 0) then Xls.ActiveSheet := Tabs.TabIndex + 1 else Xls.ActiveSheet := 1;
  //Clear data in previous grid
  ClearGrid;
  SheetData.RowCount := 1;
  SheetData.ColCount := 1;
  //FmtBox.Text := '';

  SheetData.RowCount := Xls.RowCount + 1; //Include fixed row
  SheetData.ColCount := Xls.ColCount + 1; //Include fixed col. NOTE THAT COLCOUNT IS SLOW. We use it here because we really need it. See the Performance.pdf doc.

  if (SheetData.ColCount > 1) then SheetData.FixedCols := 1; //it is deleted when we set the width to 1.
  if (SheetData.RowCount > 1) then SheetData.FixedRows := 1;

  for r := 1 to Xls.RowCount do
  begin
    //Instead of looping in all the columns, we will just loop in the ones that have data. This is much faster.
    for cIndex := 1 to Xls.ColCountInRow(r) do
    begin
      c := Xls.ColFromIndex(r, cIndex); //The real column.
      if Formatted then
      begin
        SheetData.Cells[c, r] := Xls.GetStringFromCell(r, c);
      end
      else
      begin
        v := Xls.GetCellValue(r, c);
        SheetData.Cells[c, r] := v.ToString;
      end;
    end;
  end;

  //Fill the row headers
  for r := 1 to SheetData.RowCount - 1 do
  begin
    SheetData.Cells[0, r] := IntToStr(r);
    SheetData.RowHeights[r] := Round(Xls.GetRowHeight(r) / TExcelMetrics.RowMultDisplay(Xls));
  end;

  //Fill the column headers
  for c := 1 to SheetData.ColCount - 1 do
  begin
    SheetData.Cells[c, 0] := TCellAddress.EncodeColumn(c);
    SheetData.ColWidths[c] := Round(Xls.GetColWidth(c) / TExcelMetrics.ColMult(Xls));
  end;

  //SelectedCell(1,1);

end;

function TfmSheetImport.ConvertCol2Int(AnyString : string) : integer;
var
  itmp    : integer;
  tmpStr  : string;
  tmpChar : char;
begin
    AnyString := UpperCase(AnyString);
    tmpStr := AnyString;
    ClearNull(tmpStr);
    Result := 0;
    if (length(tmpStr) = 2) then
    begin
      tmpChar := tmpStr[1];
      itmp := (ord(tmpChar)-64)*26;
      tmpChar := tmpStr[2];
      Result := itmp+(ord(tmpChar)-64);
    end else
    begin
      tmpChar := tmpStr[1];
      Result := (ord(tmpChar)-64);
    end;
end;

procedure TfmSheetImport.bbImportClick(Sender: TObject);
var
  j     : integer;
  iCode : integer;
  i : integer;
  tmpStr : string;
  x1,x2 : double;
  v : TCellValue;
begin
  ImportSheetNumber := cbSheetName.ItemIndex + 1;
  FromRowValueString := UpperCase(eFromRow.Text);
  ToRowValueString := UpperCase(eToRow.Text);
  eSampleNo.Text := UpperCase(eSampleNo.Text);
  eShowData.Text := UpperCase(eShowData.Text);
  e238204.Text := UpperCase(e238204.Text);
  e232238.Text := UpperCase(e232238.Text);
  e206204.Text := UpperCase(e206204.Text);
  e207204.Text := UpperCase(e207204.Text);
  e208204.Text := UpperCase(e208204.Text);
  e207206.Text := UpperCase(e207206.Text);
  e208206.Text := UpperCase(e208206.Text);
  SampleNoStr := eSampleNo.Text;
  ShowDataStr := eShowData.Text;
  s238204Str := e238204.Text;
  s232238Str := e232238.Text;
  s206204Str := e206204.Text;
  s207204Str := e207204.Text;
  s208204Str := e208204.Text;
  s207206Str := e207206.Text;
  s208206Str := e208206.Text;
  {check row variables}
  iCode := 1;
  repeat
    {From Row}
    tmpStr := eFromRow.Text;
    Val(tmpStr, FromRow, iCode);
    {To Row}
    if (iCode = 0) then
    begin
      tmpStr := eToRow.Text;
      Val(tmpStr, ToRow, iCode);
    end else
    begin
      ShowMessage('Incorrect value entered for From row');
      Exit;
    end;
    if (iCode = 0) then
    begin
      if (ToRow >= FromRow) then iCode := 0
                            else iCode := -1;
    end else
    begin
      ShowMessage('Incorrect value entered for To row');
      Exit;
    end;
    if (iCode <> 0)
      then begin
        ShowMessage('Incorrect values entered for rows to import');
        Exit;
      end;
  until (iCode = 0);
  {convert input columns for variables to numeric}
  SampleNoCol := ConvertCol2Int(eSampleNo.Text);
  ShowDataCol := ConvertCol2Int(eShowData.Text);
  n238204Col := ConvertCol2Int(e238204.Text);
  n232238Col := ConvertCol2Int(e232238.Text);
  n206204Col := ConvertCol2Int(e206204.Text);
  n207204Col := ConvertCol2Int(e207204.Text);
  n208204Col := ConvertCol2Int(e208204.Text);
  n207206Col := ConvertCol2Int(e207206.Text);
  n208206Col := ConvertCol2Int(e208206.Text);
  dmUPb.cdsPbModel.EmptyDataSet;
  sbSheet.SimpleText := 'Appending new definitions';

  Xls := TXlsFile.Create(false);
  try
    //By default, FlexCel returns the formula text for the formulas, besides its calculated value.
    //If you are not interested in formula texts, you can gain a little performance by ignoring it.
    //This also works in non virtual mode.
    xls.IgnoreFormulaText := true; //bme - hard code this for this situation
    xls.VirtualMode := false;
    try
      xls.Open(OpenDialogSprdSheet.FileName);
    finally
    end;
    j := 1;
    for i := FromRow to ToRow do
    begin
      try
        dmUPb.cdsPbModel.Append;
        v := Xls.GetCellValue(i,SampleNoCol);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,SampleNoCol];
        dmUPb.cdsPbModelSample_No.AsString := tmpStr;
        v := Xls.GetCellValue(i,ShowDataCol);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,ShowDataCol];
        dmUPb.cdsPbModelShow.AsString := tmpStr;
        v := Xls.GetCellValue(i,n238204Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n238204Col];
        dmUPb.cdsPbModelU238Pb204.AsString := tmpStr;
        v := Xls.GetCellValue(i,n232238Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n232238Col];
        dmUPb.cdsPbModelTh232U238.AsString := tmpStr;
        v := Xls.GetCellValue(i,n206204Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n206204Col];
        dmUPb.cdsPbModelPb206Pb204.AsString := tmpStr;
        v := Xls.GetCellValue(i,n207204Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n207204Col];
        dmUPb.cdsPbModelPb207Pb204.AsString := tmpStr;
        v := Xls.GetCellValue(i,n208204Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n208204Col];
        dmUPb.cdsPbModelPb208Pb204.AsString := tmpStr;
        v := Xls.GetCellValue(i,n207206Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n207206Col];
        dmUPb.cdsPbModelPb207Pb206.AsString := tmpStr;
        v := Xls.GetCellValue(i,n208206Col);
        tmpStr := v.ToString;
        //tmpStr := FlexCelImport1.CellValue[i,n208206Col];
        dmUPb.cdsPbModelPb208Pb206.AsString := tmpStr;
        //ShowMessage(IntToStr(i)+'**'+tmpStr);
        dmUPb.cdsPbModel.Post;
      except
      end;
    end;
    dmUPb.cdsPbModel.SaveToFile(PbModelFileName,dfXML);
    dmUPb.cdsPbModel.LoadFromFile(PbModelFileName);
  finally
    Xls.Free;
  end;
  if (iCode = 0) then
  begin
    ModalResult := mrOK;
  end else
  begin
    ModalResult := mrNone;
  end;
end;

procedure TfmSheetImport.FormShow(Sender: TObject);
var
  i, j : integer;
begin
  bbImport.Visible := false;
  gbDefineFields.Visible := false;
  {
  SprdSheet.Visible := false;
  }
  gbDefineRows.Visible := false;
  eFromRow.Text := FromRowValueString;
  eToRow.Text := ToRowValueString;
  eSampleNo.Text := SampleNoStr;
  eShowData.Text := ShowDataStr;
  e238204.Text := s238204Str;
  e232238.Text := s232238Str;
  e206204.Text := s206204Str;
  e207204.Text := s207204Str;
  e208204.Text := s208204Str;
  e207206.Text := s207206Str;
  e208206.Text := s208206Str;
  bbOpenSheetClick(Sender);
end;

procedure TfmSheetImport.sbFindLastRowClick(Sender: TObject);
var
  iCode : integer;
  tmpStr : string;
  i,j : integer;
  v : TCellValue;
begin
  ImportSheetNumber := cbSheetName.ItemIndex;
  //ShowMessage('ImportSheetNumber = '+Int2Str(ImportSheetNumber));
  ToRow := 0;
  try
    i := 2;
    j := ConvertCol2Int(eSampleNo.Text);
    //ShowMessage('i = '+Int2Str(i));
    //ShowMessage('j = '+Int2Str(j));
    //ShowMessage('tmpStr = #'+tmpStr+'#');
    v := Xls.GetCellValue(i,j);
    tmpStr := v.ToString;
    //tmpStr := FlexCelImport1.CellValue[i,j];
    //('tmpStr = *'+tmpStr+'*');
    while (tmpStr <> '') do
    begin
      if (i > ToRow) then ToRow := i;
      i := i + 1;
      v := Xls.GetCellValue(i,j);
      tmpStr := v.ToString;
      //tmpStr := FlexCelImport1.CellValue[i,j];
    end;
    eToRow.Text := IntToStr(ToRow);
  except
    //MessageDlg('Error reading data for main variable grade',mtwarning,[mbOK],0);
  end;
end;

procedure TfmSheetImport.SheetDataSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
begin
  //SelectedCell(aCol, aRow);
  CanSelect := true;
end;

procedure TfmSheetImport.TabControlChange(Sender: TObject);
begin
  {
  Data.ApplySheet;
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  }
  cbSheetname.ItemIndex := TabControl.TabIndex;
  {
  Data.Zoom := 70;
  Data.LoadSheet;
  }
  //sbFindLastRowClick(Sender);
end;

function TfmSheetImport.GetStringFromCell(iRow,iCol : integer) : string;
begin
  Result := Xls.GetStringFromCell(iRow,iCol);
end;

procedure TfmSheetImport.FormDestroy(Sender: TObject);
begin
  //FreeAndNil(Xls);
  //inherited;
end;

procedure TfmSheetImport.cbSheetNameChange(Sender: TObject);
begin
  //ImportSheetNumber := cbSheetName.ItemIndex+1;

  Tabs.TabIndex := cbSheetname.ItemIndex;
  ClearGrid;
  FillGrid(true);
  {
  FlexCelImport1.ActiveSheet:= TabControl.TabIndex+1;
  Data.ApplySheet;
  Data.Zoom := 70;
  Data.LoadSheet;
  }
end;

procedure TfmSheetImport.bbCancelClick(Sender: TObject);
begin
  ModalResult := mrNone;
  bbImport.ModalResult := mrNone;
  Close;
end;




end.
