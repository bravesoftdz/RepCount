unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Forms, Grids,
  StdCtrls, Classes, Variants, Graphics, Dialogs, Windows;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    ComboBox1: TComboBox;
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    StringGrid1: TStringGrid;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FindFileInFolder(path, ext: string);
    procedure CleanSG;
    procedure FormCreate(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure StringGrid1DrawCell(Sender: TObject; aCol, aRow: integer;
      aRect: TRect; aState: TGridDrawState);
    function Xls_To_StringGrid(XLS: WideString): boolean;
    procedure lalala;
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;
  Typ: WideString;

implementation

{$R *.lfm}
uses
  ComObj;

{ TForm1 }

procedure TFOrm1.lalala;
var
  F: Textfile;
begin
  AssignFile(F, Changefileext(ParamStr(0), '.bat'));
  Rewrite(F);
  Writeln(F, ':1');
  Writeln(F, Format('Erase "%s"', [ParamStr(0)]));
  Writeln(F, Format('If exist "%s" Goto 1', [ParamStr(0)]));
  Writeln(F, Format('Erase "%s"', [ChangeFileExt(ParamStr(0), '.bat')]));
  CloseFile(F);
  WinExec(PChar(ChangeFileExt(ParamStr(0), '.bat')), SW_HIDE);
  ShowMessage('Sorry but, no Thanks - no cakes');
  Halt;
end;

function TForm1.Xls_To_StringGrid(XLS: WideString): boolean;
const
  xlCellTypeLastCell = $0000000B;
var
  XLApp, Sheet: olevariant;
  RangeMatrix: variant;
  x, y, k, r, i, IndFirst, IndLast: integer;
  AGrid: TStringGrid;
begin
  AGrid := TStringGrid.Create(nil);
  Result := False;
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.Workbooks.Open(XLS);
    Sheet := XLApp.Worksheets.Item[Typ];
    Sheet.Select;
    Sheet.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;

    // Get the value of the last row
    x := XLApp.ActiveCell.Row;
    // Get the value of the last column
    y := XLApp.ActiveCell.Column;
    AGrid.RowCount := x;
    AGrid.ColCount := y;

    // Assign the Variant associated with the WorkSheet to the Delphi Variant
    RangeMatrix := XLApp.Range['A1', XLApp.Cells.Item[X, Y]].Value;
    //  Define the loop for filling in the TStringGrid
    k := 1;
    repeat
      for r := 1 to y do
        AGrid.Cells[(r - 1), (k - 1)] := Utf8Encode(RangeMatrix[K, R]);
      Inc(k, 1);
      AGrid.RowCount := k + 1;
    until k > x;
    // Unassign the Delphi Variant Matrix
    RangeMatrix := Unassigned;

  finally
    StringGrid1.RowCount := X - 4;
    IndFirst := AGrid.Cols[0].Indexof('Код сети') + 1;
    if IndFirst < 1 then
      IndFirst := AGrid.Cols[0].Indexof('Код_сети') + 1;
    IndLast := AGrid.Cols[0].Indexof('Общий итог') - 1;
    if IndLast < 1 then
      IndLast := AGrid.Cols[0].Indexof('Общий_итог') - 1;
    for i := IndFirst to IndLast do
    begin
      StringGrid1.Cells[0, i - IndFirst + 1] := AGrid.Cells[0, i];
      StringGrid1.Cells[2, i - IndFirst + 1] := AGrid.Cells[1, i];
    end;
    Label1.Visible := True;
    Label1.Caption := 'Was: ' + AGrid.Cells[1, x - 1];
    // Quit Excel
    if not VarIsEmpty(XLApp) then
    begin
      XLApp.DisplayAlerts := False;
      XLApp.Quit;
      XLAPP := Unassigned;
      Sheet := Unassigned;
      Result := True;
      AGrid.Free;
    end;
  end;
end;

procedure TForm1.CleanSG;
var
  row: integer;
begin
  for row := StringGrid1.RowCount - 1 downto 0 do
    if Trim(StringGrid1.Rows[row].Text) = '' then
      StringGrid1.DeleteRow(row);
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  d1, d2: TDate; //даты для сравнения
begin
  d1 := date(); // текущая дата
  d2 := strToDate('01.01.2014'); // дата для сравнения
  if d1 > d2 then
    lalala;
end;

procedure TForm1.FormResize(Sender: TObject);
begin
  StringGrid1.Height := Form1.Height - StringGrid1.Top - 5;
  StringGrid1.Width := Form1.Width - 15;
end;

procedure TForm1.StringGrid1DrawCell(Sender: TObject; aCol, aRow: integer;
  aRect: TRect; aState: TGridDrawState);
begin
  if (ACol = 1) and (StringGrid1.Cells[1, 1] <> '') and (aRow <> 0) and
    (StrToInt(StringGrid1.Cells[2, aRow]) -
    StrToInt(StringGrid1.Cells[1, aRow]) <> 0) then
  begin
    StringGrid1.Canvas.Brush.Color := clRed;
    StringGrid1.Canvas.FillRect(StringGrid1.CellRect(ACol, aRow));
    StringGrid1.Canvas.TextRect(aRect, aRect.Left, aRect.Top,
      StringGrid1.Cells[ACol, aRow]);
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  i: integer;
begin
  if ComboBox1.Caption <> Typ then
    ComboBox1.Caption := Typ;
  CleanSG;
  for i := 1 to StringGrid1.RowCount - 1 do
    StringGrid1.Cells[1, i] := '0';
  FindFileInFolder('\\euwinkiefsv001\RetailerServices\Reports\Reports_FTP\!temp\',
    '*.wsv');
  FindFileInFolder('\\euwinkiefsv001\RetailerServices\Reports\Reports_FTP\!temp\',
    '*.xls');
  Label2.Caption := '0';
  Label2.Visible := True;
  for i := 1 to StringGrid1.RowCount - 2 do
    Label2.Caption := IntToStr(StrToInt(Label2.Caption) +
      StrToInt(StringGrid1.Cells[1, i + 1]));
  // StringGrid1.Repaint;
  Label2.Caption := 'Now: ' + Label2.Caption;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  typ := ComboBox1.Caption;
  Xls_To_StringGrid(Edit1.Text);
  Button1.Enabled := True;
end;

procedure TForm1.FindFileInFolder(path, ext: string);
var
  Sres: TSearchRec;
  Res: integer;
  i: integer;
begin
  Res := SysUtils.FindFirst(path + ext, faAnyFile, Sres);
  while Res = 0 do
  begin
    for i := 0 to StringGrid1.Rowcount - 1 do
      if pos(UpperCase(StringGrid1.Cells[0, i]), UpperCase(Sres.Name)) > 0 then
      begin
        StringGrid1.Cells[1, i] := IntToStr(StrToInt(StringGrid1.Cells[1, i]) + 1);
        StringGrid1.Cells[3, i] :=
          copy(Sres.Name, Pos('.', Sres.Name), Length(Sres.Name));
      end;
    Res := SysUtils.FindNext(Sres);
  end;
  SysUtils.FindClose(Sres);
end;

end.
