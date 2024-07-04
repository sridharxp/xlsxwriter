program chart;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

procedure write_worksheet_data(var aworksheet: Plxw_worksheet);
var
  row, col, i: integer;
begin
  i := 1;
  for row := 0 to 4 do
  for col := 0 to 2 do
  begin
    worksheet_write_number(aworksheet, row, col, i, nil);
    Inc(i);
  end;
end;

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  row, col, i: integer;
  chart_: plxw_chart;
  font: lxw_chart_font;
  rRow, rCol: DWord;
  ReportName: PAnsiChar;
begin
    ReportName := 'chart.xlsx';
    workbook  := workbook_new(ReportName);
    worksheet := workbook_add_worksheet(workbook, nil);

    (* Write some data for the chart. *)
    write_worksheet_data(worksheet);

    (* Create a chart object. *)
    chart_ := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Configure the chart. In simplest case we just add some value data
     * series. The NULL categories will default to 1 to 5 like in Excel.
     *)
    chart_add_series(chart_, nil, 'Sheet1!$A$1:$A$5');
    chart_add_series(chart_, nil, 'Sheet1!$B$1:$B$5');
    chart_add_series(chart_, nil, 'Sheet1!$C$1:$C$5');

  font.bold := Byte(LXW_EXPLICIT_FALSE);
  font.color := Cardinal(LXW_COLOR_BLUE);
  chart_title_set_name(chart_, 'Year End Results');
  chart_title_set_name_font(chart_, @font);

    (* Insert the chart into the worksheet. *)
    decodecell('B7', rRow, rCol);
    worksheet_insert_chart(worksheet, rRow, rCol, chart_);

  workbook_close(workbook);
end.
