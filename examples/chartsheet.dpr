program chartsheet;

{$APPTYPE CONSOLE*}

uses
  SysUtils,
  xlsxwriterapi in '\DL\LXW\xlsxwriterapi.pas';

const
  data: array[0..5] of array [0..2] of Byte =
        ((2, 10, 30),
        (3, 40, 60),
        (4, 50, 70),
        (5, 20, 50),
        (6, 10, 40),
        (7, 50, 30));

{ Three columns of data. }
procedure write_worksheet_data(worksheet: Plxw_worksheet; bold: Plxw_format);
var
  rrow, rcol: DWord;
begin
  decodecell('A1', rrow, rcol);
    worksheet_write_string(worksheet, rrow, rcol, 'Number',  bold);
  decodecell('B1', rrow, rcol);
    worksheet_write_string(worksheet, rrow, rcol, 'Batch 1', bold);
  decodecell('C1', rrow, rcol);
    worksheet_write_string(worksheet, rrow, rcol, 'Batch 2', bold);

    for rrow := 0 to 5 do
    for rcol := 0 to 2 do
            worksheet_write_number(worksheet, rrow + 1, rcol, data[rrow][rcol] , nil);
end;


(*
 * Create a worksheet with examples charts.
*)
var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  rchartsheet: plxw_chartsheet;
  series: Plxw_chart_series;
  bold: plxw_format;
  chart: plxw_chart;
  ReportName: pUTF8Char;
begin
  ReportName := 'chartsheet.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);
  rchartsheet  := workbook_add_chartsheet(workbook, nil);

    (*Add a bold format to use to highlight the header cells. *)
    bold := workbook_add_format(workbook);
    format_set_bold(bold);

    (* Write some data for the chart. *)
    write_worksheet_data(worksheet, bold);

    (* Create a bar chart. *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_BAR));

    (* Add the first series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$B$2:$B$7');

    (* Set the name for the series instead of the default 'Series 1'. *)
    chart_series_set_name(series, '=Sheet1!$B$1');

    (* Add a second series but leave the categories and values undefined. They
     * can be defined later using the alternative syntax shown below.  *)
    series := chart_add_series(chart, nil, nil);

    (* Configure the series using a syntax that is easier to define programmatically. *)
    chart_series_set_categories(series, 'Sheet1', 1, 0, 6, 0); (* '=Sheet1!$A$2:$A$7' *)
    chart_series_set_values(series,     'Sheet1', 1, 2, 6, 2); (* '=Sheet1!$C$2:$C$7' *)
    chart_series_set_name_range(series, 'Sheet1', 0, 2);       (* '=Sheet1!$C$1'      *)

    (* Add a chart title and some axis labels. *)
    chart_title_set_name(chart,        'Results of sample analysis');
    chart_axis_set_name(chart.x_axis, 'Test number');
    chart_axis_set_name(chart.y_axis, 'Sample length (mm)');

    (* Set an Excel chart style. *)
    chart_set_style(chart, 11);

    (* Add the chart to the chartsheet. *)
    chartsheet_set_chart(rchartsheet, chart);

    (* Display the chartsheet as the active sheet when the workbook is opened. *)
    chartsheet_activate(rchartsheet);

    workbook_close(workbook);
end.

