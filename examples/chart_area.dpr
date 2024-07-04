program chartsheet;

{$APPTYPE CONSOLE*}

uses
  SysUtils,
  xlsxwriterapi in '\DL\LXW\xlsxwriterapi.pas';

const
  data: array[0..5] of array [0..2] of Byte=
        (* Three columns of data. *)
        ((2, 40, 30),
        (3, 40,  25),
        (4, 50,  30),
        (5, 30,  10),
        (6, 25,   5),
        (7, 50, 10));

{ Three columns of data. }
procedure write_worksheet_data(worksheet: Plxw_worksheet; bold: Plxw_format);
var
  rrow, rcol: Word;
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
  series: Plxw_chart_series;
  bold: plxw_format;
  chart: Plxw_chart;
  rRow1, rCol1: Word;
  ReportName: pUTF8Char;
begin
  ReportName := 'chart_area.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

    (*Add a bold format to use to highlight the header cells. *)
    bold := workbook_add_format(workbook);
    format_set_bold(bold);

    (* Write some data for the chart. *)
    write_worksheet_data(worksheet, bold);

    (*
     * Chart 1. Create a area chart.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_AREA));

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
    chart_axis_set_name(chart^.x_axis, 'Test number');
    chart_axis_set_name(chart^.y_axis, 'Sample length (mm)');

    (* Set an Excel chart style. *)
    chart_set_style(chart, 11);

    (* Insert the chart into the worksheet. *)
    decodecell('E2', rRow1, rCol1);
    worksheet_insert_chart(worksheet, rRow1, rCol1, chart);


    (*
     * Chart 2. Create a stacked area chart.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_AREA_STACKED));

    (* Add the first series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$B$2:$B$7');

    (* Set the name for the series instead of the default 'Series 1'. *)
    chart_series_set_name(series, '=Sheet1!$B$1');

    (* Add the second series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$C$2:$C$7');

    (* Set the name for the series instead of the default 'Series 2'. *)
    chart_series_set_name(series, '=Sheet1!$C$1');

    (* Add a chart title and some axis labels. *)
    chart_title_set_name(chart,        'Results of sample analysis');
    chart_axis_set_name(chart^.x_axis, 'Test number');
    chart_axis_set_name(chart^.y_axis, 'Sample length (mm)');

    (* Set an Excel chart style. *)
    chart_set_style(chart, 12);

    (* Insert the chart into the worksheet. *)
    decodecell('E18', rRow1, rCol1);
    worksheet_insert_chart(worksheet, rRow1, rCol1, chart);


    (*
     * Chart 3. Create a percent stacked area chart.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_AREA_STACKED_PERCENT));

    (* Add the first series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$B$2:$B$7');

    (* Set the name for the series instead of the default 'Series 1'. *)
    chart_series_set_name(series, '=Sheet1!$B$1');

    (* Add the second series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$C$2:$C$7');

    (* Set the name for the series instead of the default 'Series 2'. *)
    chart_series_set_name(series, '=Sheet1!$C$1');

    (* Add a chart title and some axis labels. *)
    chart_title_set_name(chart,        'Results of sample analysis');
    chart_axis_set_name(chart^.x_axis, 'Test number');
    chart_axis_set_name(chart^.y_axis, 'Sample length (mm)');

    (* Set an Excel chart style. *)
    chart_set_style(chart, 13);

    (* Insert the chart into the worksheet. *)
    decodecell('E34', rRow1, rCol1);
    worksheet_insert_chart(worksheet, rRow1, rCol1, chart);

    workbook_close(workbook);
end.

