program chart_data_table;

{$APPTYPE CONSOLE}

uses
  FastMM4 in '..\..\..\DL\FastMM4_4993\FastMM4.pas',
  FastMM4Messages in '..\..\..\DL\FastMM4_4993\FastMM4Messages.pas',
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

const
  data : array [0..5] of array [0..2] of Byte =
        (* Three columns of data. *)
        ((2, 10, 30),
        (3, 40, 60),
        (4, 50, 70),
        (5, 20, 50),
        (6, 10, 40),
        (7, 50, 30));
(*
 * Write some data to the worksheet.
 *)
procedure write_worksheet_data(worksheet: Plxw_worksheet; bold: Plxw_format);
var
  row, col: Integer;
  rRow, rCol: DWord;
begin
    decodecell('A1', rRow, rCol);
    worksheet_write_string(worksheet, rRow, rCol, 'Number',  bold);
    decodecell('B1', rRow, rCol);
    worksheet_write_string(worksheet, rRow, rCol, 'Batch 1', bold);
    decodecell('C1', rRow, rCol);
    worksheet_write_string(worksheet, rRow, rCol, 'Batch 2', bold);

    for row := 0 to 5 do
        for col := 0 to 2 do
            worksheet_write_number(worksheet, row + 1, col, data[row][col] , nil);
end;

(*
 * Create a worksheet with examples charts.
 *)

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  series: Plxw_chart_series;
  bold: Plxw_format;
  chart: Plxw_chart;
  rRow, rCol: DWord;
  ReportName: PAnsiChar;
begin
  ReportName := '.\Report\chart_data_table.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

    (* Add a bold format to use to highlight the header cells. *)
    bold := workbook_add_format(workbook);
    format_set_bold(bold);

    (* Write some data for the chart. *)
    write_worksheet_data(worksheet, bold);


    (*
     * Chart 1. Create a column chart with a data table.
     *)
     chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

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
    chart_title_set_name(chart,        'Chart with Data Table');
    chart_axis_set_name(chart^.x_axis, 'Test number');
    chart_axis_set_name(chart^.y_axis, 'Sample length (mm)');

    (* Set a default data table on the X-axis. *)
    chart_set_table(chart);

    (* Insert the chart into the worksheet. *)
    decodecell('E2', rRow, rCol);
    worksheet_insert_chart(worksheet, rRow, rCol, chart);


    (*
     * Chart 2. Create a column chart with a data table and legend keys.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add the first series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$B$2:$B$7');

    (* Set the name for the series instead of the default 'Series 1'. *)
    chart_series_set_name(series, '=Sheet1!$B$1');

    (* Add the second series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7', '=Sheet1!$C$2:$C$7');

    (* Set the name for the series instead of the default 'Series 2'. *)
    chart_series_set_name(series, '=Sheet1!$C$1');

    (* Add a chart title and some axis labels. *)
    chart_title_set_name(chart,        'Data Table with legend keys');
    chart_axis_set_name(chart^.x_axis, 'Test number');
    chart_axis_set_name(chart^.y_axis, 'Sample length (mm)');

    (* Set a data table on the X-axis with the legend keys shown. *)
    chart_set_table(chart);
    chart_set_table_grid(chart, Byte(LXW_TRUE), Byte(LXW_TRUE), Byte(LXW_TRUE), Byte(LXW_TRUE));

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));


    (* Insert the chart into the worksheet. *)
    decodecell('E18', rRow, rCol);
    worksheet_insert_chart(worksheet, rRow, rCol, chart);


  workbook_close(workbook);
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
