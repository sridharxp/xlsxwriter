program chart_clustered;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';


(*
 * Write some data to the worksheet.
 *)
procedure write_worksheet_data(worksheet: Plxw_worksheet; bold: Plxw_format);
begin

    worksheet_write_string(worksheet, 0, 0, 'Types',      bold);
    worksheet_write_string(worksheet, 1, 0, 'Type 1',     nil);
    worksheet_write_string(worksheet, 4, 0, 'Type 2',     nil);

    worksheet_write_string(worksheet, 0, 1, 'Sub Type',   bold);
    worksheet_write_string(worksheet, 1, 1, 'Sub Type A', nil);
    worksheet_write_string(worksheet, 2, 1, 'Sub Type B', nil);
    worksheet_write_string(worksheet, 3, 1, 'Sub Type C', nil);
    worksheet_write_string(worksheet, 4, 1, 'Sub Type D', nil);
    worksheet_write_string(worksheet, 5, 1, 'Sub Type E', nil);

    worksheet_write_string(worksheet, 0, 2, 'Value 1',    bold);
    worksheet_write_number(worksheet, 1, 2, 5000,         nil);
    worksheet_write_number(worksheet, 2, 2, 2000,         nil);
    worksheet_write_number(worksheet, 3, 2, 250,          nil);
    worksheet_write_number(worksheet, 4, 2, 6000,         nil);
    worksheet_write_number(worksheet, 5, 2, 500,          nil);

    worksheet_write_string(worksheet, 0, 3, 'Value 2',    bold);
    worksheet_write_number(worksheet, 1, 3, 8000,         nil);
    worksheet_write_number(worksheet, 2, 3, 3000,         nil);
    worksheet_write_number(worksheet, 3, 3, 1000,         nil);
    worksheet_write_number(worksheet, 4, 3, 6000,         nil);
    worksheet_write_number(worksheet, 5, 3, 300,          nil);

    worksheet_write_string(worksheet, 0, 4, 'Value 3',    bold);
    worksheet_write_number(worksheet, 1, 4, 6000,         nil);
    worksheet_write_number(worksheet, 2, 4, 4000,         nil);
    worksheet_write_number(worksheet, 3, 4, 2000,         nil);
    worksheet_write_number(worksheet, 4, 4, 6500,         nil);
    worksheet_write_number(worksheet, 5, 4, 200,          nil);
end;


var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  chart: Plxw_chart;
  bold: Plxw_format;
  rRow1, rCol1: Word;
  ReportName: PAnsiChar;
begin
  ReportName := '.\Report\chart_clustered2.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);
    chart     := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a bold format to use to highlight the header cells. *)
    bold := workbook_add_format(workbook);
    format_set_bold(bold);

    (* Write some data for the chart. *)
    write_worksheet_data(worksheet, bold);

    (*
     * Configure the series. Note, that the categories are 2D ranges (from
     * column A to column B). This creates the clusters. The series are shown
     * as formula strings for clarity but you can also use variables with the
     * chart_series_set_categories() and chart_series_set_values()
     * functions. See the docs.
     *)
    chart_add_series(chart,
                     '=Sheet1!$A$2:$B$6',
                     '=Sheet1!$C$2:$C$6');

    chart_add_series(chart,
                     '=Sheet1!$A$2:$B$6',
                     '=Sheet1!$D$2:$D$6');

    chart_add_series(chart,
                     '=Sheet1!$A$2:$B$6',
                     '=Sheet1!$E$2:$E$6');

    (* Set an Excel chart style. *)
    chart_set_style(chart, 37);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('G3', rRow1, rCol1);
    worksheet_insert_chart(worksheet, rRow1, rCol1, chart);

  workbook_close(workbook);
end.
