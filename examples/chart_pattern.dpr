program chart_pattern;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  sHELLaPI,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
  chart: Plxw_chart;
  bold: Plxw_format;
  series1: Plxw_chart_series;
  series2: Plxw_chart_series;
  pattern1: lxw_chart_pattern;
  pattern2: lxw_chart_pattern;
  line1: lxw_chart_line;
  line2: lxw_chart_line;
  Row1, Col1: DWord;
  Row2, Col2: DWord;
begin
  ReportName := 'chart_pattern.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);


    (* Add a bold format to use to highlight the header cells. )
    bold := workbook_add_format(workbook);
    format_set_bold(bold);

    (* Write some data for the chart. *)
    worksheet_write_string(worksheet, 0, 0, 'Shingle', bold);
    worksheet_write_number(worksheet, 1, 0, 105,       nil);
    worksheet_write_number(worksheet, 2, 0, 150,       nil);
    worksheet_write_number(worksheet, 3, 0, 130,       nil);
    worksheet_write_number(worksheet, 4, 0, 90,        nil);
    worksheet_write_string(worksheet, 0, 1, 'Brick',   bold);
    worksheet_write_number(worksheet, 1, 1, 50,        nil);
    worksheet_write_number(worksheet, 2, 1, 120,       nil);
    worksheet_write_number(worksheet, 3, 1, 100,       nil);
    worksheet_write_number(worksheet, 4, 1, 110,       nil);

    (* Create a chart object. )
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Configure the chart. *)
    series1 := chart_add_series(chart, nil, 'Sheet1!$A$2:$A$5');
    series2 := chart_add_series(chart, nil, 'Sheet1!$B$2:$B$5');

    chart_series_set_name(series1, '=Sheet1!$A$1');
    chart_series_set_name(series2, '=Sheet1!$B$1');

    chart_title_set_name(chart,        'Cladding types');
    chart_axis_set_name(chart.x_axis, 'Region');
    chart_axis_set_name(chart.y_axis, 'Number of houses');


    (* Configure an add the chart series patterns. )
    pattern1._type := Byte(LXW_CHART_PATTERN_SHINGLE);
                                  pattern1.fg_color := $804000;
                                  pattern1.bg_color := $C68C53;

    pattern2._type := Byte(LXW_CHART_PATTERN_HORIZONTAL_BRICK);
                                  pattern2.fg_color := $B30000;
                                  pattern2.bg_color := $FF6666;

    chart_series_set_pattern(series1, @pattern1);
    chart_series_set_pattern(series2, @pattern2);

    (* Configure and set the chart series borders. *)
    line1.color := $804000;
    line2.color := $b30000;

    chart_series_set_line(series1, @line1);
    chart_series_set_line(series2, @line2);

    (* Widen the gap between the series/categories. *)
    chart_set_series_gap(chart, 70);

    // Insert the chart into the worksheet.
    decodecell('D2', Row1, Col1);
    worksheet_insert_chart(worksheet, Row1, Col1, chart);

  lxw_error(ExitCode) := workbook_close(workbook);
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
