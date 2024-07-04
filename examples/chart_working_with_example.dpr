program chart_working_with_example;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  chart: Plxw_chart;
  series: Plxw_chart_series;
  rRow1, rCol1: Word;
  ReportName: PAnsiChar;
begin
  ReportName := '.\Report\chart_line.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

  (* Create a chart object. *)
  chart := workbook_add_chart(workbook, Byte(LXW_CHART_LINE_CC));

    (* Configure the chart. *)
    series := chart_add_series(chart, nil, 'Sheet1!$A$1:$A$6');

//  series; (* Do something with series in the real examples. *)

    (* Insert the chart into the worksheet. *)
  decodecell('C1', rRow1, rCol1);
  worksheet_insert_chart(worksheet, rRow1, rCol1, chart);

  workbook_close(workbook);
end.
