program comments1;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
begin
  ReportName := 'comments1.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);
  worksheet_write_string( worksheet, 0, 0, 'Hello' , nil);
  worksheet_write_comment(worksheet, 0, 0, 'This is a comment');
  workbook_close(workbook);
end.
