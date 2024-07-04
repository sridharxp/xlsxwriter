program background;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';
var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
begin
  ReportName := 'background.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

   worksheet_set_background(worksheet, 'logo.png');
end.
