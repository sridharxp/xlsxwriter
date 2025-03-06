program utf8;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  sHELLaPI,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
begin
  ReportName := '.\Report\utf8.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

  worksheet_write_string(worksheet, 2, 1, 'some utf8 string', nil);

  lxw_error(ExitCode) := workbook_close(workbook);
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
