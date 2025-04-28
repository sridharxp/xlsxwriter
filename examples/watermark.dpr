program watermark;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
  header_options: lxw_header_footer_options;
begin
  ReportName := 'watermark.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

  (* Set a worksheet header with the watermark image. )
  header_options.image_center := 'watermark.png';
  worksheet_set_header_opt(worksheet, '&C&[Picture]', @header_options);

  workbook_close(workbook);

  ExitCode := 0;
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
