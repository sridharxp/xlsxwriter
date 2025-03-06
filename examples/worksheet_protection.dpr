program worksheet_protection;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
  unlocked: Plxw_format;
  hidden: Plxw_format;
begin
  ReportName := '.protection.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

    unlocked := workbook_add_format(workbook);
    format_set_unlocked(unlocked);

    hidden := workbook_add_format(workbook);
    format_set_hidden(hidden);

    (* Widen the first column to make the text clearer. *)
    worksheet_set_column(worksheet, 0, 0, 40, nil);

    (* Turn worksheet protection on without a password. *)
    worksheet_protect(worksheet, nil, nil);


    (* Write a locked, unlocked and hidden cell. *)
    worksheet_write_string(worksheet, 0, 0, 'B1 is locked. It cannot be edited.',       nil);
    worksheet_write_string(worksheet, 1, 0, 'B2 is unlocked. It can be edited.',        nil);
    worksheet_write_string(worksheet, 2, 0, 'B3 is hidden. The formula isn''t visible.', nil);

    worksheet_write_formula(worksheet, 0, 1, '=1+2', nil);     (* Locked by default. *)
    worksheet_write_formula(worksheet, 1, 1, '=1+2', unlocked);
    worksheet_write_formula(worksheet, 2, 1, '=1+2', hidden);

    workbook_close(workbook);

    ExitCode := 0;
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
