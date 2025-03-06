program array_formula;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  rRow1, rCol1, rRow2, rCol2: DWord;
  ReportName: PAnsiChar;
begin
    (* Create a new workbook and add a worksheet. *)
  ReportName := '.\Report\array_formula.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);
    (* Write some data for the formulas. *)
    worksheet_write_number(worksheet, 0, 1, 500, nil);
    worksheet_write_number(worksheet, 1, 1, 10, nil);
    worksheet_write_number(worksheet, 4, 1, 1, nil);
    worksheet_write_number(worksheet, 5, 1, 2, nil);
    worksheet_write_number(worksheet, 6, 1, 3, nil);

    worksheet_write_number(worksheet, 0, 2, 300, nil);
    worksheet_write_number(worksheet, 1, 2, 15, nil);
    worksheet_write_number(worksheet, 4, 2, 20234, nil);
    worksheet_write_number(worksheet, 5, 2, 21003, nil);
    worksheet_write_number(worksheet, 6, 2, 10000, nil);

    (* Write an array formula that returns a single value. *)
    worksheet_write_array_formula(worksheet, 0, 0, 0, 0, '{=SUM(B1:C1*B2:C2)}', nil);

    (* Similar to above but using the RANGE macro. *)
    decoderange('A2:A2', rRow1, rCol1, rRow2, rCol2);
    worksheet_write_array_formula(worksheet, rRow1, rCol1, rRow2, rCol2, '{=SUM(B1:C1*B2:C2)}', nil);

    (* Write an array formula that returns a range of values. *)
    worksheet_write_array_formula(worksheet, 4, 0, 6, 0, '{=TREND(C5:C7,B5:B7)}', nil);

  workbook_close(workbook);

  ExitCode := 0;
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
