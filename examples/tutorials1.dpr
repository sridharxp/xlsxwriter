program tutorials1;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';
type
  expense = Record
    item: pAnsiChar;
    cost: Integer;
  end;

const
  expenses: array [0..3] of expense =
    (
    (item: 'Rent'; cost: 1000),
    (item: 'Gas';  cost:  100),
    (item: 'Food'; cost:  300),
    (item: 'Gym';  cost:   50)
    );
var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  row, col: integer;
  ReportName: PAnsiChar;
begin
  ReportName := '.tutorial01.xlsx';

  (* Create a workbook and add a worksheet. *)
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

  (* Start from the first cell. Rows and columns are zero indexed. *)
  row := 0;
  col := 0;

    (* Iterate over the data and write it out element by element. *)
  for row := 0 to 3 do
  begin
    worksheet_write_string(worksheet, row, col,     expenses[row].item, nil);
    worksheet_write_number(worksheet, row, col + 1, expenses[row].cost, nil);
  end;

 (* Write a total using a formula. *)
  worksheet_write_string (worksheet, row, col,     'Total',       nil);
  worksheet_write_formula(worksheet, row, col + 1, '=SUM(B1:B4)', nil);

  (* Save the workbook and free any allocated memory. *)
  lxw_error(ExitCode) := workbook_close(workbook);
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
