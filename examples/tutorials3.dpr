program tutorials3;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '\DL\LXW\xlsxwriterapi.pas';

type
/* Some data we want to write to the worksheet. */
  expense = Record
    item: pAnsiChar;
    cost: Integer;
    datetime: lxw_datetime;
  end;


const
  expenses: array [0..3] of expense =
    (
    (item: 'Rent'; cost: 1000),
    (item: 'Gas';  cost:  100),
    (item: 'Food'; cost:  300),
    (item: 'Gym';  cost:   50)
    );
  datetimes: array [0..3] of lxw_datetime =
    (
    (year: 2013; month: 1; day: 13),
    (year: 2013; month: 1; day: 14),
    (year: 2013; month: 1; day: 15),
    (year: 2013; month: 1; day: 16)
    );

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  bold, money: Plxw_format;
  date_format: Plxw_format;
  row, col, i: integer;
  ReportName: PAnsiChar;
begin
  ReportName := 'tutorial03.xlsx';

  (* Create a workbook and add a worksheet. *)
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);
  row := 0;
  col := 0;

  (* Add a bold format to use to highlight cells. *)
  bold := workbook_add_format(workbook);
  format_set_bold(bold);

  (* Add a number format for cells with money. *)
  money := workbook_add_format(workbook);
  format_set_num_format(money, '$#,##0');

    (* Add an Excel date format. *)
  date_format := workbook_add_format(workbook);
  format_set_num_format(date_format, 'd mmmm yyyy');

    (* Adjust the column width. *)
  worksheet_set_column(worksheet, 0, 0, 15, nil);
  worksheet_set_column(worksheet, 0, 1, 15, nil);

    (* Write some data header. *)
  worksheet_write_string(worksheet, row, col,     'Item', bold);
  worksheet_write_string(worksheet, row, col + 2, 'Cost', bold);

  (* Iterate over the data and write it out element by element. *)
  for i := 0 to 3  do
  begin
        (* Write from the first cell below the headers. *)
    row := i + 1;
    worksheet_write_string(worksheet, row, col,     expenses[i].item, nil);
    worksheet_write_datetime(worksheet, row, col + 1, @datetimes[i], date_format);
    worksheet_write_number(worksheet, row, col + 2, expenses[i].cost, money);
  end;

  (* Write a total using a formula. *)
  worksheet_write_string (worksheet, row + 1, col,     'Total',       bold);
  worksheet_write_formula(worksheet, row + 1, col + 2, '=SUM(B2:B5)', money);

  (* Save the workbook and free any allocated memory. *)
  workbook_close(workbook);
end.
