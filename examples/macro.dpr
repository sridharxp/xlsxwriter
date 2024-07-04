program macro;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  rCol1, rCol2: DWord;
  options: lxw_button_options;
  ReportName: PAnsiChar;
begin
    (* Note the xlsm extension of the filename *)
  ReportName := 'macro.xlsm';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

    decodecols('A:A', rCol1, rCol2);
    worksheet_set_column(worksheet, rCol1, rCol2, 30, nil);

    (* Add a macro file extracted from an Excel workbook. *)
    workbook_add_vba_project(workbook, '.\examples\vbaProject.bin');

    worksheet_write_string(worksheet, 2, 0, 'Press the button to say hello.', nil);

    options.caption := 'Press Me';
    options.macro := 'say_hello';
    options.width := 80;
    options.height := 30;

     worksheet_insert_button(worksheet, 2, 1, @options);


  workbook_close(workbook);

end.
