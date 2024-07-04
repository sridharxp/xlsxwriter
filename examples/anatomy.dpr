program anatomy;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet1, worksheet2: Plxw_worksheet;
  myformat1, myformat2: Plxw_format;
  error: lxw_error;
  ReportName: PAnsiChar;
begin
  ReportName := 'anatomy.xlsx';
    (* Create a new workbook. *)
  workbook  := workbook_new(ReportName);
    (* Add a worksheet with a user defined sheet name. *)
  worksheet1 := workbook_add_worksheet(workbook, nil);
    (* Add a worksheet with Excel's default sheet name: Sheet2. *)
  worksheet2 := workbook_add_worksheet(workbook, nil);

    (* Add some cell formats. *)
    myformat1    := workbook_add_format(workbook);
    myformat2    := workbook_add_format(workbook);

    (* Set the bold property for the first format. *)
    format_set_bold(myformat1);

    (* Set a number format for the second format. *)
    format_set_num_format(myformat2, '$#,##0.00');

    (* Widen the first column to make the text clearer. *)
    worksheet_set_column(worksheet1, 0, 0, 20, nil);

    (* Write some unformatted data. *)
    worksheet_write_string(worksheet1, 0, 0, 'Peach', nil);
    worksheet_write_string(worksheet1, 1, 0, 'Plum',  nil);

    (* Write formatted data. *)
    worksheet_write_string(worksheet1, 2, 0, 'Pear',  myformat1);

    (* Formats can be reused. *)
    worksheet_write_string(worksheet1, 3, 0, 'Persimmon',  myformat1);


    (* Write some numbers. *)
    worksheet_write_number(worksheet1, 5, 0, 123,       nil);
    worksheet_write_number(worksheet1, 6, 0, 4567.555,  myformat2);


    (* Write to the second worksheet. *)
    worksheet_write_string(worksheet2, 0, 0, 'Some text', myformat1);


    (* Close the workbook, save the file and free any memory. *)
   error := workbook_close(workbook);

    (* Check if there was any error creating the xlsx file. *)
    if (Integer(error) > 0) then
{
        printf("Error in workbook_close().\n"
               "Error %d = %s\n", error, lxw_strerror(error));
}
        Writeln(Format('Error in workbook_close().\n'+
               'Error %d = ' + lxw_strerror(error) + '\n', [Integer(error)]));
end.
