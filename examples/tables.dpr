program tables;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

(* Write some data to the worksheet. *)
procedure write_worksheet_data(worksheet: Plxw_worksheet; format: Plxw_format);
var
  rRow, rCol: DWord;
begin
    decodecell('B4', rCol, rRow);
    worksheet_write_string(worksheet, rCol, rRow, 'Apples',  nil);
    decodecell('B5', rCol, rRow);
    worksheet_write_string(worksheet, rCol, rRow, 'Pears',   nil);
    decodecell('B6', rCol, rRow);
    worksheet_write_string(worksheet, rCol, rRow, 'Bananas', nil);
    decodecell('B7', rCol, rRow);
    worksheet_write_string(worksheet, rCol, rRow, 'Oranges', nil);

    decodecell('C4', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow, 10000,  format);
    decodecell('C5', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  2000,  format);
    decodecell('C6', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  6000,  format);
    decodecell('C7', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,   500,  format);

    decodecell('D4', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  5000,  format);
    decodecell('D5', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  3000,  format);
    decodecell('D6', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  6000,  format);
    decodecell('D7', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,   300,  format);

    decodecell('E4', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  8000,  format);
    decodecell('E5', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  4000,  format);
    decodecell('E6', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  6500,  format);
    decodecell('E7', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,   200,  format);

    decodecell('F4', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  6000,  format);
    decodecell('F5', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  5000,  format);
    decodecell('F6', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,  6000,  format);
    decodecell('F7', rCol, rRow);
    worksheet_write_number(worksheet, rCol, rRow,   700,  format);

end;

var
  workbook: Plxw_workbook;
  worksheet1, worksheet2, worksheet3, worksheet4: Plxw_worksheet;
  worksheet5, worksheet6, worksheet7, worksheet8: Plxw_worksheet;
  worksheet9, worksheet10, worksheet11, worksheet12: Plxw_worksheet;
  worksheet13: Plxw_worksheet;
  currency_format: Plxw_format;
  rCol1, rCol2: DWord;
  rRow1, rRow2: DWord;
  options3, options4, options5, options6: lxw_table_options;
  options7, options8, options9: lxw_table_options;
  options10, options11, options12, options13: lxw_table_options;
  col7_1, col7_2, col7_3, col7_4, col7_5: lxw_table_column;
  col8_1, col8_2, col8_3, col8_4, col8_5, col8_6: lxw_table_column;
  col9_1, col9_2, col9_3, col9_4, col9_5, col9_6: lxw_table_column;
  col10_1, col10_2, col10_3, col10_4, col10_5, col10_6: lxw_table_column;
  col11_1, col11_2, col11_3, col11_4, col11_5, col11_6: lxw_table_column;
  col12_1, col12_2, col12_3, col12_4, col12_5, col12_6: lxw_table_column;
  col13_1, col13_2, col13_3, col13_4, col13_5, col13_6: lxw_table_column;
  columns7: array [0..5] of Plxw_table_column;
  columns8: array [0..7] of Plxw_table_column;
  columns9: array [0..7] of Plxw_table_column;
  columns10: array [0..7] of Plxw_table_column;
  columns11: array [0..7] of Plxw_table_column;
  columns12: array [0..7] of Plxw_table_column;
  columns13: array [0..7] of Plxw_table_column;
  ReportName: PAnsiChar;

begin
  { TODO -oUser -cConsole Main : Insert code here }
  ReportName := 'tables.xlsx';
    workbook    := workbook_new(ReportName);
    worksheet1  := workbook_add_worksheet(workbook, nil);
    worksheet2  := workbook_add_worksheet(workbook, nil);
    worksheet3  := workbook_add_worksheet(workbook, nil);
    worksheet4  := workbook_add_worksheet(workbook, nil);
    worksheet5  := workbook_add_worksheet(workbook, nil);
    worksheet6  := workbook_add_worksheet(workbook, nil);
    worksheet7  := workbook_add_worksheet(workbook, nil);
    worksheet8  := workbook_add_worksheet(workbook, nil);
    worksheet9  := workbook_add_worksheet(workbook, nil);
    worksheet10 := workbook_add_worksheet(workbook, nil);
    worksheet11 := workbook_add_worksheet(workbook, nil);
    worksheet12 := workbook_add_worksheet(workbook, nil);
    worksheet13 := workbook_add_worksheet(workbook, nil);

    currency_format := workbook_add_format(workbook);
    format_set_num_format(currency_format, '$#,##0');

    (*
     * Example 1. Default table with no data
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet1, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecols('B:1', rCol1, rCol2);
    worksheet_write_string(worksheet1, rCol1, rCol2, 'Default table with no data.', nil);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2,rCol2);
    worksheet_add_table(worksheet1, rRow1, rCol1, rRow2,rCol2, nil);

    (*
     * Example 2. Default table with data
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet2, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet2, rRow1, rCol1, 'Default table with data.', nil);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet2, rRow1, rCol1, rRow2, rCol2, nil);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet2, nil);

    (*
     * Example 3. Table without default autofilter
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet3, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet3, rRow1, rCol1, 'Table without default autofilter.', nil);

    (* Set the table options. *)
    options3.no_autofilter := Byte(LXW_TRUE);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet3, rRow1, rCol1, rRow2, rCol2, @options3);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet3, nil);


    (*
     * Example 4. Table without default header row
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet4, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet4, rRow1, rCol1, 'Table without default header row.', nil);

    (* Set the table options. *)
    options4.no_header_row := Byte(LXW_TRUE);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet4, rRow1, rCol1, rRow2, rCol2, @options4);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet4, nil);


    (*
     * Example 5. Default table with "First Column" and "Last Column" options
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet5, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet5, rRow1, rCol1,
                           'Default table with "First Column" and "Last Column" options.',
                           nil);

    (* Set the table options. *)
    options5.first_column := Byte(LXW_TRUE);
    options5.last_column := Byte(LXW_TRUE);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet5, rRow1, rCol1, rRow2, rCol2, @options5);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet5, nil);


    (*
     * Example 6. Table with banded columns but without default banded rows
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet6, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet6, rRow1, rCol1,
                           'Table with banded columns but without default banded rows.',
                           nil);

    (* Set the table options. *)
    options6.no_banded_rows := Byte(LXW_TRUE);
    options6.banded_columns := Byte(LXW_TRUE);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet6, rRow1, rCol1, rRow2, rCol2, @options6);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet6, nil);


    (*
     * Example 7. Table with user defined column headers
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rcol1, rcol2);
    worksheet_set_column(worksheet7, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet7, rRow1, rCol1, 'Table with user defined column headers.', nil);


    (* Set the table options. *)
    col7_1.header := 'Product';

    col7_2.header := 'Quarter 1';
    col7_3.header := 'Quarter 2';
    col7_4.header := 'Quarter 3';
    col7_5.header := 'Quarter 4';

    columns7[0] := @col7_1;
    columns7[1] := @col7_2;
    columns7[2] := @col7_3;
    columns7[3] := @col7_4;
    columns7[4] := @col7_5;
    columns7[5] := nil;


    options7.columns := (@columns7);

    (* Add a table to the worksheet. *)
    decoderange('B3:F7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet7, rRow1, rCol1, rRow2, rCol2, @options7);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet7, nil);


    (*
     * Example 8. Table with user defined column headers
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet8, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet8, rRow1, rCol1, 'Table with user defined column headers.', nil);

    (* Set the table options. *)
    col8_1.header := 'Product';
    col8_2.header := 'Quarter 1';
    col8_3.header := 'Quarter 2';
    col8_4.header := 'Quarter 3';
    col8_5.header := 'Quarter 4';
    col8_6.header := 'Year';
    col8_6.formula := '=SUM(Table8[@[Quarter 1]:[Quarter 4]])';

    columns8[0] := @col8_1;
    columns8[1] := @col8_2;
    columns8[2] := @col8_3;
    columns8[3] := @col8_4;
    columns8[4] := @col8_5;
    columns8[5] := @col8_6;
    columns8[6] := nil;

    options8.columns := @columns8;

    (* Add a table to the worksheet. *)
    decoderange('B3:G7', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet8, rRow1, rCol1, rRow2, rCol2, @options8);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet8, nil);


    (*
     * Example 9. Table with totals row (but no caption or totals)
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet9, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet9, rRow1, rCol1,
                           'Table with totals row (but no caption or totals).',
                           nil);


    (* Set the table options. *)
    col9_1.header := 'Product';
    col9_2.header := 'Quarter 1';
    col9_3.header := 'Quarter 2';
    col9_4.header := 'Quarter 3';
    col9_5.header := 'Quarter 4';
    col9_6.header := 'Year';
    col9_6.formula := '=SUM(Table9[@[Quarter 1]:[Quarter 4]])';

    columns9[0] := @col9_1;
    columns9[1] := @col9_2;
    columns9[2] := @col9_3;
    columns9[3] := @col9_4;
    columns9[4] := @col9_5;
    columns9[5] := @col9_6;
    columns9[6] := nil;

  options9.total_row := Byte(LXW_TRUE);
  options9.columns := @columns9;

    (* Add a table to the worksheet. *)
    decoderange('B3:G8', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet9, rRow1, rCol1, rRow2, rCol2, @options9);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet9, nil);


    (*
     * Example 10. Table with totals row with user captions and functions
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet10, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet10, rRow1, rCol1,
                           'Table with totals row with user captions and functions.',
                           nil);

    (* Set the table options. *)
    col10_1.header         := 'Product';
    col10_1.total_string   := 'Totals';

  col10_2.header         := 'Quarter 1';
  col10_2.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col10_3.header         := 'Quarter 2';
  col10_3.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col10_4.header         := 'Quarter 3';
  col10_4.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col10_5.header         := 'Quarter 4';
  col10_5.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col10_6.header         := 'Year';
  col10_6.formula        := '=SUM(Table10[@[Quarter 1]:[Quarter 4]])';
  col10_6.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

    columns10[0] := @col10_1;
    columns10[1] := @col10_2;
    columns10[2] := @col10_3;
    columns10[3] := @col10_4;
    columns10[4] := @col10_5;
    columns10[5] := @col10_6;
    columns10[6] := nil;

    options10.total_row := Byte(LXW_TRUE);
    options10.columns := @columns10;

    (* Add a table to the worksheet. *)
    decoderange('B3:G8', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet10, rRow1, rCol1, rRow2, rCol2, @options10);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet10, nil);


    (*
     * Example 11. Table with alternative Excel style
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet11, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet11, rRow1, rCol1, 'Table with alternative Excel style.', nil);

    (* Set the table options. *)
  col11_1.header         := 'Product';
  col11_1.total_string   := 'Totals';

  col11_2.header         := 'Quarter 1';
  col11_2.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col11_3.header         := 'Quarter 2';
  col11_3.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col11_4.header         := 'Quarter 3';
  col11_4.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col11_5.header         := 'Quarter 4';
  col11_5.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col11_6.header         := 'Year';
  col11_6.formula        := '=SUM(Table11[@[Quarter 1]:[Quarter 4]])';
  col11_6.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

    columns11[0] := @col11_1;
    columns11[1] := @col11_2;
    columns11[2] := @col11_3;
    columns11[3] := @col11_4;
    columns11[4] := @col11_5;
    columns11[5] := @col11_6;
    columns11[6] := nil;

   options11.style_type := Byte(LXW_TABLE_STYLE_TYPE_LIGHT);
   options11.style_type_number := 11;
   options11.total_row := Byte(LXW_TRUE);
   options11.columns := @columns11;

    (* Add a table to the worksheet. *)
    decoderange('B3:G8', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet11, rRow1, rCol1, rRow2, rCol2, @options11);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet11, nil);

    (*
     * Example 12. Table with Excel style removed
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet12, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet12, rRow1, rCol1, 'Table with Excel style removed.', nil);

    (* Set the table options. *)
  col12_1.header         := 'Product';
  col12_1.total_string   := 'Totals';

  col12_2.header         := 'Quarter 1';
  col12_2.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col12_3.header         := 'Quarter 2';
  col12_3.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col12_4.header         := 'Quarter 3';
  col12_4.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col12_5.header         := 'Quarter 4';
  col12_5.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

  col12_6.header         := 'Year';
  col12_6.formula        := '=SUM(Table12[@[Quarter 1]:[Quarter 4]])';
  col12_6.total_function := Byte(LXW_TABLE_FUNCTION_SUM);

    columns12[0] := @col12_1;
    columns12[1] := @col12_2;
    columns12[2] := @col12_3;
    columns12[3] := @col12_4;
    columns12[4] := @col12_5;
    columns12[5] := @col12_6;
    columns12[6] := nil;

    options12.style_type := Byte(LXW_TABLE_STYLE_TYPE_LIGHT);
    options12.style_type_number := $0;
    options12.total_row := Byte(LXW_TRUE);

    options12.columns := @columns12;

    (* Add a table to the worksheet. *)
    decoderange('B3:G8', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet12, rRow1, rCol1, rRow2, rCol2, @options12);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet12, nil);

    (*
     * Example 13. Table with column formats
     *)

    (* Set the columns widths for clarity. *)
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column(worksheet13, rCol1, rCol2, 12, nil);

    (* Write the worksheet caption to explain the example. *)
    decodecell('B1', rRow1, rCol1);
    worksheet_write_string(worksheet13, rRow1, rCol1, 'Table with column formats.', nil);

    (* Set the table options. *)
    col13_1.header         := 'Product';
    col13_1.total_string   := 'Totals';

    col13_2.header         := 'Quarter 1';
    col13_2.total_function := Byte(LXW_TABLE_FUNCTION_SUM);
    col13_2.format         := currency_format;

    col13_3.header         := 'Quarter 2';
    col13_3.total_function := Byte(LXW_TABLE_FUNCTION_SUM);
    col13_3.format         := currency_format;

    col13_4.header         := 'Quarter 3';
    col13_4.total_function := Byte(LXW_TABLE_FUNCTION_SUM);
    col13_4.format         := currency_format;

    col13_5.header         := 'Quarter 4';
    col13_5.total_function := Byte(LXW_TABLE_FUNCTION_SUM);
    col13_5.format         := currency_format;

    col13_6.header         := 'Year';
    col13_6.formula        := '=SUM(Table13[@[Quarter 1]:[Quarter 4]])';
    col13_6.total_function := Byte(LXW_TABLE_FUNCTION_SUM);
    col13_6.format         := currency_format;

    columns13[0] := @col13_1;
    columns13[1] := @col13_2;
    columns13[2] := @col13_3;
    columns13[3] := @col13_4;
    columns13[4] := @col13_5;
    columns13[5] := @col13_6;
    columns13[6] := nil;


    options13.total_row := Byte(LXW_TRUE);
    options13.columns := @columns13;

    (* Add a table to the worksheet. *)
    decoderange('B3:G8', rRow1, rCol1, rRow2, rCol2);
    worksheet_add_table(worksheet13, rRow1, rCol1, rRow2, rCol2, @options13);

    (* Write the data into the worksheet cells. *)
    write_worksheet_data(worksheet13, currency_format);


    workbook_close(workbook);
end.
