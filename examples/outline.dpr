program outline;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet1, worksheet2, worksheet3, worksheet4: Plxw_worksheet;
  bold: Plxw_format;
  options1, options2: lxw_row_col_options;
  options3, options4, options5: lxw_row_col_options;
  options6: lxw_row_col_options;
  level1, level2: lxw_row_col_options;
  level3, level4, level5: lxw_row_col_options;
  level6, level7: lxw_row_col_options;
  rCol1, rCol2: DWord;
  rRow, rCol: DWord;
  ReportName: PAnsiChar;
begin
  ReportName := '.\Report\outline.xlsx';
  workbook  := workbook_new(ReportName);
    worksheet1 := workbook_add_worksheet(workbook, 'Outlined Rows');
    worksheet2 := workbook_add_worksheet(workbook, 'Collapsed Rows');
    worksheet3 := workbook_add_worksheet(workbook, 'Outline Columns');
    worksheet4 := workbook_add_worksheet(workbook, 'Outline levels');

    bold := workbook_add_format(workbook);
    format_set_bold(bold);

   (*
    * Example 1: Create a worksheet with outlined rows. It also includes
    * SUBTOTAL() functions so that it looks like the type of automatic
    * outlines that are generated when you use the 'Sub Totals' option.
    *
    * For outlines the important parameters are 'hidden' and 'level'. Rows
    * with the same 'level' are grouped together. The group will be collapsed
    * if 'hidden' is non-zero.
    *)

    (* The option structs with the outline level set. *)
    options1.hidden := 0;
    options1.level := 2;
    options1.collapsed := 0;
    options2.hidden := 0;
    options2.level := 1;
    options2.collapsed := 0;

    (* Set the column width for clarity. *)
    decodecols('A:A', rCol1, rCol2);
    worksheet_set_column(worksheet1, rCol1, rCol2, 20, nil);

    (* Set the row options with the outline level. *)
    worksheet_set_row_opt(worksheet1, 1,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 2,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 3,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 4,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 5,  LXW_DEF_ROW_HEIGHT, nil, @options2);

    worksheet_set_row_opt(worksheet1, 6,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 7,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 8,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 9,  LXW_DEF_ROW_HEIGHT, nil, @options1);
    worksheet_set_row_opt(worksheet1, 10, LXW_DEF_ROW_HEIGHT, nil, @options2);

    (* Add data and formulas to the worksheet. *)
    decodecell('A1', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'Region', bold);
    decodecell('A2', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'North',  nil);
    decodecell('A3', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'North',  nil);
    decodecell('A4', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'North',  nil);
    decodecell('A5', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'North',  nil);
    decodecell('A6', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'North Total', bold);

    decodecell('B1', rRow, rCol);
    worksheet_write_string(worksheet1,rRow, rCol, 'Sales', bold);
    decodecell('B2', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol, 1000,    nil);
    decodecell('B3', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol, 1200,    nil);
    decodecell('4', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol, 900,     nil);
    decodecell('B5', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol, 1200,    nil);
    decodecell('B6', rRow, rCol);
    worksheet_write_formula(worksheet1, rRow, rCol, '=SUBTOTAL(9,B2:B5)', bold);

    decodecell('A7', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'South',  nil);
    decodecell('A8', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'South',  nil);
    decodecell('A9', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'South',  nil);
    decodecell('A10', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'South', nil);
    decodecell('A11', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'South Total', bold);

    decodecell('B7', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol,  400, nil);
    decodecell('B8', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol,  600, nil);
    decodecell('B9', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol,  500, nil);
    decodecell('B10', rRow, rCol);
    worksheet_write_number(worksheet1, rRow, rCol, 600, nil);
    decodecell('B11', rRow, rCol);
    worksheet_write_formula(worksheet1, rRow, rCol, '=SUBTOTAL(9,B7:B10)', bold);

    decodecell('A12', rRow, rCol);
    worksheet_write_string(worksheet1, rRow, rCol, 'Grand Total', bold);
    decodecell('B12', rRow, rCol);
    worksheet_write_formula(worksheet1, rRow, rCol, '=SUBTOTAL(9,B2:B10)', bold);


   (*
    * Example 2: Create a worksheet with outlined rows. This is the same as
    * the previous example except that the rows are collapsed.  Note: We need
    * to indicate the row that contains the collapsed symbol '+' with the
    * optional parameter, 'collapsed'.
    *)

    (* The option structs with the outline level and collapsed property set. *)
    options3.hidden := 1;
    options3.level := 2;
    options3.collapsed := 0;
    options4.hidden := 1;
    options4.level := 1;
    options4.collapsed := 0;
    options5.hidden := 0;
    options5.level := 0;
    options5.collapsed := 1;

    (* Set the column width for clarity. *)
    decodecols('A:A', rCol1, rCol2);
    worksheet_set_column(worksheet2, rCol1, rCol2, 20, nil);

    (* Set the row options with the outline level. *)
    worksheet_set_row_opt(worksheet2, 1,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 2,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 3,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 4,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 5,  LXW_DEF_ROW_HEIGHT, nil, @options4);

    worksheet_set_row_opt(worksheet2, 6,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 7,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 8,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 9,  LXW_DEF_ROW_HEIGHT, nil, @options3);
    worksheet_set_row_opt(worksheet2, 10, LXW_DEF_ROW_HEIGHT, nil, @options4);
    worksheet_set_row_opt(worksheet2, 11, LXW_DEF_ROW_HEIGHT, nil, @options5);

    (* Add data and formulas to the worksheet. *)
    decodecell('A1', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'Region', bold);
    decodecell('A2', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'North',  nil);
    decodecell('A3', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'North',  nil);
    decodecell('A4', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'North',  nil);
    decodecell('A5', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'North',  nil);
    decodecell('A6', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'North Total', bold);

    decodecell('B1', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'Sales', bold);
    decodecell('B2', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol, 1000, nil);
    decodecell('B3', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol, 1200, nil);
    decodecell('B4', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol, 900,  nil);
    decodecell('B5', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol, 1200, nil);
    decodecell('B6', rRow, rCol);
    worksheet_write_formula(worksheet2, rRow, rCol, '=SUBTOTAL(9,B2:B5)', bold);

    decodecell('A7', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol,  'South', nil);
    decodecell('A8', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol,  'South', nil);
    decodecell('A9', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol,  'South', nil);
    decodecell('A10', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'South', nil);
    decodecell('A11', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'South Total', bold);

    decodecell('B7', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol,  400, nil);
    decodecell('B8', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol,  600, nil);
    decodecell('B9', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol,  500, nil);
    decodecell('B10', rRow, rCol);
    worksheet_write_number(worksheet2, rRow, rCol, 600, nil);
    decodecell('B11', rRow, rCol);
    worksheet_write_formula(worksheet2, rRow, rCol, '=SUBTOTAL(9,B7:B10)', bold);

    decodecell('A12', rRow, rCol);
    worksheet_write_string(worksheet2, rRow, rCol, 'Grand Total', bold);
    decodecell('B12', rRow, rCol);
    worksheet_write_formula(worksheet2, rRow, rCol, '=SUBTOTAL(9,B2:B10)', bold);


    (*
     * Example 3: Create a worksheet with outlined columns.
     *)
    options6.hidden := 0;
    options6.level := 1;
    options6.collapsed := 0;

    (* Add data and formulas to the worksheet. *)
    decodecell('A1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Month', nil);
    decodecell('B1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Jan',   nil);
    decodecell('C1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Feb',   nil);
    decodecell('D1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Mar',   nil);
    decodecell('E1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Apr',   nil);
    decodecell('F1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'May',   nil);
    decodecell('G1', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Jun',   nil);
    decodecell('', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'Total', nil);

    decodecell('A2', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'North', nil);
    decodecell('B2', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 50,      nil);
    decodecell('C2', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 20,      nil);
    decodecell('D2', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 15,      nil);
    decodecell('E2', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 25,      nil);
    decodecell('F2', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 65,      nil);
    decodecell('G2', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 80,      nil);
    decodecell('H2', rRow, rCol);
    worksheet_write_formula(worksheet3, rRow, rCol, '=SUM(B2:G2)', nil);

    decodecell('A3', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'South', nil);
    decodecell('B3', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 10,      nil);
    decodecell('C3', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 20,      nil);
    decodecell('D3', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 30,      nil);
    decodecell('E3', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 50,      nil);
    decodecell('F3', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 50,      nil);
    decodecell('G3', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 50,      nil);
    decodecell('H3', rRow, rCol);
    worksheet_write_formula(worksheet3, rRow, rCol, '=SUM(B3:G3)', nil);

    decodecell('A4', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'East',  nil);
    decodecell('B4', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 45,      nil);
    decodecell('C4', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 75,      nil);
    decodecell('D4', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 50,      nil);
    decodecell('E4', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 15,      nil);
    decodecell('F4', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 75,      nil);
    decodecell('G4', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 100,     nil);
    decodecell('H4', rRow, rCol);
    worksheet_write_formula(worksheet3, rRow, rCol, '=SUM(B4:G4)', nil);

    decodecell('A5', rRow, rCol);
    worksheet_write_string(worksheet3, rRow, rCol, 'West',  nil);
    decodecell('B5', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 15,      nil);
    decodecell('C5', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 15,      nil);
    decodecell('D5', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 55,      nil);
    decodecell('E5', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 35,      nil);
    decodecell('F5', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 20,      nil);
    decodecell('G5', rRow, rCol);
    worksheet_write_number(worksheet3, rRow, rCol, 50,      nil);
    decodecell('H5', rRow, rCol);
    worksheet_write_formula(worksheet3, rRow, rCol, '=SUM(B5:G5)', nil);

    decodecell('H6', rRow, rCol);
    worksheet_write_formula(worksheet3, rRow, rCol, '=SUM(H2:H5)', bold);

    (* Add bold format to the first row. *)
    worksheet_set_row(worksheet3, 0, LXW_DEF_ROW_HEIGHT, bold);

    (* Set column formatting and the outline level. *)
    decodecols('A:A', rCol1, rCol2);
    worksheet_set_column(    worksheet3, rCol1, rCol2, 10, bold);
    decodecols('B:G', rCol1, rCol2);
    worksheet_set_column_opt(worksheet3, rCol1, rCol2,  5, nil, @options6);
    decodecols('H:H', rCol1, rCol2);
    worksheet_set_column(    worksheet3, rCol1, rCol2, 10, nil);



    (*
     * Example 4: Show all possible outline levels.
     *)
    level1.level := 1;
    level1.hidden := 0;
    level1.collapsed := 0;

    level2.level := 2;
    level2.hidden := 0;
    level2.collapsed := 0;

    level3.level := 3;
    level3.hidden := 0;
    level3.collapsed := 0;

    level4.level := 4;
    level4.hidden := 0;
    level4.collapsed := 0;

    level5.level := 5;
    level5.hidden := 0;
    level5.collapsed := 0;

    level6.level := 6;
    level6.hidden := 0;
    level6.collapsed := 0;

    level7.level := 7;
    level7.hidden := 0;
    level7.collapsed := 0;

    worksheet_write_string(worksheet4, 0,  0, 'Level 1', nil);
    worksheet_write_string(worksheet4, 1,  0, 'Level 2', nil);
    worksheet_write_string(worksheet4, 2,  0, 'Level 3', nil);
    worksheet_write_string(worksheet4, 3,  0, 'Level 4', nil);
    worksheet_write_string(worksheet4, 4,  0, 'Level 5', nil);
    worksheet_write_string(worksheet4, 5,  0, 'Level 6', nil);
    worksheet_write_string(worksheet4, 6,  0, 'Level 7', nil);
    worksheet_write_string(worksheet4, 7,  0, 'Level 6', nil);
    worksheet_write_string(worksheet4, 8,  0, 'Level 5', nil);
    worksheet_write_string(worksheet4, 9,  0, 'Level 4', nil);
    worksheet_write_string(worksheet4, 10, 0, 'Level 3', nil);
    worksheet_write_string(worksheet4, 11, 0, 'Level 2', nil);
    worksheet_write_string(worksheet4, 12, 0, 'Level 1', nil);

    worksheet_set_row_opt(worksheet4, 0,  LXW_DEF_ROW_HEIGHT, nil, @level1);
    worksheet_set_row_opt(worksheet4, 1,  LXW_DEF_ROW_HEIGHT, nil, @level2);
    worksheet_set_row_opt(worksheet4, 2,  LXW_DEF_ROW_HEIGHT, nil, @level3);
    worksheet_set_row_opt(worksheet4, 3,  LXW_DEF_ROW_HEIGHT, nil, @level4);
    worksheet_set_row_opt(worksheet4, 4,  LXW_DEF_ROW_HEIGHT, nil, @level5);
    worksheet_set_row_opt(worksheet4, 5,  LXW_DEF_ROW_HEIGHT, nil, @level6);
    worksheet_set_row_opt(worksheet4, 6,  LXW_DEF_ROW_HEIGHT, nil, @level7);
    worksheet_set_row_opt(worksheet4, 7,  LXW_DEF_ROW_HEIGHT, nil, @level6);
    worksheet_set_row_opt(worksheet4, 8,  LXW_DEF_ROW_HEIGHT, nil, @level5);
    worksheet_set_row_opt(worksheet4, 9,  LXW_DEF_ROW_HEIGHT, nil, @level4);
    worksheet_set_row_opt(worksheet4, 10, LXW_DEF_ROW_HEIGHT, nil, @level3);
    worksheet_set_row_opt(worksheet4, 11, LXW_DEF_ROW_HEIGHT, nil, @level2);
    worksheet_set_row_opt(worksheet4, 12, LXW_DEF_ROW_HEIGHT, nil, @level1);

    workbook_close(workbook);

    ExitCode := 0;
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
