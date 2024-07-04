program autofilter;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

  type
    row = Record
      region: pUTF8Char;
      item: pUTF8Char;
      volume: Integer;
      month: pUTF8Char;
    end;
  const
    data: array [0..49] of row =
       ((region: 'East';  item:   'Apple';   volume: 9000; month: 'July'      ),
        (region: 'East';  item: 'Apple';   volume:   5000; month:  'July'      ),
        (region: 'South'; item:  'Orange';   volume:  9000; month:  'September' ),
        (region: 'North'; item:  'Apple';   volume:   2000; month:  'November'  ),
        (region: 'West';  item:   'Apple';   volume:   9000; month:  'November'  ),
        (region: 'South'; item:  'Pear';   volume:    7000; month:  'October'   ),
        (region: 'North'; item:  'Pear';   volume:    9000; month:  'August'    ),
        (region: 'West';  item:  'Orange';   volume:  1000; month: 'December'  ),
        (region: 'West';  item:  'Grape';   volume:   1000; month: 'November'  ),
        (region: 'South'; item: 'Pear';   volume:   10000; month: 'April'     ),
        (region: 'West';  item:  'Grape';   volume:   6000; month: 'January'   ),
        (region: 'South'; item: 'Orange';   volume:  3000; month: 'May'       ),
        (region: 'North'; item: 'Apple';   volume:   3000; month: 'December'  ),
        (region: 'South'; item: 'Apple';   volume:   7000; month: 'February'  ),
        (region: 'West';  item:  'Grape';   volume:   1000; month: 'December'  ),
        (region: 'East';  item:  'Grape';   volume:   8000; month: 'February'  ),
        (region: 'South'; item: 'Grape';   volume:  10000; month: 'June'      ),
        (region: 'West';  item:  'Pear';   volume:    7000; month: 'December'  ),
        (region: 'South'; item: 'Apple';   volume:   2000; month: 'October'   ),
        (region: 'East';  item:  'Grape';   volume:   7000; month: 'December'  ),
        (region: 'North'; item: 'Grape';   volume:   6000; month: 'April'     ),
        (region: 'East';  item:  'Pear';   volume:    8000; month: 'February'  ),
        (region: 'North'; item: 'Apple';   volume:   7000; month: 'August'    ),
        (region: 'North'; item: 'Orange';   volume:  7000; month: 'July'      ),
        (region: 'North'; item: 'Apple';   volume:   6000; month: 'June'      ),
        (region: 'South'; item: 'Grape';   volume:   8000; month: 'September' ),
        (region: 'West';  item:  'Apple';   volume:   3000; month: 'October'   ),
        (region: 'South'; item: 'Orange';   volume: 10000; month: 'November'  ),
        (region: 'West';  item:  'Grape';   volume:   4000; month: 'July'      ),
        (region: 'North'; item: 'Orange';   volume:  5000; month: 'August'    ),
        (region: 'East';  item:  'Orange';   volume:  1000; month: 'November'  ),
        (region: 'East';  item:  'Orange';   volume:  4000; month: 'October'   ),
        (region: 'North'; item: 'Grape';   volume:   5000; month: 'August'    ),
        (region: 'East';  item:  'Apple';   volume:   1000; month: 'December'  ),
        (region: 'South'; item: 'Apple';   volume:   10000; month: 'March'    ),
        (region: 'East';  item:  'Grape';   volume:   7000; month: 'October'   ),
        (region: 'West';  item:  'Grape';   volume:   1000; month: 'September' ),
        (region: 'East';  item:  'Grape';   volume:  10000; month: 'October'   ),
        (region: 'South'; item: 'Orange';   volume:  8000; month: 'March'     ),
        (region: 'North'; item: 'Apple';   volume:   4000; month: 'July'      ),
        (region: 'South'; item: 'Orange';   volume:  5000; month: 'July'      ),
        (region: 'West';  item:  'Apple';   volume:   4000; month: 'June'      ),
        (region: 'East';  item:  'Apple';   volume:   5000; month: 'April'     ),
        (region: 'North'; item: 'Pear';   volume:    3000; month: 'August'    ),
        (region: 'East';  item:  'Grape';   volume:   9000; month: 'November'  ),
        (region: 'North'; item: 'Orange';   volume:  8000; month: 'October'   ),
        (region: 'East';  item:  'Apple';   volume:  10000; month: 'June'      ),
        (region: 'South'; item: 'Pear';   volume:    1000; month: 'December'  ),
        (region: 'North'; item: 'Grape';   volume:   10000; month: 'July'     ),
        (region: 'East';  item:  'Grape';   volume:   6000; month: 'February'  )
        );

  list: array [0..3] of PUtf8Char = ('East', 'North', 'South', nil);

procedure write_worksheet_header(worksheet: Plxw_worksheet; header: Plxw_format);
begin
    (* Make the columns wider for clarity. *)
    worksheet_set_column(worksheet, 0, 3, 12, nil);

    (* Write the column headers. *)
    worksheet_set_row(worksheet, 0, 20, header);
    worksheet_write_string(worksheet, 0, 0, 'Region', nil);
    worksheet_write_string(worksheet, 0, 1, 'Item',   nil);
    worksheet_write_string(worksheet, 0, 2, 'Volume', nil);
    worksheet_write_string(worksheet, 0, 3, 'Month',  nil);
end;

var
  workbook: Plxw_workbook;
  worksheet1: Plxw_worksheet;
  worksheet2: Plxw_worksheet;
  worksheet3: Plxw_worksheet;
  worksheet4: Plxw_worksheet;
  worksheet5: Plxw_worksheet;
  worksheet6: Plxw_worksheet;
  worksheet7: Plxw_worksheet;
  i: Word;
  hidden: lxw_row_col_options;
  filter_rule2: lxw_filter_rule;
  filter_rule3a, filter_rule3b: lxw_filter_rule;
  filter_rule4a, filter_rule4b, filter_rule4c: lxw_filter_rule;
  filter_rule6: lxw_filter_rule;
  filter_rule7: lxw_filter_rule;
  header: Plxw_format;
  ReportName: PAnsiChar;
begin
  ReportName := 'autofilter.xlsx';

  workbook   := workbook_new(ReportName);
  worksheet1 := workbook_add_worksheet(workbook, nil);
  worksheet2 := workbook_add_worksheet(workbook, nil);
  worksheet3 := workbook_add_worksheet(workbook, nil);
  worksheet4 := workbook_add_worksheet(workbook, nil);
  worksheet5 := workbook_add_worksheet(workbook, nil);
  worksheet6 := workbook_add_worksheet(workbook, nil);
  worksheet7 := workbook_add_worksheet(workbook, nil);


    hidden.hidden := Byte(LXW_TRUE);

    header := workbook_add_format(workbook);
    format_set_bold(header);



    (*
     * Example 1. Autofilter without conditions.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet1, header);

    (* Write the row data. *)
    for i := 0 to Length(data)-1 do
    begin
        worksheet_write_string(worksheet1, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet1, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet1, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet1, i + 1, 3, data[i].month,  nil);
    end;


    (* Add the autofilter. *)
    worksheet_autofilter(worksheet1, 0, 0, 50, 3);


    (*
     * Example 2. Autofilter with a filter condition in the first column.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet2, header);

    (* Write the row data. *)
    for i := 0 to Length(data)-1 do
    begin
        worksheet_write_string(worksheet2, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet2, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet2, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet2, i + 1, 3, data[i].month,  nil);

        (* It isn't sufficient to just apply the filter condition below. We
         * must also hide the rows that don't match the criteria since Excel
         * doesn't do that automatically. *)
        if (data[i].region = 'East') then
            (* Row matches the filter, no further action required. *)

        else
            (* Hide rows that don't match the filter. *)
            worksheet_set_row_opt(worksheet2, i + 1, LXW_DEF_ROW_HEIGHT, nil, @hidden);


        (* Note, the if() statement above is written to match the logic of the
         * criteria in worksheet_filter_column() below. However you could get
         * the same results with the following simpler, but reversed, code:
         *
         *     if (strcmp(data[i].region, 'East') != 0) {
         *         worksheet_set_row_opt(worksheet2, i + 1, LXW_DEF_ROW_HEIGHT, nil, &hidden);
         *     }
         *
         * The same applies to the Examples 3-6 as well.
         *)
    end;


    (* Add the autofilter. *)
    worksheet_autofilter(worksheet2, 0, 0, 50, 3);

    (* Add the filter criteria. *)
    filter_rule2.criteria     := Byte(LXW_FILTER_CRITERIA_EQUAL_TO);
    filter_rule2.value_string := 'East';

    worksheet_filter_column(worksheet2, 0, @filter_rule2);


    (*
     * Example 3. Autofilter with a dual filter condition in one of the columns.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet3, header);

    (* Write the row data. *)
    for i := 0 to Length(data)-1 do
    begin
        worksheet_write_string(worksheet3, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet3, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet3, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet3, i + 1, 3, data[i].month,  nil);

        if (data[i].region = 'East') or (data[i].region = 'South') then
            (* Row matches the filter, no further action required. *)
        else
            (* We need to hide rows that don't match the filter. *)
            worksheet_set_row_opt(worksheet3, i + 1, LXW_DEF_ROW_HEIGHT, nil, @hidden);

    end;

    (* Add the autofilter. *)
    worksheet_autofilter(worksheet3, 0, 0, 50, 3);

    (* Add the filter criteria. *)
    filter_rule3a.criteria     := Byte(LXW_FILTER_CRITERIA_EQUAL_TO);
    filter_rule3a.value_string := 'East';

    filter_rule3b.criteria     := Byte(LXW_FILTER_CRITERIA_EQUAL_TO);
    filter_rule3b.value_string := 'South';

    worksheet_filter_column2(worksheet3, 0, @filter_rule3a, @filter_rule3b, Byte(LXW_FILTER_OR));



    (*
     * Example 4. Autofilter with filter conditions in two columns.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet4, header);

    (* Write the row data. *)
    for i := 0 to Length(data)-1 do
    begin
        worksheet_write_string(worksheet4, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet4, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet4, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet4, i + 1, 3, data[i].month,  nil);

        if (data[i].region = 'East') and
            (data[i].volume > 3000) and (data[i].volume < 8000) then

            (* Row matches the filter, no further action required. *)

        else
            (* We need to hide rows that don't match the filter. *)
            worksheet_set_row_opt(worksheet4, i + 1, LXW_DEF_ROW_HEIGHT, nil, @hidden);

    end;

    (* Add the autofilter. *)
    worksheet_autofilter(worksheet4, 0, 0, 50, 3);

    (* Add the filter criteria. *)
    filter_rule4a.criteria := Byte(LXW_FILTER_CRITERIA_EQUAL_TO);
    filter_rule4a.value_string := 'East';

    filter_rule4b.criteria     := Byte(LXW_FILTER_CRITERIA_GREATER_THAN);
    filter_rule4b.value        := 3000;

    filter_rule4c.criteria     := Byte(LXW_FILTER_CRITERIA_LESS_THAN);
    filter_rule4c.value        := 8000;

    worksheet_filter_column(worksheet4,  0, @filter_rule4a);
    worksheet_filter_column2(worksheet4, 2, @filter_rule4b, @filter_rule4c, Byte(LXW_FILTER_AND));


    (*
     * Example 5. Autofilter with a dual filter condition in one of the columns.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet5, header);

    (* Write the row data. *)
    for i := 0 to Length(data)-1 do begin

        worksheet_write_string(worksheet5, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet5, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet5, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet5, i + 1, 3, data[i].month,  nil);

        if (data[i].region = 'East') or
            (data[i].region = 'North') or
            (data[i].region = 'South') then
            (* Row matches the filter, no further action required. *)
        else
            (* We need to hide rows that don't match the filter. *)
            worksheet_set_row_opt(worksheet5, i + 1, LXW_DEF_ROW_HEIGHT, nil, @hidden);
    end;

    (* Add the autofilter. *)
    worksheet_autofilter(worksheet5, 0, 0, 50, 3);

    (* Add the filter criteria. *)
//    const char* list[] = {'East', 'North', 'South', nil};

    worksheet_filter_list(worksheet5, 0, @list);


    (*
     * Example 6. Autofilter with filter for blanks.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet6, header);

// Delphi does not allow altering constants
    (* Simulate one blank cell in the data, to test the filter. *)
//    data[5].region[0] := Char($0);


    (* Write the row data. *)
    for i := 0 to Length(data)-1 do begin
        worksheet_write_string(worksheet6, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet6, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet6, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet6, i + 1, 3, data[i].month,  nil);

        if data[i].region =  '' then
            (* Row matches the filter, no further action required. *)
        else
            (* We need to hide rows that don't match the filter. *)
            worksheet_set_row_opt(worksheet6, i + 1, LXW_DEF_ROW_HEIGHT, nil, @hidden);
    end;

    (* Add the autofilter. *)
    worksheet_autofilter(worksheet6, 0, 0, 50, 3);

    (* Add the filter criteria. *)
    filter_rule6.criteria  := Byte(LXW_FILTER_CRITERIA_BLANKS);

    worksheet_filter_column(worksheet6, 0, @filter_rule6);


    (*
     * Example 7. Autofilter with filter for non-blanks.
     *)

    (* Set up the worksheet data. *)
    write_worksheet_header(worksheet7, header);

    (* Write the row data. *)
    for i := 0 to Length(data)-1 do
    begin
        worksheet_write_string(worksheet7, i + 1, 0, data[i].region, nil);
        worksheet_write_string(worksheet7, i + 1, 1, data[i].item,   nil);
        worksheet_write_number(worksheet7, i + 1, 2, data[i].volume, nil);
        worksheet_write_string(worksheet7, i + 1, 3, data[i].month,  nil);

        if (data[i].region <> '') then
            (* Row matches the filter, no further action required. *)
        else
            (* We need to hide rows that don't match the filter. *)
            worksheet_set_row_opt(worksheet7, i + 1, LXW_DEF_ROW_HEIGHT, nil, @hidden);
    end;

    (* Add the autofilter. *)
    worksheet_autofilter(worksheet7, 0, 0, 50, 3);

    (* Add the filter criteria. *)
    filter_rule7.criteria  := Byte(LXW_FILTER_CRITERIA_NON_BLANKS);

    worksheet_filter_column(worksheet7, 0, @filter_rule7);

    workbook_close(workbook);

end.
