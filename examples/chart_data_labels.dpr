program chart_data_labels;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  bold: Plxw_format;
  options: lxw_chart_options;
  chart: Plxw_chart;
  series: Plxw_chart_series;
  rRow, rCol: DWord;
  font1: lxw_chart_font;
  line1: lxw_chart_line;
  fill1: lxw_chart_fill;
  data_label5_1, data_label5_2, data_label5_3, data_label5_4: lxw_chart_data_label;
  data_label5_5, data_label5_6: lxw_chart_data_label;
  data_labels5: array [0..6] of Plxw_chart_data_label;
  data_label6_1, data_label6_2, data_label6_3, data_label6_4: lxw_chart_data_label;
  data_label6_5, data_label6_6: lxw_chart_data_label;
  data_labels6: array [0..6] of Plxw_chart_data_label;
  font2: lxw_chart_font;
  data_label7_1, data_label7_2, data_label7_3: lxw_chart_data_label;
  data_label7_4: lxw_chart_data_label;
  data_labels7: array [0..4] of Plxw_chart_data_label;
  hide, keep: lxw_chart_data_label;
  data_labels8: array [0..6] of Plxw_chart_data_label;
  line2, line3: lxw_chart_line;
  fill2, fill3: lxw_chart_fill;
  data_label9_1, data_label9_2, data_label9_3, data_label9_4: lxw_chart_data_label;
  data_label9_5, data_label9_6: lxw_chart_data_label;
  data_labels9: array [0..6] of Plxw_chart_data_label;
  ReportName: PAnsiChar;
begin
  ReportName := 'chart_data_labels.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);

    (* Add a bold format to use to highlight the header cells. *)
    bold := workbook_add_format(workbook);
    format_set_bold(bold);

    (* Some chart positioning options. *)
    options.x_offset := 25;
    options.y_offset := 10;

    (* Write some data for the chart. *)
    worksheet_write_string(worksheet, 0, 0, 'Number',  bold);
    worksheet_write_number(worksheet, 1, 0, 2,         nil);
    worksheet_write_number(worksheet, 2, 0, 3,         nil);
    worksheet_write_number(worksheet, 3, 0, 4,         nil);
    worksheet_write_number(worksheet, 4, 0, 5,         nil);
    worksheet_write_number(worksheet, 5, 0, 6,         nil);
    worksheet_write_number(worksheet, 6, 0, 7,         nil);

    worksheet_write_string(worksheet, 0, 1, 'Data',    bold);
    worksheet_write_number(worksheet, 1, 1, 20,        nil);
    worksheet_write_number(worksheet, 2, 1, 10,        nil);
    worksheet_write_number(worksheet, 3, 1, 20,        nil);
    worksheet_write_number(worksheet, 4, 1, 30,        nil);
    worksheet_write_number(worksheet, 5, 1, 40,        nil);
    worksheet_write_number(worksheet, 6, 1, 30,        nil);

    worksheet_write_string(worksheet, 0, 2, 'Text',    bold);
    worksheet_write_string(worksheet, 1, 2, 'Jan',     nil);
    worksheet_write_string(worksheet, 2, 2, 'Feb',     nil);
    worksheet_write_string(worksheet, 3, 2, 'Mar',     nil);
    worksheet_write_string(worksheet, 4, 2, 'Apr',     nil);
    worksheet_write_string(worksheet, 5, 2, 'May',     nil);
    worksheet_write_string(worksheet, 6, 2, 'Jun',     nil);


    (*
     * Chart 1. Example with standard data labels.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Chart with standard data labels');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                                       '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D2', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 2. Example with value and category data labels.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Category and Value data labels');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Turn on Value and Category labels. *)
    chart_series_set_labels_options(series, Byte(LXW_FALSE), Byte(LXW_TRUE), Byte(LXW_TRUE));

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D18', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 3. Example with standard data labels with different font.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Data labels with user defined font');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    font1.bold := Byte(LXW_TRUE);
    font1.color := Cardinal(LXW_COLOR_RED);
    font1.rotation := -30;

    chart_series_set_labels_font(series, @font1);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D34', rRow, rCol);
    worksheet_insert_chart_opt(worksheet,  rRow, rCol, chart, @options);


    (*
     * Chart 4. Example with standard data labels and formatting.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Data labels with formatting');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Set the border/line and fill for the data labels. *)
    line1.color := Cardinal(LXW_COLOR_RED);
    fill1.color := Cardinal(LXW_COLOR_YELLOW);

    chart_series_set_labels_line(series, @line1);
    chart_series_set_labels_fill(series, @fill1);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D50', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 5.Example with custom string data labels.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Chart with custom string data labels');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Create some custom labels. *)
    data_label5_1.value := 'Amy';
    data_label5_2.value := 'Bea';
    data_label5_3.value := 'Eva';
    data_label5_4.value := 'Fay';
    data_label5_5.value := 'Liv';
    data_label5_6.value := 'Una';

    (* Create an array of label pointers. nil indicates the end of the array. *)
    data_labels5[0] := @data_label5_1;
    data_labels5[1] := @data_label5_2;
    data_labels5[2] := @data_label5_3;
    data_labels5[3] := @data_label5_4;
    data_labels5[4] := @data_label5_5;
    data_labels5[5] := @data_label5_6;
    data_labels5[6] := nil;

    (* Set the custom labels. *)
    chart_series_set_labels_custom(series, @data_labels5);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D66', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 6. Example with custom data labels from cells.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Chart with custom data labels from cells');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Create some custom labels. *)
    data_label6_1.value := '=Sheet1!$C$2';
    data_label6_2.value := '=Sheet1!$C$3';
    data_label6_3.value := '=Sheet1!$C$4';
    data_label6_4.value := '=Sheet1!$C$5';
    data_label6_5.value := '=Sheet1!$C$6';
    data_label6_6.value := '=Sheet1!$C$7';

    (* Create an array of label pointers. nil indicates the end of the array. *)
    data_labels6[0] := @data_label6_1;
    data_labels6[1] := @data_label6_2;
    data_labels6[2] := @data_label6_3;
    data_labels6[3] := @data_label6_4;
    data_labels6[4] := @data_label6_5;
    data_labels6[5] := @data_label6_6;
    data_labels6[6] := nil;

    (* Set the custom labels. *)
    chart_series_set_labels_custom(series, @data_labels6);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D82', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 7. Example with custom and default data labels.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Mixed custom and default data labels');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    font2.color := Cardinal(LXW_COLOR_RED);

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Create some custom labels. *)

    (* The following is used to get a mix of default and custom labels. The
     * items initialized with '{0}' and items without a custom label (points 5
     * and 6 which come after nil) will get the default value. We also set a
     * font for the custom items as an extra example.
     *)
    data_label7_1.value := '=Sheet1!$C$2';
    data_label7_1.font := @font2;
    chart_data_label_init(data_label7_2);
    data_label7_3.value := '=Sheet1!$C$4';
    data_label7_3.font := @font2;
    data_label7_4.value := '=Sheet1!$C$5';
    data_label7_4.font := @font2;

    (* Create an array of label pointers. nil indicates the end of the array. *)
    data_labels7[0] := @data_label7_1;
    data_labels7[1] := @data_label7_2;
    data_labels7[2] := @data_label7_3;
    data_labels7[3] := @data_label7_4;
    data_labels7[4] := nil;

    (* Set the custom labels. *)
    chart_series_set_labels_custom(series, @data_labels7);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D98', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 8. Example with deleted/hidden custom data labels.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Chart with deleted data labels');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Create some custom labels. *)
    hide.hide := Byte(LXW_TRUE);
    keep.hide := Byte(LXW_FALSE);

    (* An initialized struct like this would also work: *)
    (* lxw_chart_data_label keep := {0}; *)
{    chart_data_label_init(keep); }

    (* Create an array of label pointers. nil indicates the end of the array. *)
    data_labels8[0] := @hide;
    data_labels8[1] := @keep;
    data_labels8[2] := @hide;
    data_labels8[3] := @hide;
    data_labels8[4] := @keep;
    data_labels8[5] := @hide;
    data_labels8[6] := nil;

    (* Set the custom labels. *)
    chart_series_set_labels_custom(series, @data_labels8);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D114', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);


    (*
     * Chart 9.Example with custom string data labels and formatting.
     *)
    chart := workbook_add_chart(workbook, Byte(LXW_CHART_COLUMN));

    (* Add a chart title. *)
    chart_title_set_name(chart, 'Chart with custom labels and formatting');

    (* Add a data series to the chart. *)
    series := chart_add_series(chart, '=Sheet1!$A$2:$A$7',
                                     '=Sheet1!$B$2:$B$7');

    (* Add the series data labels. *)
    chart_series_set_labels(series);

    (* Set the border/line and fill for the data labels. *)
    line2.color := Cardinal(LXW_COLOR_RED);
    fill2.color := Cardinal(LXW_COLOR_YELLOW);
    line3.color := Cardinal(LXW_COLOR_BLUE);
    fill3.color := Cardinal(LXW_COLOR_GREEN);

    (* Create some custom labels. *)
    data_label9_1.value := 'Amy';
    data_label9_1.line := @line3;
    data_label9_2.value := 'Bea';
    data_label9_3.value := 'Eva';
    data_label9_4.value := 'Fay';
    data_label9_5.value := 'Liv';
    data_label9_6.value := 'Una';
    data_label9_6.fill := @fill3;

    (* Set the default formatting for the data labels in the series. *)
    chart_series_set_labels_line(series, @line2);
    chart_series_set_labels_fill(series, @fill2);

    (* Create an array of label pointers. nil indicates the end of the array. *)
    data_labels9[0] :=  @data_label9_1;
    data_labels9[1] :=  @data_label9_2;
    data_labels9[2] :=  @data_label9_3;
    data_labels9[3] :=  @data_label9_4;
    data_labels9[4] :=  @data_label9_5;
    data_labels9[5] :=  @data_label9_6;
    data_labels9[6] :=  nil;

    (* Set the custom labels. *)
    chart_series_set_labels_custom(series, @data_labels9);

    (* Turn off the legend. *)
    chart_legend_set_position(chart, Byte(LXW_CHART_LEGEND_NONE));

    (* Insert the chart into the worksheet. *)
    decodecell('D130', rRow, rCol);
    worksheet_insert_chart_opt(worksheet, rRow, rCol, chart, @options);

  workbook_close(workbook);
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
