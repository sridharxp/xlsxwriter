program merge_range;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  xlsxwriterapi in '..\..\..\DL\LXW\xlsxwriterapi.pas';

var
  workbook: Plxw_workbook;
  worksheet: Plxw_worksheet;
  ReportName: PAnsiChar;
  merge_format: Plxw_format;
begin
  ReportName := 'merge_range.xlsx';
  workbook  := workbook_new(ReportName);
  worksheet := workbook_add_worksheet(workbook, nil);
  merge_format := workbook_add_format(workbook);

  // Configure a format for the merged range.
  format_set_align(merge_format, Byte(LXW_ALIGN_CENTER));
  format_set_align(merge_format, Byte(LXW_ALIGN_VERTICAL_CENTER));
  format_set_bold(merge_format);
  format_set_bg_color(merge_format, UInt32(LXW_COLOR_YELLOW));
  format_set_border(merge_format, Byte(LXW_BORDER_THIN));

  // Increase the cell size of the merged cells to highlight the formatting.
  worksheet_set_column(worksheet, 1, 3, 12, nil);
  worksheet_set_row(worksheet, 3, 30, nil);
  worksheet_set_row(worksheet, 6, 30, nil);
  worksheet_set_row(worksheet, 7, 30, nil);

  // Merge 3 cells.
  worksheet_merge_range(worksheet, 3, 1, 3, 3, 'Merged Range', merge_format);

  // Merge 3 cells over two rows.
  worksheet_merge_range(worksheet, 6, 1, 7, 3, 'Merged Range', merge_format);

  workbook_close(workbook);

  ExitCode := 0;
{
  ShellExecute(Self.Handle, Pchar('Open'), ReportName,
      nil, nil, SW_SHOWNORMAL);
}
end.
