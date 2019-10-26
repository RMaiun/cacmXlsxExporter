package mairo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class UberWriter {
  public static final String SHEET_NAME = "CACM_2.2&2.3 flow-based";

  public void generateTemplate(NtcBasedAllocationXlsxTemplate data, FileOutputStream fos) throws IOException {
    XSSFWorkbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet(SHEET_NAME);
    // row #1
    prepareVioletRow(workbook, sheet, "NTC-based capacity allocation and network utilisation", 0, true);
    //row #2
    prepareVioletRow(workbook, sheet, "NTC-based capacity allocation and network utilisation [CACM 2.2&2.3]", 1, false);
    // row #3
    prepareVioletRow(workbook, sheet, "NTC-based capacity allocation and network utilisation - Result [CACM 2.2&2.3]", 2, false);
    // violet interval row
    String intervalWithTimeZone  = data.getTimeInterval()+ " (" + data.getTimezone()+ ")";
    prepareVioletRow(workbook, sheet, intervalWithTimeZone, 3, false);
    // grey time interval with region
    prepareTimeIntervalWithRegion(workbook, sheet);
    // yellow ti and region values
    prepareTimeIntervalWithRegionValues(workbook, sheet, data);
    // mtu header
    prepareMtuHeader(workbook, sheet);
    int row = 8;

    for (XlsxMtu mtu : data.getMtuList()) {
      prepareMtuValues(workbook, sheet, row,mtu);
      row ++;
      prepareConstraintPtdfHeaders(workbook, sheet, row);
      row ++;
      prepareConstraintsTableHeaders(workbook, sheet, row, mtu.getConstraintContainer());
      row ++;


    }
    workbook.write(fos);
  }

  private void prepareConstraintsTable(XSSFWorkbook workbook, Sheet sheet, int rowLine, XlsxConstraintRow constraintRow){
    Row row = sheet.createRow(rowLine);
    Cell c0= row.createCell(0);
    c0.setCellValue(constraintRow.getConstraintId());
    c0.setCellStyle(defaultCellStyle(workbook));

    Cell c1= row.createCell(1);
    c1.setCellValue(constraintRow.getCbCoName());
    c1.setCellStyle(defaultCellStyle(workbook));

    Cell c2= row.createCell(2);
    c2.setCellValue(constraintRow.getCbCoEic());
    c2.setCellStyle(defaultCellStyle(workbook));

    Cell c3= row.createCell(3);
    c3.setCellValue(constraintRow.getCbCoType());
    c3.setCellStyle(defaultCellStyle(workbook));

    Cell c4= row.createCell(4);
    c4.setCellValue(constraintRow.getCbCoLocation());
    c4.setCellStyle(defaultCellStyle(workbook));

    Cell c5= row.createCell(5);
    c5.setCellValue(constraintRow.getCbCoAreaIn());
    c5.setCellStyle(defaultCellStyle(workbook));

    Cell c6= row.createCell(6);
    c6.setCellValue(constraintRow.getCbCoAreaOut());
    c6.setCellStyle(defaultCellStyle(workbook));

    Cell c7= row.createCell(7);
    c7.setCellValue(constraintRow.getArea1());
    c7.setCellStyle(defaultCellStyle(workbook));

    Cell c8= row.createCell(8);
    c8.setCellValue(constraintRow.getArea2());
    c8.setCellStyle(defaultCellStyle(workbook));

    Cell c9= row.createCell(9);
    c9.setCellValue(constraintRow.getArea3());
    c9.setCellStyle(defaultCellStyle(workbook));

    Cell c10= row.createCell(10);
    c10.setCellValue(constraintRow.getArea4());
    c10.setCellStyle(defaultCellStyle(workbook));

    Cell c11= row.createCell(11);
    c11.setCellValue(constraintRow.getArea5());
    c11.setCellStyle(defaultCellStyle(workbook));

    Cell c12 = row.createCell(12);
    c12.setCellValue(constraintRow.getfRef());
    c12.setCellStyle(defaultCellStyle(workbook));

    Cell c13 = row.createCell(13);
    c13.setCellValue(constraintRow.getfMax());
    c13.setCellStyle(defaultCellStyle(workbook));

    Cell c14 = row.createCell(14);
    c14.setCellValue(constraintRow.getActualFlow());
    c14.setCellStyle(defaultCellStyle(workbook));
  }


  private void prepareConstraintsTableHeaders(XSSFWorkbook workbook, Sheet sheet, int rowLine, ConstraintsContainer constraintsContainer){
    Row row = sheet.createRow(rowLine);
    Cell c0= row.createCell(0);
    c0.setCellValue("Constraint ID");
    c0.setCellStyle(greyCellStyle(workbook, true));

    Cell c1= row.createCell(1);
    c1.setCellValue("cb/co name");
    c1.setCellStyle(greyCellStyle(workbook, true));

    Cell c2= row.createCell(2);
    c2.setCellValue("cb/co eic");
    c2.setCellStyle(greyCellStyle(workbook, true));

    Cell c3= row.createCell(3);
    c3.setCellValue("cb/co type");
    c3.setCellStyle(greyCellStyle(workbook, true));

    Cell c4= row.createCell(4);
    c4.setCellValue("cb/co location");
    c4.setCellStyle(greyCellStyle(workbook, true));

    Cell c5= row.createCell(5);
    c5.setCellValue("cb/co In Area");
    c5.setCellStyle(greyCellStyle(workbook, true));

    Cell c6= row.createCell(6);
    c6.setCellValue("cb/co Out Area");
    c6.setCellStyle(greyCellStyle(workbook, true));

    Cell c7= row.createCell(7);
    c7.setCellValue(constraintsContainer.getArea1Name());
    c7.setCellStyle(greyCellStyle(workbook, true));

    Cell c8= row.createCell(8);
    c8.setCellValue(constraintsContainer.getArea2Name());
    c8.setCellStyle(greyCellStyle(workbook, true));

    Cell c9= row.createCell(9);
    c9.setCellValue(constraintsContainer.getArea3Name());
    c9.setCellStyle(greyCellStyle(workbook, true));

    Cell c10= row.createCell(10);
    c10.setCellValue(constraintsContainer.getArea4Name());
    c10.setCellStyle(greyCellStyle(workbook, true));

    Cell c11= row.createCell(11);
    c11.setCellValue(constraintsContainer.getArea5Name());
    c11.setCellStyle(greyCellStyle(workbook, true));

    Cell c12 = row.createCell(12);
    c12.setCellValue("Fref [MW]");
    c12.setCellStyle(greyCellStyle(workbook, true));

    Cell c13 = row.createCell(13);
    c13.setCellValue("Fmax [MW]");
    c13.setCellStyle(greyCellStyle(workbook, true));

    Cell c14 = row.createCell(14);
    c14.setCellValue("Actual flow [MW]");
    c14.setCellStyle(greyCellStyle(workbook, true));

    sheet.autoSizeColumn(0);
    sheet.autoSizeColumn(1);
    sheet.autoSizeColumn(2);
    sheet.autoSizeColumn(3);
    sheet.autoSizeColumn(4);
    sheet.autoSizeColumn(5);
    sheet.autoSizeColumn(6);
    sheet.autoSizeColumn(7);
    sheet.autoSizeColumn(8);
    sheet.autoSizeColumn(9);
    sheet.autoSizeColumn(10);
    sheet.autoSizeColumn(11);
    sheet.autoSizeColumn(12);
    sheet.autoSizeColumn(13);
    sheet.autoSizeColumn(14);
  }

  private void prepareConstraintPtdfHeaders(XSSFWorkbook workbook, Sheet sheet, int row){
    Row headerRow = sheet.createRow(row);
    CellRangeAddress mergedCell1 = new CellRangeAddress(row, row, 0, 6);
    CellRangeAddress mergedCell2 = new CellRangeAddress(row, row, 7, 11);
    CellRangeAddress mergedCell3 = new CellRangeAddress(row, row, 12, 14);
    sheet.addMergedRegion(mergedCell1);
    sheet.addMergedRegion(mergedCell2);
    sheet.addMergedRegion(mergedCell3);
    Cell c1= headerRow.createCell(0);
    c1.setCellValue("Constraint");
    c1.setCellStyle(greyCellStyle(workbook, true));

    Cell c2 = headerRow.createCell(7);
    c2.setCellValue("PTDF");
    c2.setCellStyle(greyCellStyle(workbook, true));
    Cell c3 = headerRow.createCell(12);
    c3.setCellValue("");
    c3.setCellStyle(greyCellStyle(workbook, false));

  }

  private void prepareMtuValues(XSSFWorkbook workbook, Sheet sheet, int row, XlsxMtu mtu){
    Row headerRow = sheet.createRow(row);
    CellRangeAddress mergedCell1 = new CellRangeAddress(row, row, 0, 8);
    CellRangeAddress mergedCell2 = new CellRangeAddress(row, row, 9, 19);
    sheet.addMergedRegion(mergedCell1);
    sheet.addMergedRegion(mergedCell2);
    Cell tiCell= headerRow.createCell(0);
    tiCell.setCellValue(mtu.getInterval());
    tiCell.setCellStyle(yellowCellStyle(workbook, false));

    Cell regionCell = headerRow.createCell(9);
    regionCell.setCellValue(mtu.getSummary());
    regionCell.setCellStyle(yellowCellStyle(workbook, true));

  }

  private void prepareMtuHeader(XSSFWorkbook workbook, Sheet sheet){
    int rowNum = 7;
    Row headerRow = sheet.createRow(rowNum);
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
    Cell tiCell= headerRow.createCell(0);
    tiCell.setCellValue("MTU");
    tiCell.setCellStyle(greyCellStyle(workbook, true));

    Cell regionCell = headerRow.createCell(9);
    regionCell.setCellValue("MTU Summary");
    regionCell.setCellStyle(greyCellStyle(workbook, true));
  }
  private void prepareTimeIntervalWithRegionValues(XSSFWorkbook workbook, Sheet sheet, NtcBasedAllocationXlsxTemplate data) {
    int rowNum = 6;
    Row headerRow = sheet.createRow(rowNum);
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
    Cell tiCell= headerRow.createCell(0);
    tiCell.setCellValue(data.getTimeInterval()+ " (" + data.getTimezone()+ ")");
    tiCell.setCellStyle(yellowCellStyle(workbook, false));

    Cell regionCell = headerRow.createCell(9);
    regionCell.setCellValue(data.getRegion());
    regionCell.setCellStyle(yellowCellStyle(workbook, false));
  }


    private void prepareTimeIntervalWithRegion(XSSFWorkbook workbook, Sheet sheet){
    int rowNum = 5;
    Row headerRow = sheet.createRow(rowNum);
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
    Cell tiCell= headerRow.createCell(0);
    tiCell.setCellValue("Time Interval");
    tiCell.setCellStyle(greyCellStyle(workbook, true));

    Cell regionCell = headerRow.createCell(9);
    regionCell.setCellValue("Region");
    regionCell.setCellStyle(greyCellStyle(workbook, true));

  }

  private void prepareVioletRow(XSSFWorkbook workbook, Sheet sheet, String text, int row, boolean boldFont) {
    Row headerRow = sheet.createRow(row);
    CellRangeAddress cellRangeAddress = new CellRangeAddress(row, row, 0, 19);
    sheet.addMergedRegion(cellRangeAddress);
    Cell cell = headerRow.createCell(0);
    cell.setCellValue(text);
    cell.setCellStyle(violetCellStyle(workbook, boldFont));
  }

  private CellStyle violetCellStyle(XSSFWorkbook workbook, boolean withBoldFont) {
    XSSFColor violetColor = new XSSFColor(new java.awt.Color(166, 166, 207));
    return customColorCellStyle(violetColor, withBoldFont, workbook);
  }

  private CellStyle greyCellStyle(XSSFWorkbook workbook, boolean withBoldFont) {
    XSSFColor greyColor = new XSSFColor(new java.awt.Color(191, 191, 191));
    return customColorCellStyle(greyColor, withBoldFont, workbook);
  }

  private CellStyle defaultCellStyle(XSSFWorkbook workbook) {
    return customColorCellStyle(null, false, workbook);
  }

  private CellStyle yellowCellStyle(XSSFWorkbook workbook, boolean withBoldFont) {
    XSSFColor yellowColor = new XSSFColor(new java.awt.Color(244, 180, 91));
    return customColorCellStyle(yellowColor, withBoldFont, workbook);
  }

  private CellStyle customColorCellStyle(XSSFColor color, boolean withBoldFont, XSSFWorkbook workbook){
    Font font = defaultFont(workbook, withBoldFont);
    XSSFCellStyle style = workbook.createCellStyle();
//    style.setBorderBottom(BorderStyle.THIN);
//    style.setBorderTop(BorderStyle.THIN);
//    style.setBorderLeft(BorderStyle.THIN);
//    style.setBorderRight(BorderStyle.THIN);
    style.setFont(font);
    if (color != null){
      style.setFillForegroundColor(color);
    }
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//    style.setWrapText(false);

    return style;
  }

  private Font defaultFont(XSSFWorkbook workbook, boolean isBold) {
    Font font = workbook.createFont();
    font.setFontName("Arial");
    font.setBold(isBold);
    font.setFontHeightInPoints((short) 10);
    font.setColor(IndexedColors.BLACK.getIndex());
    return font;
  }
}
