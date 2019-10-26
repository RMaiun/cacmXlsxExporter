package mairo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.IntStream;

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
    String intervalWithTimeZone = data.getTimeInterval() + " (" + data.getTimezone() + ")";
    prepareVioletRow(workbook, sheet, intervalWithTimeZone, 3, false);
    // grey time interval with region
    prepareTimeIntervalWithRegionHeaders(workbook, sheet);
    // yellow ti and region values
    prepareTimeIntervalWithRegionValues(workbook, sheet, data);
    // mtu header
    prepareMtuHeader(workbook, sheet);
    int row = 8;

    for (XlsxMtu mtu : data.getMtuList()) {
      prepareMtuValues(workbook, sheet, row, mtu);
      row++;
      prepareConstraintPtdfHeaders(workbook, sheet, row);
      row++;
      prepareConstraintsTableHeaders(workbook, sheet, row, mtu.getConstraintContainer());
      row++;
      for (XlsxConstraintRow constraintRow : mtu.getConstraintContainer().getConstraintRows()) {
        prepareConstraintsTable(workbook, sheet, row, constraintRow);
        row++;
      }
      prepareResultsHeader(workbook, sheet, row);
      row++;
      prepareResultTableHeaders(workbook, sheet, row);
      row++;
      for (ResultXlsxTemplate result : mtu.getResults()) {
        prepareResultRow(workbook, sheet, row, result);
        row++;
      }
    }
    workbook.write(fos);
  }

  private void prepareResultRow(XSSFWorkbook wb, Sheet sheet, int rowLine, ResultXlsxTemplate result) {
    Row row = sheet.createRow(rowLine);
    standardCell(row, 0, result.getDirection(), wb);
    standardCell(row, 1, result.getName(), wb);
    standardCell(row, 2, result.getEic(), wb);
    standardCell(row, 3, result.getType(), wb);
    standardCell(row, 4, result.getLocation(), wb);
    standardCell(row, 5, result.getMaxFlow(), wb);
    standardCell(row, 6, result.getMaxExchange(), wb);
    standardCell(row, 7, result.getMaxExchangeAfterRemedy(), wb);
    standardCell(row, 8, result.getAac(), wb);
    standardCell(row, 9, result.getTrm(), wb);
    standardCell(row, 10, result.getTtc(), wb);
  }

  private void prepareResultTableHeaders(XSSFWorkbook wb, Sheet sheet, int rowLine) {
    Row row = sheet.createRow(rowLine);
    tripleLineHeaderCell(row, 0, "Out Area \n > \n In Area", wb, sheet);
    headerCell(row, 1, "Name", wb);
    headerCell(row, 2, "EIC", wb);
    headerCell(row, 3, "Type", wb);
    headerCell(row, 4, "Location", wb);
    headerCell(row, 5, "Max flow [MW]", wb);
    headerCell(row, 6, "Max Exchange [MW]", wb);
    tripleLineHeaderCell(row, 7, "Max Exchange \n After Remedy [MW]", wb, sheet);
    headerCell(row, 8, "AAC", wb);
    headerCell(row, 9, "TRM", wb);
    headerCell(row, 10, "TTC", wb);
  }

  private void prepareResultsHeader(XSSFWorkbook wb, Sheet sheet, int rowLine) {
    Row row = sheet.createRow(rowLine);
    IntStream.rangeClosed(0, 10).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, "Result", wb);
      } else {
        headerCell(row, i, "", wb);
      }
    });

    CellRangeAddress mergedCell1 = new CellRangeAddress(rowLine, rowLine, 0, 10);
    sheet.addMergedRegion(mergedCell1);
  }

  private void prepareConstraintsTable(XSSFWorkbook workbook, Sheet sheet, int rowLine, XlsxConstraintRow constraintRow) {
    Row row = sheet.createRow(rowLine);
    Cell c0 = row.createCell(0);
    c0.setCellValue(constraintRow.getConstraintId());
    c0.setCellStyle(defaultCellStyle(workbook));

    Cell c1 = row.createCell(1);
    c1.setCellValue(constraintRow.getCbCoName());
    c1.setCellStyle(defaultCellStyle(workbook));

    Cell c2 = row.createCell(2);
    c2.setCellValue(constraintRow.getCbCoEic());
    c2.setCellStyle(defaultCellStyle(workbook));

    Cell c3 = row.createCell(3);
    c3.setCellValue(constraintRow.getCbCoType());
    c3.setCellStyle(defaultCellStyle(workbook));

    Cell c4 = row.createCell(4);
    c4.setCellValue(constraintRow.getCbCoLocation());
    c4.setCellStyle(defaultCellStyle(workbook));

    Cell c5 = row.createCell(5);
    c5.setCellValue(constraintRow.getCbCoAreaIn());
    c5.setCellStyle(defaultCellStyle(workbook));

    Cell c6 = row.createCell(6);
    c6.setCellValue(constraintRow.getCbCoAreaOut());
    c6.setCellStyle(defaultCellStyle(workbook));

    Cell c7 = row.createCell(7);
    c7.setCellValue(constraintRow.getArea1());
    c7.setCellStyle(defaultCellStyle(workbook));

    Cell c8 = row.createCell(8);
    c8.setCellValue(constraintRow.getArea2());
    c8.setCellStyle(defaultCellStyle(workbook));

    Cell c9 = row.createCell(9);
    c9.setCellValue(constraintRow.getArea3());
    c9.setCellStyle(defaultCellStyle(workbook));

    Cell c10 = row.createCell(10);
    c10.setCellValue(constraintRow.getArea4());
    c10.setCellStyle(defaultCellStyle(workbook));

    Cell c11 = row.createCell(11);
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


  private void prepareConstraintsTableHeaders(XSSFWorkbook wb, Sheet sheet, int rowLine, ConstraintsContainer constraintsContainer) {
    Row row = sheet.createRow(rowLine);
    headerCell(row, 0, "Constraint ID", wb);
    headerCell(row, 1, "cb/co name", wb);
    headerCell(row, 2, "cb/co eic", wb);
    headerCell(row, 3, "cb/co type", wb);
    headerCell(row, 4, "cb/co location", wb);
    headerCell(row, 5, "cb/co In Area", wb);
    headerCell(row, 6, "cb/co Out Area", wb);
    headerCell(row, 7, constraintsContainer.getArea1Name(), wb);
    headerCell(row, 8, constraintsContainer.getArea2Name(), wb);
    headerCell(row, 9, constraintsContainer.getArea3Name(), wb);
    headerCell(row, 10, constraintsContainer.getArea4Name(), wb);
    headerCell(row, 11, constraintsContainer.getArea5Name(), wb);
    headerCell(row, 12, "Fref [MW]", wb);
    headerCell(row, 13, "Fmax [MW]", wb);
    headerCell(row, 14, "Actual flow [MW]", wb);
    IntStream.rangeClosed(0, 14).forEach(sheet::autoSizeColumn);
  }

  private void prepareConstraintPtdfHeaders(XSSFWorkbook wb, Sheet sheet, int rowLine) {
    Row row = sheet.createRow(rowLine);
    IntStream.rangeClosed(0, 14).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, "Constraint", wb);
      } else if (i == 7) {
        headerCell(row, i, "PTDF", wb);
      } else {
        headerCell(row, i, "", wb);
      }
    });
    CellRangeAddress mergedCell1 = new CellRangeAddress(rowLine, rowLine, 0, 6);
    CellRangeAddress mergedCell2 = new CellRangeAddress(rowLine, rowLine, 7, 11);
    CellRangeAddress mergedCell3 = new CellRangeAddress(rowLine, rowLine, 12, 14);
    sheet.addMergedRegion(mergedCell1);
    sheet.addMergedRegion(mergedCell2);
    sheet.addMergedRegion(mergedCell3);
  }

  private void prepareMtuValues(XSSFWorkbook wb, Sheet sheet, int rowLine, XlsxMtu mtu) {
    Row row = sheet.createRow(rowLine);
    IntStream.rangeClosed(0, 19).forEach(i -> {
      if (i == 0) {
        yellowCell(row, i, mtu.getInterval(), wb);
      } else if (i == 9) {
        yellowCell(row, i, mtu.getSummary(), wb);
      } else {
        yellowCell(row, i, "", wb);
      }
    });
    CellRangeAddress mergedCell1 = new CellRangeAddress(rowLine, rowLine, 0, 8);
    CellRangeAddress mergedCell2 = new CellRangeAddress(rowLine, rowLine, 9, 19);
    sheet.addMergedRegion(mergedCell1);
    sheet.addMergedRegion(mergedCell2);
  }

  private void prepareMtuHeader(XSSFWorkbook wb, Sheet sheet) {
    int rowNum = 7;
    Row row = sheet.createRow(rowNum);
    IntStream.rangeClosed(0, 19).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, "MTU", wb);
      } else if (i == 9) {
        headerCell(row, i, "MTU Summary", wb);
      } else {
        headerCell(row, i, "", wb);
      }
    });
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
  }

  private void prepareTimeIntervalWithRegionValues(XSSFWorkbook wb, Sheet sheet, NtcBasedAllocationXlsxTemplate data) {
    int rowNum = 6;
    Row row = sheet.createRow(rowNum);
    IntStream.rangeClosed(0, 19).forEach(i -> {
      if (i == 0) {
        yellowCell(row, i, data.getTimeInterval() + " (" + data.getTimezone() + ")", wb);
      } else if (i == 19) {
        yellowCell(row, i, data.getRegion(), wb);
      } else {
        yellowCell(row, i, "", wb);
      }
    });
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
  }


  private void prepareTimeIntervalWithRegionHeaders(XSSFWorkbook wb, Sheet sheet) {
    int rowNum = 5;
    Row row = sheet.createRow(rowNum);
    IntStream.rangeClosed(0, 19).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, "Time Interval", wb);
      } else if (i == 9) {
        headerCell(row, i, "Region", wb);

      } else {
        headerCell(row, i, "", wb);
      }
    });
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
  }

  private void prepareVioletRow(XSSFWorkbook wb, Sheet sheet, String text, int rowLine, boolean boldFont) {
    Row headerRow = sheet.createRow(rowLine);
    int firstCol = 0;
    int lastCol = 19;
    IntStream.rangeClosed(firstCol, lastCol).forEach(i -> {
      Cell cell = headerRow.createCell(i);
      CellStyle violetStyle = violetCellStyle(wb, boldFont);
      if (i == 0) {
        cell.setCellValue(text);
      }
      cell.setCellStyle(violetStyle);
    });
    CellRangeAddress cellRangeAddress = new CellRangeAddress(rowLine, rowLine, firstCol, lastCol);
    sheet.addMergedRegion(cellRangeAddress);
  }

  private CellStyle violetCellStyle(XSSFWorkbook workbook, boolean withBoldFont) {
    XSSFColor violetColor = new XSSFColor(new java.awt.Color(166, 166, 207));
    return customColorCellStyle(violetColor, withBoldFont, workbook);
  }

  private CellStyle greyCellStyle(XSSFWorkbook workbook) {
    XSSFColor greyColor = new XSSFColor(new java.awt.Color(191, 191, 191));
    return customColorCellStyle(greyColor, true, workbook);
  }

  private CellStyle defaultCellStyle(XSSFWorkbook workbook) {
    return customColorCellStyle(null, false, workbook);
  }

  private CellStyle yellowCellStyle(XSSFWorkbook workbook) {
    XSSFColor yellowColor = new XSSFColor(new java.awt.Color(244, 180, 91));
    return customColorCellStyle(yellowColor, false, workbook);
  }

  private CellStyle customColorCellStyle(XSSFColor color, boolean withBoldFont, XSSFWorkbook workbook) {
    Font font = defaultFont(workbook, withBoldFont);
    XSSFCellStyle style = workbook.createCellStyle();
    style.setFont(font);
    if (color != null) {
      style.setFillForegroundColor(color);
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      style.setBorderBottom(BorderStyle.THIN);
      style.setBorderLeft(BorderStyle.THIN);
      style.setBorderRight(BorderStyle.THIN);
      style.setBorderTop(BorderStyle.THIN);
    }

    return style;
  }

  private CellStyle tripleLineGreyCellStyle(XSSFWorkbook workbook, Sheet sheet) {
    Font font = defaultFont(workbook, true);
    XSSFCellStyle style = workbook.createCellStyle();
    style.setFont(font);
    XSSFColor color = new XSSFColor(new java.awt.Color(191, 191, 191));
    style.setFillForegroundColor(color);
    style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    style.setWrapText(true);
    sheet.autoSizeColumn(3);
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

  private void cell(Row row, int column, String value, boolean isHeader, XSSFWorkbook wb, boolean isYellow) {
    Cell cell = row.createCell(column);
    cell.setCellValue(value);
    if (isYellow) {
      cell.setCellStyle(yellowCellStyle(wb));
    } else {
      if (isHeader) {
        CellStyle greyStyle = greyCellStyle(wb);
        greyStyle.setBorderBottom(BorderStyle.THIN);
        greyStyle.setBorderTop(BorderStyle.THIN);
        greyStyle.setBorderLeft(BorderStyle.THIN);
        greyStyle.setBorderRight(BorderStyle.THIN);
        cell.setCellStyle(greyStyle);
      } else {
        cell.setCellStyle(defaultCellStyle(wb));
      }
    }
  }

  private void headerCell(Row row, int column, String value, XSSFWorkbook wb) {
    cell(row, column, value, true, wb, false);
  }

  private void standardCell(Row row, int column, String value, XSSFWorkbook wb) {
    cell(row, column, value, false, wb, false);
  }

  private void yellowCell(Row row, int column, String value, XSSFWorkbook wb) {
    cell(row, column, value, false, wb, true);
  }

  private void tripleLineHeaderCell(Row row, int column, String value, XSSFWorkbook wb, Sheet sheet) {
    row.setHeightInPoints((3 * sheet.getDefaultRowHeightInPoints()));
    Cell c1 = row.createCell(column);
    c1.setCellValue(value);
    CellStyle style = tripleLineGreyCellStyle(wb, sheet);
    style.setBorderBottom(BorderStyle.THIN);
    style.setBorderTop(BorderStyle.THIN);
    style.setBorderLeft(BorderStyle.THIN);
    style.setBorderRight(BorderStyle.THIN);
    c1.setCellStyle(style);
  }

}
