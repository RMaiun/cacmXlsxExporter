package mairo;

import mairo.dto.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.stream.IntStream;

import static mairo.CacmXlsxConstants.*;

public class CacmXlsxWriter {


  public void generateTemplate(NtcBasedAllocationXlsxTemplate data, FileOutputStream fos) throws IOException {
    XSSFWorkbook workbook = new XSSFWorkbook();
    Sheet sheet = workbook.createSheet(SHEET_NAME);
    // row #1
    prepareVioletRow(workbook, sheet, GENERAL_INFO, 0, true);
    //row #2
    prepareVioletRow(workbook, sheet, NTC_BASED_CAPACITY_ALLOCATION_AND_NETWORK_UTILISATION, 1, false);
    // row #3
    prepareVioletRow(workbook, sheet, NTC_BASED_CAPACITY_ALLOCATION_AND_NETWORK_UTILISATION_RESULT, 2, false);
    // violet interval row
    String intervalWithTimeZone = data.getTimeInterval() + LB_SPACED + data.getTimezone() + RB;
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
      for (ConstraintXlsxTemplate constraintRow : mtu.getConstraintContainer().getConstraintRows()) {
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
    tripleLineHeaderCell(row, 0, RESULT_DIRECTION_HEADER, wb, sheet);
    headerCell(row, 1, RESULT_NAME_HEADER, wb);
    headerCell(row, 2, RESULT_EIC_HEADER, wb);
    headerCell(row, 3, RESULT_TYPE_HEADER, wb);
    headerCell(row, 4, RESULT_LOCATION_HEADER, wb);
    headerCell(row, 5, RESULT_MAX_FLOW_HEADER, wb);
    headerCell(row, 6, RESULT_MAX_EXCHANGE_HEADER, wb);
    tripleLineHeaderCell(row, 7, RESULT_MAX_EXCHANGE_AFTER_REMEDY_HEADER, wb, sheet);
    headerCell(row, 8, RESULT_AAC_HEADER, wb);
    headerCell(row, 9, RESULT_TRM_HEADER, wb);
    headerCell(row, 10, RESULT_TTC_HEADER, wb);
  }

  private void prepareResultsHeader(XSSFWorkbook wb, Sheet sheet, int rowLine) {
    Row row = sheet.createRow(rowLine);
    IntStream.rangeClosed(0, 10).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, RESULT_HEADER, wb);
      } else {
        headerCell(row, i, EMPTY_STRING, wb);
      }
    });

    CellRangeAddress mergedCell1 = new CellRangeAddress(rowLine, rowLine, 0, 10);
    sheet.addMergedRegion(mergedCell1);
  }

  private void prepareConstraintsTable(XSSFWorkbook workbook, Sheet sheet, int rowLine, ConstraintXlsxTemplate constraintRow) {
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


  private void prepareConstraintsTableHeaders(XSSFWorkbook wb, Sheet sheet, int rowLine, ConstraintsContainerXlsxTemplate constraintsContainer) {
    Row headersRow = sheet.createRow(rowLine);
    headerCell(headersRow, 0, CONSTRAINT_ID_HEADER, wb);
    headerCell(headersRow, 1, CONSTRAINT_NAME_HEADER, wb);
    headerCell(headersRow, 2, CONSTRAINT_EIC_HEADER, wb);
    headerCell(headersRow, 3, CONSTRAINT_TYPE_HEADER, wb);
    headerCell(headersRow, 4, CONSTRAINT_LOCATION_HEADER, wb);
    headerCell(headersRow, 5, CONSTRAINT_IN_AREA_HEADER, wb);
    headerCell(headersRow, 6, CONSTRAINT_OUT_AREA_HEADER, wb);
    headerCell(headersRow, 7, constraintsContainer.getArea1Name(), wb);
    headerCell(headersRow, 8, constraintsContainer.getArea2Name(), wb);
    headerCell(headersRow, 9, constraintsContainer.getArea3Name(), wb);
    headerCell(headersRow, 10, constraintsContainer.getArea4Name(), wb);
    headerCell(headersRow, 11, constraintsContainer.getArea5Name(), wb);
    headerCell(headersRow, 12, CONSTRAINT_FREF_HEADER, wb);
    headerCell(headersRow, 13, CONSTRAINT_FMAX_HEADER, wb);
    headerCell(headersRow, 14, CONSTRAINT_ACTUAL_FLOW_HEADER, wb);
    IntStream.rangeClosed(0, 14).forEach(sheet::autoSizeColumn);
  }

  private void prepareConstraintPtdfHeaders(XSSFWorkbook wb, Sheet sheet, int rowLine) {
    Row row = sheet.createRow(rowLine);
    IntStream.rangeClosed(0, 14).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, CONSTRAINT_HEADER, wb);
      } else if (i == 7) {
        headerCell(row, i, PTDF_HEADER, wb);
      } else {
        headerCell(row, i, EMPTY_STRING, wb);
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
        yellowCell(row, i, mtuSummary(mtu.getConstraintsNum(), mtu.getResultsNum()), wb);
      } else {
        yellowCell(row, i, EMPTY_STRING, wb);
      }
    });
    CellRangeAddress mergedCell1 = new CellRangeAddress(rowLine, rowLine, 0, 8);
    CellRangeAddress mergedCell2 = new CellRangeAddress(rowLine, rowLine, 9, 19);
    sheet.addMergedRegion(mergedCell1);
    sheet.addMergedRegion(mergedCell2);
  }

  private void prepareMtuHeader(XSSFWorkbook wb, Sheet sheet) {
    prepareGreyHeadersBlock(wb, sheet, 7, MTU_HEADER, MTU_SUMMARY_HEADER);
  }

  private void prepareTimeIntervalWithRegionValues(XSSFWorkbook wb, Sheet sheet, NtcBasedAllocationXlsxTemplate data) {
    int rowNum = 6;
    Row row = sheet.createRow(rowNum);
    IntStream.rangeClosed(0, 19).forEach(i -> {
      if (i == 0) {
        yellowCell(row, i, data.getTimeInterval() + LB_SPACED + data.getTimezone() + RB, wb);
      } else if (i == 19) {
        yellowCell(row, i, data.getRegion(), wb);
      } else {
        yellowCell(row, i, EMPTY_STRING, wb);
      }
    });
    CellRangeAddress tiRange = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress regionRange = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(tiRange);
    sheet.addMergedRegion(regionRange);
  }


  private void prepareTimeIntervalWithRegionHeaders(XSSFWorkbook wb, Sheet sheet) {
    prepareGreyHeadersBlock(wb, sheet, 5, TIME_INTERVAL_HEADER, REGION_HEADER);
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
    font.setFontName(DEFAULT_FONT);
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

  private void prepareGreyHeadersBlock(XSSFWorkbook wb, Sheet sheet, int rowNum, String data0, String data19) {
    Row row = sheet.createRow(rowNum);
    IntStream.rangeClosed(0, 19).forEach(i -> {
      if (i == 0) {
        headerCell(row, i, data0, wb);
      } else if (i == 9) {
        headerCell(row, i, data19, wb);

      } else {
        headerCell(row, i, EMPTY_STRING, wb);
      }
    });
    CellRangeAddress range1 = new CellRangeAddress(rowNum, rowNum, 0, 8);
    CellRangeAddress range2 = new CellRangeAddress(rowNum, rowNum, 9, 19);
    sheet.addMergedRegion(range1);
    sheet.addMergedRegion(range2);
  }

  private String mtuSummary(int constraints, int results) {
    return String.format("There is/are %o CB/CO and %o results in MTU.", constraints, results);
  }

}
