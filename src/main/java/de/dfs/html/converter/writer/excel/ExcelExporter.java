/**
 * Copyright (C) 2012 alanhay <alanhay99@hotmail.com>
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.dfs.html.converter.writer.excel;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.jsoup.nodes.Element;

import de.dfs.html.converter.writer.AbstractExporter;

public class ExcelExporter extends AbstractExporter {
    public void exportHtml(String html, OutputStream out) throws IOException {
        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = null;
        int startRow = 0;

        for (Element element : getTables(html)) {

            if (workbook.getNumberOfSheets() == 0) {
                String sheetName = getSheetName(element);

                if (StringUtils.isNotEmpty(sheetName)) {
                    sheet = workbook.createSheet(sheetName);
                } else {
                    sheet = workbook.createSheet();
                }
            } else if (isNewSheet(element)) {
                String sheetName = getSheetName(element);

                if (StringUtils.isNotEmpty(sheetName))
                    sheet = workbook.createSheet(sheetName);
                else {
                    sheet = workbook.createSheet();
                }

                startRow = 0;
            }

            //TableWriter writer = new ExcelTableWriter(new ExcelTableRowWriter(sheet, new ExcelTableCellWriter(sheet
            //)));


            sheet.setDisplayGridlines(false);
            sheet.setMargin(Sheet.LeftMargin,500);

            //sheet.setColumnHidden(1,true);


            Cell cell;


            for (int i = 0; i < 4; i++) {
                Row row = sheet.createRow(i);


                if (i == 0) {
                    row.setHeight((short) 1200);
                    Font font = workbook.createFont();
                    font.setFontHeightInPoints((short) 15);
                    font.setFontName("Courier New");




                    cell = row.createCell(0);
                    cell.setCellValue("test");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setFont(font);
                    cell.getCellStyle().setFillBackgroundColor(IndexedColors.BLACK.index);
                    cell.getCellStyle().setFillPattern(FillPatternType.BIG_SPOTS);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.BLACK.getIndex());



                    cell = row.createCell(1);
                    cell.setCellValue("test2");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setFont(font);
                    cell.getCellStyle().setFillBackgroundColor(IndexedColors.BLACK.index);
                    cell.getCellStyle().setFillPattern(FillPatternType.BIG_SPOTS);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.BLACK.getIndex());


                    cell = row.createCell(2);
                    cell.setCellValue("test3");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setFont(font);
                    cell.getCellStyle().setFillBackgroundColor(IndexedColors.BLACK.index);
                    cell.getCellStyle().setFillPattern(FillPatternType.BIG_SPOTS);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.BLACK.getIndex());

                    cell = row.createCell(3);
                    cell.setCellValue("test4");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setFont(font);
                    cell.getCellStyle().setFillBackgroundColor(IndexedColors.BLACK.index);
                    cell.getCellStyle().setFillPattern(FillPatternType.BIG_SPOTS);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.BLACK.getIndex());

                    cell = row.createCell(4);
                    cell.setCellValue("test5");
                    cell.setCellStyle(getCellStyleStandart(workbook));
                    cell.getCellStyle().setFont(font);
                    cell.getCellStyle().setFillBackgroundColor(IndexedColors.BLACK.index);
                    cell.getCellStyle().setFillPattern(FillPatternType.BIG_SPOTS);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.BLACK.getIndex());

                    cell = row.createCell(5);
                    cell.setCellValue("test6");
                    cell.setCellStyle(getCellStyleStandart(workbook));
                    cell.getCellStyle().setFont(font);
                    cell.getCellStyle().setFillBackgroundColor(IndexedColors.BLACK.index);
                    cell.getCellStyle().setFillPattern(FillPatternType.BIG_SPOTS);
                    cell.getCellStyle().setFillForegroundColor(IndexedColors.BLACK.getIndex());
                }


                if (i == 1) {
                    row.setHeight((short) 800);

                    cell = row.createCell(0);
                    cell.setCellValue("ZZZ");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(1);
                    cell.setCellValue("16038");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(2);
                    cell.setCellValue("ZZZ");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(3);
                    cell.setCellValue("16038");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(4);
                    cell.setCellValue("10.20.20");
                    cell.setCellStyle(getCellStyleStandart(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(5);
                    cell.setCellValue("20.10.22");
                    cell.setCellStyle(getCellStyleStandart(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);


                }

                if (i == 2) {
                    row.setHeight((short) 800);

                    cell = row.createCell(0);
                    cell.setCellValue("ZWT");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);


                    cell = row.createCell(1);
                    cell.setCellValue("11659");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(2);
                    cell.setCellValue("ZWT");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                    cell = row.createCell(3);
                    cell.setCellValue("11659");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);


                    cell = row.createCell(4);
                    cell.setCellValue("Tätigkeit / Buchungsstatus / Einschränkungen");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.DISTRIBUTED);

                    cell = row.createCell(5);
                    cell.setCellValue(" ");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));
                    cell.getCellStyle().setAlignment(HorizontalAlignment.CENTER);


                }

                if (i == 3) {
                    row.setHeight((short) 800);

                    cell = row.createCell(0);
                    cell.setCellValue("ZWS");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));

                    cell = row.createCell(1);
                    cell.setCellValue("16129");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));

                    cell = row.createCell(2);
                    cell.setCellValue("ZWS");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));

                    cell = row.createCell(3);
                    cell.setCellValue("16129");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));


                    cell = row.createCell(4);
                    cell.setCellValue("*");
                    cell.setCellStyle(getCellStyleNoRightBorder(workbook));

                    cell = row.createCell(5);
                    cell.setCellValue(" ");
                    cell.setCellStyle(getCellStyleNoLeftBorder(workbook));


                }

            }


            sheet.setVerticallyCenter(true);
            //sheet.setMargin((short) 5,15);

			/*sheet.addMergedRegionUnsafe(CellRangeAddress.valueOf("A1:B1"));
			sheet.addMergedRegionUnsafe(CellRangeAddress.valueOf("C1:D1"));*/
            sheet.addMergedRegionUnsafe(CellRangeAddress.valueOf("E3:F3"));
            sheet.addMergedRegionUnsafe(CellRangeAddress.valueOf("E4:F4"));


            //startRow += writer.writeTable(element, null, startRow) + 1;
            //sheet.createRow(startRow);


        }

        for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
            formatSheet(workbook.getSheetAt(i));
        }

        workbook.write(out);
        out.flush();
        out.close();
    }

    protected CellStyle getCellHeaderStyle(Workbook workbook) {
        CellStyle cellHeaderStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 15);
        font.setFontName("Courier New");
        cellHeaderStyle.setFont(font);
        cellHeaderStyle.setAlignment(HorizontalAlignment.CENTER);

        return cellHeaderStyle;

    }

    public CellStyle changeCellBackgroundColorWithPattern(Cell cell) {
        CellStyle cellStyleBackground = cell.getCellStyle();
        if (cellStyleBackground == null) {
            cellStyleBackground = cell.getSheet().getWorkbook().createCellStyle();
        }
        cellStyleBackground.setFillBackgroundColor(IndexedColors.BLACK.index);
        cellStyleBackground.setFillPattern(FillPatternType.BIG_SPOTS);
        cellStyleBackground.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        cell.setCellStyle(cellStyleBackground);
        return cellStyleBackground;
    }


    protected CellStyle getCellStyleStandart(Workbook workbook) {
        final XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        return cellStyle;
    }

    protected CellStyle getCellStyleNoRightBorder(Workbook workbook) {
        final XSSFCellStyle cellStyleNoRightBorder = (XSSFCellStyle) workbook.createCellStyle();
        cellStyleNoRightBorder.setBorderBottom(BorderStyle.THIN);
        cellStyleNoRightBorder.setBorderTop(BorderStyle.THIN);
        cellStyleNoRightBorder.setBorderLeft(BorderStyle.THIN);
        return cellStyleNoRightBorder;
    }

    protected CellStyle getCellStyleNoLeftBorder(Workbook workbook) {
        final XSSFCellStyle cellStyleNoLeftBorder = (XSSFCellStyle) workbook.createCellStyle();
        cellStyleNoLeftBorder.setBorderBottom(BorderStyle.THIN);
        cellStyleNoLeftBorder.setBorderTop(BorderStyle.THIN);
        cellStyleNoLeftBorder.setBorderRight(BorderStyle.THIN);
        return cellStyleNoLeftBorder;
    }


    protected void formatSheet(Sheet sheet) {
        int lastRowWithData = 0;

        for (int i = sheet.getLastRowNum(); i >= 0; --i) {
            if (sheet.getRow(i) != null && sheet.getRow(i).getPhysicalNumberOfCells() > 0) {
                lastRowWithData = i;
                break;
            }
        }

        for (int i = 0; i < sheet.getRow(lastRowWithData).getPhysicalNumberOfCells(); ++i) {
            sheet.autoSizeColumn(i);
        }

        for (int i = 0; i < sheet.getRow(sheet.getLastRowNum()).getPhysicalNumberOfCells(); ++i) {
            sheet.setColumnWidth(i, (int) (sheet.getColumnWidth(i) * 1.2));
        }
    }
}
