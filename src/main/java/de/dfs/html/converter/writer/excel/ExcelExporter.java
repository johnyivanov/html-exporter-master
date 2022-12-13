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

import de.dfs.html.converter.writer.TableWriter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.jsoup.nodes.Element;

import de.dfs.html.converter.writer.AbstractExporter;

import javax.swing.table.TableCellEditor;

import static org.apache.poi.ss.usermodel.VerticalAlignment.DISTRIBUTED;

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

            //TableWriter writer = new ExcelTableWriter(new ExcelTableRowWriter(sheet, new ExcelTableCellWriter(sheet)));


            sheet.setDisplayGridlines(false);


            //sheet.setColumnHidden(1


            Cell cell;



            for (int i = 0; i < 4; i++) {
                Row row = sheet.createRow(i);



                if (i == 0) {
                    row.setHeight((short) 1000);

                    cell = row.createCell(0);
                    cell.setCellValue("test");
                    cell.setCellStyle(getCellHeaderStyle(workbook,cell));

                    cell = row.createCell(1);
                    cell.setCellValue("test2");
                    cell.setCellStyle(getCellHeaderStyle(workbook,cell));

                    cell = row.createCell(2);
                    cell.setCellValue("test3");
                    cell.setCellStyle(getCellHeaderStyle(workbook,cell));

                    cell = row.createCell(3);
                    cell.setCellValue("test4");
                    cell.setCellStyle(getCellHeaderStyle(workbook,cell));

                    cell = row.createCell(4);
                    cell.setCellValue("test5");
                    cell.setCellStyle(getCellHeaderStyle(workbook,cell));

                    cell = row.createCell(5);
                    cell.setCellValue("test6");
                    cell.setCellStyle(getCellHeaderStyle(workbook,cell));
                }


                if (i == 1) {
                    row.setHeight((short) 800);

                    cell = row.createCell(0);
                    cell.setCellValue("ZZZ");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(1);
                    cell.setCellValue("16038");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(2);
                    cell.setCellValue("ZZZ");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(3);
                    cell.setCellValue("16038");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(4);
                    cell.setCellValue("10.20.20");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));


                    cell = row.createCell(5);
                    cell.setCellValue("10.10.10");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));


                }

                if (i == 2) {
                    row.setHeight((short) 800);

                    cell = row.createCell(0);
                    cell.setCellValue("ZWT");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(1);
                    cell.setCellValue("11659");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(2);
                    cell.setCellValue("ZWT");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(3);
                    cell.setCellValue("11659");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(4);
                    cell.setCellValue("Tätigkeit / Buchungsstatus / Einschränkungen");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));


                    cell = row.createCell(5);
                    cell.setCellValue(" ");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));



                }

                if (i == 3) {
                    row.setHeight((short) 800);

                    cell = row.createCell(0);
                    cell.setCellValue("ZWS");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(1);
                    cell.setCellValue("16129");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(2);
                    cell.setCellValue("ZWS");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(3);
                    cell.setCellValue("16129");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));


                    cell = row.createCell(4);
                    cell.setCellValue("*");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));

                    cell = row.createCell(5);
                    cell.setCellValue(" ");
                    cell.setCellStyle(getCellStyleStandart(workbook,cell));


                }

            }



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

    protected CellStyle getCellHeaderStyle(Workbook workbook, Cell cell) {
        CellStyle cellHeaderStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 15);
        font.setFontName("Courier New");
        cellHeaderStyle.setFont(font);
        cellHeaderStyle.setAlignment(HorizontalAlignment.CENTER);

        cellHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellHeaderStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());

        if(cell.getColumnIndex() ==0 || cell.getColumnIndex()==2){
              //no right border

            cellHeaderStyle.setBorderTop(BorderStyle.THIN);
            cellHeaderStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderLeft(BorderStyle.THIN);
            cellHeaderStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderBottom(BorderStyle.THIN);
            cellHeaderStyle.setBottomBorderColor(IndexedColors.RED.getIndex());


        }
       else if(cell.getColumnIndex()==1 || cell.getColumnIndex()==3) {
            //no left border

            cellHeaderStyle.setBorderTop(BorderStyle.THIN);
           cellHeaderStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderRight(BorderStyle.THIN);
            cellHeaderStyle.setRightBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderBottom(BorderStyle.THIN);
            cellHeaderStyle.setBottomBorderColor(IndexedColors.RED.getIndex());



        }
        else if(cell.getColumnIndex()==5) {
            //no left border

            cellHeaderStyle.setBorderTop(BorderStyle.THIN);
            cellHeaderStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderRight(BorderStyle.THIN);
            cellHeaderStyle.setRightBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderBottom(BorderStyle.THIN);
            cellHeaderStyle.setBottomBorderColor(IndexedColors.RED.getIndex());



        }
        else {


            cellHeaderStyle.setBorderTop(BorderStyle.THIN);
            cellHeaderStyle.setTopBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderRight(BorderStyle.THIN);
           cellHeaderStyle.setRightBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderLeft(BorderStyle.THIN);
            cellHeaderStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
            cellHeaderStyle.setBorderBottom(BorderStyle.THIN);
            cellHeaderStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
        }

        return cellHeaderStyle;

    }




    protected CellStyle getCellStyleStandart(Workbook workbook, Cell cell) {
        CellStyle cellStyleStandart = workbook.createCellStyle();

        if (cell.getColumnIndex()==4 && cell.getRowIndex()==2){
            cellStyleStandart.setVerticalAlignment(DISTRIBUTED);
            cellStyleStandart.setAlignment(HorizontalAlignment.DISTRIBUTED);
            //cellStyleStandart.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            //cellStyleStandart.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        }
        else cellStyleStandart.setAlignment(HorizontalAlignment.CENTER);

        if((cell.getColumnIndex() ==0 || cell.getColumnIndex()==2) && (cell.getRowIndex()==1)) {
            //no right border
           cellStyleStandart.setBorderBottom(BorderStyle.THIN);
            //cellStyleStandart.setBorderTop(BorderStyle.THIN);
            cellStyleStandart.setBorderLeft(BorderStyle.THIN);

        }
        else if((cell.getColumnIndex()==1 || cell.getColumnIndex()==3) && (cell.getRowIndex()==1)){
            //no left border
            cellStyleStandart.setBorderBottom(BorderStyle.THIN);
            //cellStyleStandart.setBorderTop(BorderStyle.THIN);
            cellStyleStandart.setBorderRight(BorderStyle.THIN);

        }
        else {
            cellStyleStandart.setBorderBottom(BorderStyle.THIN);

            //cellStyleStandart.setBorderTop(BorderStyle.THIN);
            cellStyleStandart.setBorderLeft(BorderStyle.THIN);
            cellStyleStandart.setBorderRight(BorderStyle.THIN);
        }
        return cellStyleStandart;
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
