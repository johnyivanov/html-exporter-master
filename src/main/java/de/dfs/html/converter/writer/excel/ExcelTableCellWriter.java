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

import de.dfs.html.converter.ss.CellRange;
import de.dfs.html.converter.ss.Function;
import de.dfs.html.converter.writer.AbstractTableCellWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.jsoup.nodes.Element;

import static de.dfs.html.converter.writer.TableRowWriter.*;
import static de.dfs.html.converter.writer.TableWriter.TD_TABLE;

public class ExcelTableCellWriter extends AbstractTableCellWriter {

    private Sheet sheet;
    private ExcelStyleGenerator styleGenerator;


    public ExcelTableCellWriter(Sheet sheet) {
        this.sheet = sheet;


        styleGenerator = new ExcelStyleGenerator();
    }

    @Override
    public void renderCell(Element element, int rowIndex, int columnIndex) {
        if (element.getElementsByTag(TD_TABLE).size() == 1 ) {
            renderInnerTable(element, rowIndex,columnIndex);
            return;
        }





        if (sheet.getRow(rowIndex) == null) {
            sheet.createRow(rowIndex);
        }
        Cell cell = sheet.getRow(rowIndex).createCell(columnIndex, CellType.STRING);
        cell.setCellValue(getElementText(element));

        /*int cellHight = element.wholeText().length() - element.wholeText().replace(System.lineSeparator(), "").length();
        if (cellHight > 0) {
            sheet.getRow(rowIndex).setHeightInPoints(((cellHight - 2) * sheet.getDefaultRowHeightInPoints()));
        }*/

     sheet.getRow(rowIndex).setHeightInPoints(50);
     sheet.setColumnWidth(columnIndex,4000);

       // sheet.autoSizeColumn(columnIndex);
        if(element.tag().getName().equals(TH_TAG)) {
            cell.setCellStyle(styleGenerator.getCellHeaderStyle(sheet.getWorkbook(),cell,sheet));
        } else cell.setCellStyle(styleGenerator.getCellStandartStyle(sheet.getWorkbook(),cell,sheet));

        if (isDateCell(element)) {
            CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
            cell.getCellStyle().setDataFormat(createHelper.createDataFormat().getFormat(getDateCellFormat(element)));
        }

        String commentText;

        if ((commentText = getCellCommentText(element)) != null) {
            ExcelCellCommentGenerator.addCellComment(cell, commentText, getCellCommentDimension(element));
        }

        if (definesFreezePane(element)) {
            sheet.createFreezePane(columnIndex, rowIndex);
        }
    }
    private void renderInnerTable(Element table, int rowIndex, int columnIndex) {

        for (Element tableRow : table.getElementsByTag(TABLE_ROW_ELEMENT_NAME)) {

            for (Element element : tableRow.getAllElements()) {

                if (element.tag().getName().equals(TD_TAG) ) {
                    //columnIndex = rowTracker.getNextColumnIndexForRow(rowIndex);
                    writeCell(element, rowIndex, columnIndex,false);
                    ExcelExporter.lastInnerColumnIndex = columnIndex;
                    columnIndex++;
                }
            }
            ExcelExporter.globalColumnIndex = columnIndex;
            columnIndex = 0;
            ++rowIndex;
            ExcelExporter.lastInnerRowIndex = rowIndex;
        }

    }

    public void addFunctionCell(int rowIndex, int columnIndex, CellRange range, Function function) {
        Cell cell = sheet.getRow(rowIndex).getCell(columnIndex);

        new ExcelFunctionCell(cell, range, new ExcelCellRangeResolver(), function);
    }


}