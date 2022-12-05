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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.jsoup.nodes.Element;
import de.dfs.html.converter.ss.Function;
import de.dfs.html.converter.writer.AbstractTableCellWriter;

public class ExcelTableCellWriter extends AbstractTableCellWriter {

    private Sheet sheet;
    private ExcelStyleGenerator styleGenerator;

    public ExcelTableCellWriter (Sheet sheet){
        this.sheet = sheet;


        styleGenerator = new ExcelStyleGenerator();
    }

    @Override
    public void renderCell(Element element, int rowIndex, int columnIndex) {

        int cellHight = element.wholeText().length() - element.wholeText().replace(System.lineSeparator(), "").length();
        if(cellHight>0){
            sheet.getRow(rowIndex).setHeightInPoints(((cellHight-2)*sheet.getDefaultRowHeightInPoints()));

        }
        Cell cell = sheet.getRow(rowIndex).createCell(columnIndex, CellType.STRING);
        cell.setCellValue(getElementText(element));


        cell.setCellStyle(styleGenerator.getStyle(cell, null));

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

    public void addFunctionCell(int rowIndex, int columnIndex, CellRange range, Function function) {
        Cell cell = sheet.getRow(rowIndex).getCell(columnIndex);

        new ExcelFunctionCell(cell, range, new ExcelCellRangeResolver(), function);
    }
}
