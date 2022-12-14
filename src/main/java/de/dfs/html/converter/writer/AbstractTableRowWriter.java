/**
 * Copyright (C) 2012 alanhay <alanhay99@hotmail.com>
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *         http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package de.dfs.html.converter.writer;

import de.dfs.html.converter.writer.excel.ExcelExporter;
import de.dfs.html.converter.writer.excel.ExcelTableWriter;
import org.jsoup.nodes.Element;

public abstract class AbstractTableRowWriter implements TableRowWriter {
    private TableCellWriter cellWriter;
    private RowTracker rowTracker;

    public AbstractTableRowWriter(TableCellWriter cellWriter) {
        this.cellWriter = cellWriter;
        rowTracker = new RowTracker();
    }

    public void writeRow(Element row, int rowIndex) {
        renderRow(row, rowIndex);
        if (ExcelExporter.lastInnerRowIndex >0  && rowIndex < ExcelExporter.lastInnerRowIndex ) {

            ExcelExporter.globalColumnIndex = ExcelExporter.lastInnerColumnIndex + 1;
        } else{
            ExcelExporter.globalColumnIndex = 0;
        }
        for (Element element : row.children()) {
            if (element.tag().getName().equals(TD_TAG) || element.tag().getName().equals(TH_TAG)) {

                //ExcelExporter.globalColumnIndex = rowTracker.getNextColumnIndexForRow(rowIndex);
                cellWriter.writeCell(element, rowIndex, ExcelExporter.globalColumnIndex,true);
                int rowSpan = 1;
                int columnSpan = getColumnSpan(element);

                rowTracker.addCell(rowIndex, ExcelExporter.globalColumnIndex, rowSpan, columnSpan);

                if (rowSpan > 1 || columnSpan > 1) {
                    doMerge(rowIndex, ExcelExporter.globalColumnIndex - 1, rowSpan, columnSpan);
                }
            }
        }
    }

    protected boolean isRowGrouped(Element row) {
        return row.hasAttr("");
    }

    protected String[] getRowGroups(Element row) {
        return getAttributeValues(row, "");
    }

    protected String[] getAttributeValues(Element element, String attributeName) {
        String values[] = null;

        if (element.hasAttr(attributeName)) {
            values = element.attr(attributeName).toLowerCase().split(",");

            for (String value : values) {
                value = value.trim().toLowerCase();
            }
        }

        return values;
    }

    protected int getColumnSpan(Element element) {
        int columnSpan = 1;

        if (element.hasAttr("colspan")) {
            columnSpan = Integer.parseInt(element.attr("colspan"));
        }

        return columnSpan;
    }

    protected int getRowSpan(Element element) {
        int rowSpan = 1;

        if (element.hasAttr("rowSpan")) {
            rowSpan = Integer.parseInt(element.attr("rowSpan"));
        }

        return rowSpan;
    }

    public abstract void renderRow(Element row, int rowIndex);

    public abstract void doMerge(int rowIndex, int columnIndex, int rowSpan, int columnSpan);
}