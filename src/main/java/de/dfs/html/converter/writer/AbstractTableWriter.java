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
package de.dfs.html.converter.writer;

import org.jsoup.nodes.Element;

import java.util.List;

import static de.dfs.html.converter.writer.TableRowWriter.TABLE_BODY;
import static de.dfs.html.converter.writer.TableRowWriter.TABLE_ROW_ELEMENT_NAME;

public abstract class AbstractTableWriter implements TableWriter {
    private TableRowWriter rowWriter;

    public AbstractTableWriter(TableRowWriter rowWriter) {
        this.rowWriter = rowWriter;
    }

    public int writeTable(Element table, Object styleMapper, int startRow) {
        renderTable(table);

        int rowIndex = startRow;

        //List<Element> tbodyElement = table.getElementsByTag(TABLE_BODY);

        for (Element element : table.children()) {
            for (Element element1 : element.children()) {
                if (element1.tag().getName().equals(TABLE_ROW_ELEMENT_NAME)) {

                    rowWriter.writeRow(element1, rowIndex);
                    ++rowIndex;
                }
            }
        }
        return rowIndex - startRow;
    }

    public abstract void renderTable(Element table);
}
