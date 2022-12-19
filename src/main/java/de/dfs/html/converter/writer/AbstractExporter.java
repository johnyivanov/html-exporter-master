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

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.*;
import java.util.List;
import java.util.stream.Collectors;

public abstract class AbstractExporter implements Exporter {

    private static final String DATA_NEW_SHEET_ATTRIBUTE = "data-new-sheet";
    private static final String DATA_SHEET_NAME_ATTRIBUTE = "data-sheet-name";

    public byte[] exportHtml(String html) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        exportHtml(html, out);

        return out.toByteArray();
    }

    public void exportHtml(String html, File outputFile) throws IOException {
        FileOutputStream out = new FileOutputStream(outputFile);
        exportHtml(html, out);
    }

    protected List<Element> getTables(String html) {
        Document document = Jsoup.parse(html);// FIXME parsing twice

        return document.getElementsByTag("body").get(0).children().stream().filter(element -> element.tag().getName().equals("table")).collect(Collectors.toList());
    }


    protected boolean isNewSheet(Element element) {
        return Boolean.valueOf(element.attr(DATA_NEW_SHEET_ATTRIBUTE));
    }

    protected String getSheetName(Element element) {
        return element.attr(DATA_SHEET_NAME_ATTRIBUTE);
    }

    public abstract void exportHtml(String html, OutputStream out) throws IOException;
}
