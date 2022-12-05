package de.dfs.html.converter.writer;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;

import de.dfs.html.converter.writer.excel.ExcelExporter;

import java.io.File;

import java.io.IOException;
import java.nio.file.Paths;

import org.jsoup.nodes.Element;
import org.jsoup.parser.Tag;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

class AbstractExporterTest {



    @Test
    void testGetTables() {
        assertTrue((new ExcelExporter()).getTables("Html").isEmpty());
    }

    @Test
    void testIsNewSheet() {
        ExcelExporter excelExporter = new ExcelExporter();
        assertFalse(excelExporter.isNewSheet(new Element("Tag")));
    }

    @Test
    void testIsNewSheet2() {
        ExcelExporter excelExporter = new ExcelExporter();
        assertFalse(excelExporter.isNewSheet(new Element(mock(Tag.class), "data-new-sheet")));
    }

    @Test
    void testGetSheetName() {
        ExcelExporter excelExporter = new ExcelExporter();
        assertEquals("", excelExporter.getSheetName(new Element("Tag")));
    }

    @Test
    void testGetSheetName3() {
        ExcelExporter excelExporter = new ExcelExporter();
        assertEquals("", excelExporter.getSheetName(new Element(mock(Tag.class), "data-sheet-name")));
    }
}

