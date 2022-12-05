package de.dfs.html.converter.writer;

import de.dfs.html.converter.ss.Dimension;
import de.dfs.html.converter.writer.excel.ExcelTableCellWriter;
import org.jsoup.nodes.Element;
import org.jsoup.parser.Tag;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.mockito.Mock;
import org.mockito.Mockito;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class AbstractTableCellWriterTest {



    @Test
    void testSpansMultipleColumns() {
        ExcelTableCellWriter excelTableCellWriter = new ExcelTableCellWriter(null);
        assertFalse(excelTableCellWriter.spansMultipleColumns(new Element("Tag")));
    }

    @Test
    void testSpansMultipleColumns3() {
        ExcelTableCellWriter excelTableCellWriter = new ExcelTableCellWriter(null);
        assertFalse(excelTableCellWriter
                .spansMultipleColumns(new Element(mock(Tag.class), TableCellWriter.COLUMN_SPAN_ATTRIBUTE)));
    }

    @Test
    void testGetMergedColumnCount() {
        ExcelTableCellWriter excelTableCellWriter = new ExcelTableCellWriter(null);
        Element element = new Element("tag");


        int actualMergedColumnCount = excelTableCellWriter.getMergedColumnCount(element);
        assertEquals(1, actualMergedColumnCount);

    }

    @Test
    @Disabled
    void testGetMergedColumnCount2() {

        ExcelTableCellWriter excelTableCellWriter = new ExcelTableCellWriter(null);
        Element element = new Element("tag");
        element.attr("2");

        // Act

        int actualMergedColumnCount = excelTableCellWriter.getMergedColumnCount(element);
       ExcelTableCellWriter myMock = Mockito.mock(ExcelTableCellWriter.class);
        ExcelTableCellWriter mySpy = Mockito.spy(excelTableCellWriter);

        // Assert

        when(myMock.spansMultipleColumns(element)).thenReturn(true);
        //doNothing().when(mySpy.spansMultipleColumns(element));

        assertEquals(2, actualMergedColumnCount);

    }

}

