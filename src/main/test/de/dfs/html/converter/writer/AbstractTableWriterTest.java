package de.dfs.html.converter.writer;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.any;
import static org.mockito.Mockito.anyInt;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

import de.dfs.html.converter.writer.excel.ExcelTableWriter;
import org.jsoup.nodes.Attributes;
import org.jsoup.nodes.CDataNode;
import org.jsoup.nodes.Comment;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.PseudoTextElement;
import org.jsoup.parser.Tag;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

class AbstractTableWriterTest {

    @Test
    void testWriteTable() {
        ExcelTableWriter excelTableWriter = new ExcelTableWriter(mock(TableRowWriter.class));
        assertEquals(0, excelTableWriter.writeTable(new Element("Tag"), "Style Mapper", 1));
    }


   /* @Test
    @Disabled
    void testWriteTable2() {
        TableRowWriter tableRowWriter = mock(TableRowWriter.class);
        doNothing().when(tableRowWriter).writeRow((Element) any(), anyInt());
        ExcelTableWriter excelTableWriter = new ExcelTableWriter(tableRowWriter);
        assertEquals(1, excelTableWriter.writeTable(new Element(TableWriter.TABLE_ROW_ELEMENT_NAME), "Style Mapper", 1));
        verify(tableRowWriter).writeRow((Element) any(), anyInt());
    }

        @Test
        @Disabled
    void testWriteTable3() {
        TableRowWriter tableRowWriter = mock(TableRowWriter.class);
        doNothing().when(tableRowWriter).writeRow((Element) any(), anyInt());
        ExcelTableWriter excelTableWriter = new ExcelTableWriter(tableRowWriter);
        assertEquals(0,
                excelTableWriter.writeTable(Document.createShell(TableWriter.TABLE_ROW_ELEMENT_NAME), "Style Mapper", 1));
    }


    @Test
    @Disabled
    void testWriteTable4() {
        TableRowWriter tableRowWriter = mock(TableRowWriter.class);
        doNothing().when(tableRowWriter).writeRow((Element) any(), anyInt());
        ExcelTableWriter excelTableWriter = new ExcelTableWriter(tableRowWriter);
        Tag tag = mock(Tag.class);
        when(tag.normalName()).thenReturn("Normal Name");
        assertEquals(0,
                excelTableWriter.writeTable(new Element(tag, TableWriter.TABLE_ROW_ELEMENT_NAME), "Style Mapper", 1));
        verify(tag).normalName();
    }*/

    @Test
    void testWriteTable5() {
        TableRowWriter tableRowWriter = mock(TableRowWriter.class);
        doNothing().when(tableRowWriter).writeRow((Element) any(), anyInt());
        ExcelTableWriter excelTableWriter = new ExcelTableWriter(tableRowWriter);
        Tag tag = mock(Tag.class);
        when(tag.normalName()).thenReturn("Normal Name");

        PseudoTextElement pseudoTextElement = new PseudoTextElement(tag, "Base Uri", new Attributes());
        pseudoTextElement.appendChild(new CDataNode("Text"));
        assertEquals(0, excelTableWriter.writeTable(pseudoTextElement, "Style Mapper", 1));
        verify(tag).normalName();
    }


    @Test
    void testWriteTable6() {
        TableRowWriter tableRowWriter = mock(TableRowWriter.class);
        doNothing().when(tableRowWriter).writeRow((Element) any(), anyInt());
        ExcelTableWriter excelTableWriter = new ExcelTableWriter(tableRowWriter);
        Tag tag = mock(Tag.class);
        when(tag.normalName()).thenReturn("Normal Name");

        Element element = new Element(tag, "Base Uri", new Attributes());
        element.appendChild(new Comment("Data"));
        assertEquals(0, excelTableWriter.writeTable(element, "Style Mapper", 1));
        verify(tag).normalName();
    }
}

