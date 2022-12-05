package de.dfs.html.converter.writer;

import de.dfs.html.converter.writer.excel.ExcelTableRowWriter;
import org.jsoup.nodes.Element;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.mockito.Mockito;


import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.when;

class AbstractTableRowWriterTest {


@Test
    void testGetColumnSpan() {

        AbstractTableRowWriter abstractTableRowWriter = new ExcelTableRowWriter(null,null);
        Element element = new Element("tag");

        int actualColumnSpan = abstractTableRowWriter.getColumnSpan(element);


       assertEquals(1,actualColumnSpan);
    }


    @Test

    void testGetRowSpan() {

        AbstractTableRowWriter abstractTableRowWriter = new ExcelTableRowWriter(null,null);
        Element element = new Element("tag");

        // Act
        int actualRowSpan = abstractTableRowWriter.getRowSpan(element);

        assertEquals(1,actualRowSpan);
    }
}

