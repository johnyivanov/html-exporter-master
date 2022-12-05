package de.dfs.html.converter.ss;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

class CellRangeRowTest {

    @Test
    void testConstructor() {
        assertEquals(1, (new CellRangeRow(1)).getIndex());
    }


    @Test
    void testAddCell() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertFalse(cellRangeRow.isEmpty());
        assertEquals(2, cellRangeRow.getCells().size());
    }

    @Test
    void testAddCell2() {
        assertThrows(IllegalArgumentException.class, () -> (new CellRangeRow(1)).addCell(null));
    }


    @Test
    void testAddCell3() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertFalse(cellRangeRow.isEmpty());
        assertEquals(2, cellRangeRow.getCells().size());
    }

    @Test
    void testIsContiguous() {
        assertFalse((new CellRangeRow(1)).isContiguous());
    }


    @Test
    void testIsContiguous2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertTrue(cellRangeRow.isContiguous());
    }


    @Test
    void testIsContiguous3() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 0));
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertTrue(cellRangeRow.isContiguous());
    }


    @Test
    void testIsContiguous4() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 0));
        cellRangeRow.addCell(new DefaultTableCellReference(1, 2));
        assertFalse(cellRangeRow.isContiguous());
    }


    @Test
    void testGetFirstPopulatedColumn() {
        assertEquals(-1, (new CellRangeRow(1)).getFirstPopulatedColumn());
    }


    @Test
    void testGetFirstPopulatedColumn2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertEquals(1, cellRangeRow.getFirstPopulatedColumn());
    }

    @Test
    void testGetLastPopulatedColumn() {
        assertEquals(-1, (new CellRangeRow(1)).getLastPopulatedColumn());
    }

    @Test
    void testGetLastPopulatedColumn2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertEquals(1, cellRangeRow.getLastPopulatedColumn());
    }

    @Test
    void testIsEmpty() {
        assertTrue((new CellRangeRow(1)).isEmpty());
    }

    @Test
    void testIsEmpty2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertFalse(cellRangeRow.isEmpty());
    }

    @Test
    void testGetWidth() {
        assertEquals(0, (new CellRangeRow(1)).getWidth());
    }

    @Test
    void testGetWidth2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertEquals(1, cellRangeRow.getWidth());
    }

    @Test
    void testGetFirstCell2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        DefaultTableCellReference defaultTableCellReference = new DefaultTableCellReference(1, 1);

        cellRangeRow.addCell(defaultTableCellReference);
        assertSame(defaultTableCellReference, cellRangeRow.getFirstCell());
    }

    @Test
    void testGetLastCell2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        DefaultTableCellReference defaultTableCellReference = new DefaultTableCellReference(1, 1);

        cellRangeRow.addCell(defaultTableCellReference);
        assertSame(defaultTableCellReference, cellRangeRow.getLastCell());
    }


    @Test
    void testToString() {
        assertEquals("", (new CellRangeRow(1)).toString());
    }

    @Test
    void testToString2() {
        CellRangeRow cellRangeRow = new CellRangeRow(1);
        cellRangeRow.addCell(new DefaultTableCellReference(1, 1));
        assertEquals("[x][1,1]", cellRangeRow.toString());
    }
}

