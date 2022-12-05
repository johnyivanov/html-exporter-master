package de.dfs.html.converter.ss;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.any;
import static org.mockito.Mockito.atLeast;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

class CellRangeTest {

    @Test
    void testConstructor() {
        CellRange actualCellRange = new CellRange();
        assertEquals(0, actualCellRange.getHeight());
        assertTrue(actualCellRange.isEmpty());
        assertTrue(actualCellRange.isContiguous());
        assertTrue(actualCellRange.getRows().isEmpty());
    }


    @Test
    void testAddCell() {
        CellRange cellRange = new CellRange();
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        assertFalse(cellRange.isEmpty());
        assertEquals(1, cellRange.getRows().size());
    }




    @Test
    void testAddCell2() {
        CellRange cellRange = new CellRange();
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        assertFalse(cellRange.isEmpty());
    }


    @Test
    void testAddCell3() {
        CellRangeObserver cellRangeObserver = mock(CellRangeObserver.class);
        doNothing().when(cellRangeObserver).cellRangeUpdated((CellRange) any());

        CellRange cellRange = new CellRange();
        cellRange.addCellRangeObserver(cellRangeObserver);
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        verify(cellRangeObserver, atLeast(1)).cellRangeUpdated((CellRange) any());
        assertFalse(cellRange.isEmpty());
    }


    @Test
    void testAddCell4() {
        CellRangeObserver cellRangeObserver = mock(CellRangeObserver.class);
        doNothing().when(cellRangeObserver).cellRangeUpdated((CellRange) any());

        CellRange cellRange = new CellRange();
        cellRange.addCellRangeObserver(cellRangeObserver);
        cellRange.addCell(new DefaultTableCellReference(2, 1));
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        verify(cellRangeObserver, atLeast(1)).cellRangeUpdated((CellRange) any());
        assertFalse(cellRange.isEmpty());
    }


    @Test
    void testAddCell5() {
        CellRangeObserver cellRangeObserver = mock(CellRangeObserver.class);
        doNothing().when(cellRangeObserver).cellRangeUpdated((CellRange) any());

        CellRange cellRange = new CellRange();
        cellRange.addCellRangeObserver(cellRangeObserver);
        cellRange.addCell(new DefaultTableCellReference(0, 1));
        cellRange.addCell(new DefaultTableCellReference(1, 1));
        verify(cellRangeObserver, atLeast(1)).cellRangeUpdated((CellRange) any());
        assertTrue(cellRange.isContiguous());
        assertEquals(2, cellRange.getHeight());
    }


}

