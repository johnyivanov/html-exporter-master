package de.dfs.html.converter.writer;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

class RowTrackerTest {

    @Test
    void testConstructor() {
        assertEquals(0, (new RowTracker()).getTrackedRowCount());
    }

       @Test
    void testAddCell() {
        RowTracker rowTracker = new RowTracker();
        rowTracker.addCell(1, 1, 2, 2);
        assertEquals(2, rowTracker.getTrackedRowCount());
    }


    @Test
    void testUpdateRows() {
        RowTracker rowTracker = new RowTracker();
        rowTracker.updateRows(1, 1, 1, 1);
        assertEquals(1, rowTracker.getTrackedRowCount());
    }

    @Test
    void testGetNextColumnIndexForRow() {
        assertEquals(0, (new RowTracker()).getNextColumnIndexForRow(1));
    }
    @Test
    void testGetNextColumnIndexForRow2() {
        RowTracker rowTracker = new RowTracker();
        rowTracker.addCell(1, 1, 2, 2);
        assertEquals(0, rowTracker.getNextColumnIndexForRow(1));
    }

    @Test
    void testGetNextColumnIndexForRow3() {
        RowTracker rowTracker = new RowTracker();
        rowTracker.addCell(1, 0, 2, 2);
        assertEquals(2, rowTracker.getNextColumnIndexForRow(1));
    }

    @Test
    void testToString() {
        assertEquals("", (new RowTracker()).toString());
    }

        @Test
    void testToString2() {
        RowTracker rowTracker = new RowTracker();
        rowTracker.addCell(1, 1, 2, 2);
        assertEquals("Row 1:[x,x][1,1][1,2]\nRow 2:[x,x][2,1][2,2]\n", rowTracker.toString());
    }
}

