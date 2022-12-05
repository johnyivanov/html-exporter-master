package de.dfs.html.converter.ss;

import static org.junit.jupiter.api.Assertions.assertNull;

import org.junit.jupiter.api.Test;

class RangeReferenceTrackerTest {

    @Test
    void testGetCellRange() {
        assertNull((new RangeReferenceTracker()).getCellRange("Name"));
    }
}

