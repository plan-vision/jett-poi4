package net.sf.jett.test;

import java.util.List;

import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.exception.FormulaParseException;
import net.sf.jett.formula.CellRef;
import net.sf.jett.parser.FormulaParser;

/**
 * This JUnit Test class tests the <code>FormulaParser</code>.
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class FormulaParserTest
{
    /**
     * Simple test.
     */
    @Test
    public void testSimple()
    {
        String formula = "SUM(B2)";
        FormulaParser parser = new FormulaParser(formula);
        parser.parse();

        List<CellRef> cellRefs = parser.getCellReferences();
        assertEquals(1, cellRefs.size());
        CellRef cellRef = cellRefs.get(0);
        assertEquals("B2", cellRef.formatAsString());
        assertEquals("B2", cellRef.formatAsStringWithDef());
        assertNull(cellRef.getDefaultValue());
    }

    /**
     * Tests default value.  Negative default values are allowed.
     */
    @Test
    public void testDefault()
    {
        String formula = "IF(M11||-1 = 0, M12||$Z$1, M13)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();

        List<CellRef> cellRefs = parser.getCellReferences();
        assertEquals(3, cellRefs.size());

        CellRef cellRef0 = cellRefs.get(0);
        assertEquals("M11", cellRef0.formatAsString());
        assertEquals("M11||-1", cellRef0.formatAsStringWithDef());
        assertEquals("-1", cellRef0.getDefaultValue());

        CellRef cellRef1 = cellRefs.get(1);
        assertEquals("M12", cellRef1.formatAsString());
        assertEquals("M12||$Z$1", cellRef1.formatAsStringWithDef());
        assertEquals("$Z$1", cellRef1.getDefaultValue());

        CellRef cellRef2 = cellRefs.get(2);
        assertEquals("M13", cellRef2.formatAsString());
        assertEquals("M13", cellRef2.formatAsStringWithDef());
        assertNull(cellRef2.getDefaultValue());
    }

    /**
     * If there is a sheet delimiter ("!"), but no sheet name, ensure that a
     * <code>FormulaParseException</code> is thrown.
     */
    @Test(expected = FormulaParseException.class)
    public void testNoSheetName()
    {
        String formula = "SUM(!B2)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();
    }

    /**
     * If there is a sheet delimiter ("!") while expecting a default value,
     * ensure that a <code>FormulaParseException</code> is thrown.
     */
    @Test(expected = FormulaParseException.class)
    public void testSheetDelimiterWhileExpectingDefaultValue()
    {
        String formula = "COUNTA(B2||!$Z$1)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();
    }

    /**
     * If there is another default value indicator ("||") while already
     * expecting a default value, ensure that a
     * <code>FormulaParseException</code> is thrown.
     */
    @Test(expected = FormulaParseException.class)
    public void testDefaultValueDelimiterWhileExpectingDefaultValue()
    {
        String formula = "COUNTA(B2||-||$Z$1)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();
    }

    /**
     * If there is a default value indicator ("||") without a cell reference,
     * ensure that a <code>FormulaParseException</code> is thrown.
     */
    @Test(expected = FormulaParseException.class)
    public void testDefaultValueWithoutCellReference()
    {
        String formula = "SUM(||$Z$1)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();
    }

    /**
     * Test what happens when there are no actual Excel formulas, and the cell
     * reference is actually last in the string.
     */
    @Test
    public void testNoFormulaOnlyOperators()
    {
        String formula = "B2 + C2";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();

        List<CellRef> cellRefs = parser.getCellReferences();
        assertEquals(2, cellRefs.size());

        CellRef cellRef0 = cellRefs.get(0);
        assertEquals("B2", cellRef0.formatAsString());

        CellRef cellRef1 = cellRefs.get(1);
        assertEquals("C2", cellRef1.formatAsString());
    }

    /**
     * Ensure that the list of cell references does not contain any duplicates.
     */
    @Test
    public void testDuplicateCellReferences()
    {
        String formula = "AB123 * (B1 + B1) * AB123";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();

        List<CellRef> cellRefs = parser.getCellReferences();
        assertEquals(2, cellRefs.size());

        CellRef cellRef0 = cellRefs.get(0);
        assertEquals("AB123", cellRef0.formatAsString());

        CellRef cellRef1 = cellRefs.get(1);
        assertEquals("B1", cellRef1.formatAsString());
    }

    /**
     * Ensure that sheet names with spaces are parsed properly.
     */
    @Test
    public void testSheetNameWithSpace()
    {
        String formula = "COUNTA('Sheet Name With Spaces'!F5||$Z$1)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();

        List<CellRef> cellRefs = parser.getCellReferences();
        assertEquals(1, cellRefs.size());

        CellRef cellRef0 = cellRefs.get(0);
        assertEquals("'Sheet Name With Spaces'!F5", cellRef0.formatAsString());
        assertEquals("'Sheet Name With Spaces'!F5||$Z$1", cellRef0.formatAsStringWithDef());
        assertEquals("$Z$1", cellRef0.getDefaultValue());
        assertEquals("Sheet Name With Spaces", cellRef0.getSheetName());
    }

    /**
     * Ensure that formulas that have a colon still pick up the distinct cell
     * references.
     * @since 0.8.0
     */
    @Test
    public void testFormulaWithColon()
    {
        String formula = "SUM(A9:A10)";

        FormulaParser parser = new FormulaParser(formula);
        parser.parse();

        List<CellRef> cellRefs = parser.getCellReferences();
        assertEquals(2, cellRefs.size());

        CellRef cellRef0 = cellRefs.get(0);
        assertEquals("A9", cellRef0.formatAsString());

        CellRef cellRef1 = cellRefs.get(1);
        assertEquals("A10", cellRef1.formatAsString());
    }
}
