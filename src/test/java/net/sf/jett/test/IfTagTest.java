package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "if" tag in entire rows,
 * block area, and bodiless modes.
 *
 * @author Randy Gettman
 */
public class IfTagTest extends TestCase
{
    /**
     * Tests the .xls template spreadsheet.
     * @throws java.io.IOException If an I/O error occurs.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the input spreadsheet is invalid.
     */
    @Override
    @Test
    public void testXls() throws IOException, InvalidFormatException
    {
        super.testXls();
    }

    /**
     * Tests the .xlsx template spreadsheet.
     * @throws IOException If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     */
    @Override
    @Test
    public void testXlsx() throws IOException, InvalidFormatException
    {
        super.testXlsx();
    }

    /**
     * Returns the Excel name base for the template and resultant spreadsheets
     * for this test.
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "IfTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet vertical = workbook.getSheetAt(0);
        assertEquals("This", TestUtility.getStringCellValue(vertical, 1, 0));
        assertEquals("is", TestUtility.getStringCellValue(vertical, 1, 1));
        assertEquals("Block1", TestUtility.getStringCellValue(vertical, 1, 2));
        assertEquals("Merged Region", TestUtility.getStringCellValue(vertical, 1, 3));
        assertTrue(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(1, 1, 3, 4)));
        assertEquals("After1", TestUtility.getStringCellValue(vertical, 2, 0));
        assertEquals("After2", TestUtility.getStringCellValue(vertical, 3, 0));
        assertFalse(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(3, 3, 3, 4)));
        assertEquals("After3", TestUtility.getStringCellValue(vertical, 4, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 5, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 5, 1));
        assertTrue(TestUtility.isCellBlank(vertical, 5, 2));
        assertTrue(TestUtility.isCellBlank(vertical, 5, 3));
        assertTrue(TestUtility.isCellBlank(vertical, 5, 4));
        assertFalse(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(5, 5, 3, 4)));
        assertEquals("After4", TestUtility.getStringCellValue(vertical, 6, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 7, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 7, 1));
        assertTrue(TestUtility.isCellBlank(vertical, 7, 2));
        assertTrue(TestUtility.isCellBlank(vertical, 7, 3));
        assertTrue(TestUtility.isCellBlank(vertical, 7, 4));
        assertTrue(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(7, 7, 3, 4)));
        assertEquals("After5", TestUtility.getStringCellValue(vertical, 8, 0));
        assertEquals("After6", TestUtility.getStringCellValue(vertical, 9, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 10, 0));
        assertFalse(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(11, 11, 3, 4)));
        assertTrue(TestUtility.isCellBlank(vertical, 11, 0));
        assertEquals("After7", TestUtility.getStringCellValue(vertical, 12, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 13, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(14, 14, 3, 4)));
        assertTrue(TestUtility.isCellBlank(vertical, 14, 0));
        assertEquals("After8", TestUtility.getStringCellValue(vertical, 15, 0));
        assertTrue(TestUtility.isCellBlank(vertical, 16, 0));
        assertFalse(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(17, 17, 3, 4)));
        assertTrue(TestUtility.isCellBlank(vertical, 17, 0));
        assertEquals("After9", TestUtility.getStringCellValue(vertical, 18, 0));
        assertEquals("Entire", TestUtility.getStringCellValue(vertical, 20, 0));
        assertEquals("Rows", TestUtility.getStringCellValue(vertical, 20, 1));
        assertEquals("Block10", TestUtility.getStringCellValue(vertical, 20, 2));
        assertEquals("Merged Region", TestUtility.getStringCellValue(vertical, 20, 3));
        assertTrue(TestUtility.isMergedRegionPresent(vertical, new CellRangeAddress(20, 20, 3, 4)));
        assertEquals("After10", TestUtility.getStringCellValue(vertical, 21, 0));
        assertEquals(4, vertical.getNumMergedRegions());

        Sheet horizontal = workbook.getSheetAt(1);
        assertEquals("This", TestUtility.getStringCellValue(horizontal, 1, 0));
        assertEquals("is", TestUtility.getStringCellValue(horizontal, 1, 1));
        assertEquals("Block1", TestUtility.getStringCellValue(horizontal, 1, 2));
        assertEquals("Merged Region 1", TestUtility.getStringCellValue(horizontal, 1, 3));
        assertTrue(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(1, 1, 3, 4)));
        assertEquals("!", TestUtility.getStringCellValue(horizontal, 1, 5));
        assertEquals("After1", TestUtility.getStringCellValue(horizontal, 1, 6));
        assertEquals("This", TestUtility.getStringCellValue(horizontal, 2, 0));
        assertEquals("is a", TestUtility.getStringCellValue(horizontal, 2, 1));
        assertEquals("multi-row", TestUtility.getStringCellValue(horizontal, 2, 2));
        assertEquals("Merged Region 2", TestUtility.getStringCellValue(horizontal, 2, 3));
        assertTrue(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(2, 3, 3, 4)));
        assertEquals("Right", TestUtility.getStringCellValue(horizontal, 2, 5));
        assertEquals("After1", TestUtility.getStringCellValue(horizontal, 2, 6));
        assertEquals("block", TestUtility.getStringCellValue(horizontal, 3, 0));
        assertEquals("area", TestUtility.getStringCellValue(horizontal, 3, 1));
        assertEquals("block.", TestUtility.getStringCellValue(horizontal, 3, 2));
        assertEquals("!", TestUtility.getStringCellValue(horizontal, 3, 5));
        assertEquals("After1", TestUtility.getStringCellValue(horizontal, 3, 6));
        assertEquals("After2", TestUtility.getStringCellValue(horizontal, 4, 0));
        assertEquals("After4", TestUtility.getStringCellValue(horizontal, 5, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 5, 3));
        assertEquals("After3", TestUtility.getStringCellValue(horizontal, 5, 6));
        assertFalse(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(5, 5, 3, 4)));
        assertEquals("After6", TestUtility.getStringCellValue(horizontal, 6, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 6, 3));
        assertEquals("After3", TestUtility.getStringCellValue(horizontal, 6, 6));
        assertEquals("After3", TestUtility.getStringCellValue(horizontal, 7, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 8, 0));
        assertEquals("After5", TestUtility.getStringCellValue(horizontal, 9, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 9, 6));
        assertEquals("After8", TestUtility.getStringCellValue(horizontal, 10, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 10, 3));
        assertEquals("After5", TestUtility.getStringCellValue(horizontal, 10, 6));
        assertTrue(TestUtility.isCellBlank(horizontal, 11, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 11, 3));
        assertTrue(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(11, 11, 3, 4)));
        assertTrue(TestUtility.isCellBlank(horizontal, 12, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 12, 3));
        assertTrue(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(12, 13, 3, 4)));
        assertTrue(TestUtility.isCellBlank(horizontal, 13, 0));
        assertEquals("After7", TestUtility.getStringCellValue(horizontal, 13, 6));
        assertEquals("After10", TestUtility.getStringCellValue(horizontal, 14, 0));
        assertEquals("After7", TestUtility.getStringCellValue(horizontal, 14, 6));
        assertTrue(TestUtility.isCellBlank(horizontal, 15, 0));
        assertEquals("Top", TestUtility.getStringCellValue(horizontal, 15, 2));
        assertEquals("After7", TestUtility.getStringCellValue(horizontal, 15, 6));
        assertEquals("Showme1", TestUtility.getStringCellValue(horizontal, 16, 1));
        assertEquals("Showme2", TestUtility.getStringCellValue(horizontal, 16, 2));
        assertEquals("Showme3", TestUtility.getStringCellValue(horizontal, 16, 3));
        assertTrue(TestUtility.isCellBlank(horizontal, 16, 6));
        assertEquals("Left", TestUtility.getStringCellValue(horizontal, 17, 0));
        assertEquals("Showme4", TestUtility.getStringCellValue(horizontal, 17, 1));
        assertEquals("Showme5", TestUtility.getStringCellValue(horizontal, 17, 2));
        assertEquals("Showme6", TestUtility.getStringCellValue(horizontal, 17, 3));
        assertEquals("Right", TestUtility.getStringCellValue(horizontal, 17, 4));
        assertEquals("After9", TestUtility.getStringCellValue(horizontal, 17, 6));
        assertEquals("Showme7", TestUtility.getStringCellValue(horizontal, 18, 1));
        assertEquals("Showme8", TestUtility.getStringCellValue(horizontal, 18, 2));
        assertEquals("Showme9", TestUtility.getStringCellValue(horizontal, 18, 3));
        assertEquals("After9", TestUtility.getStringCellValue(horizontal, 18, 6));
        assertEquals("Bottom", TestUtility.getStringCellValue(horizontal, 19, 2));
        assertEquals("Corner", TestUtility.getStringCellValue(horizontal, 19, 4));
        assertEquals("After9", TestUtility.getStringCellValue(horizontal, 19, 6));
        assertEquals("After11", TestUtility.getStringCellValue(horizontal, 20, 0));
        assertEquals("Bottom", TestUtility.getStringCellValue(horizontal, 21, 1));
        assertFalse(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(21, 22, 1, 2)));
        assertTrue(TestUtility.isCellBlank(horizontal, 21, 3));
        assertEquals("After12", TestUtility.getStringCellValue(horizontal, 22, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 23, 1));
        assertTrue(TestUtility.isCellBlank(horizontal, 23, 3));
        assertTrue(TestUtility.isCellBlank(horizontal, 24, 0));
        assertFalse(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(23, 24, 1, 2)));
        assertEquals("Bottom", TestUtility.getStringCellValue(horizontal, 25, 1));
        assertEquals("After13", TestUtility.getStringCellValue(horizontal, 26, 0));
        assertTrue(TestUtility.isCellBlank(horizontal, 27, 1));
        assertTrue(TestUtility.isMergedRegionPresent(horizontal, new CellRangeAddress(27, 28, 1, 2)));
        assertTrue(TestUtility.isCellBlank(horizontal, 27, 3));
        assertEquals("Bottom", TestUtility.getStringCellValue(horizontal, 29, 1));
        assertEquals("After14", TestUtility.getStringCellValue(horizontal, 30, 0));
        assertEquals(5, horizontal.getNumMergedRegions());

        Sheet bodiless = workbook.getSheetAt(2);
        assertEquals("I'm true!", TestUtility.getStringCellValue(bodiless, 1, 0));
        assertEquals("Right1", TestUtility.getStringCellValue(bodiless, 1, 1));
        assertEquals("I'm false!", TestUtility.getStringCellValue(bodiless, 2, 0));
        assertEquals("Right2", TestUtility.getStringCellValue(bodiless, 2, 1));
        assertEquals("I'm true!", TestUtility.getStringCellValue(bodiless, 3, 0));
        assertEquals("Right3", TestUtility.getStringCellValue(bodiless, 3, 1));
        assertTrue(TestUtility.isCellBlank(bodiless, 4, 0));
        assertEquals("Right4", TestUtility.getStringCellValue(bodiless, 4, 1));
        assertEquals("After", TestUtility.getStringCellValue(bodiless, 5, 0));

        Sheet rts = workbook.getSheetAt(3);
        assertEquals(Math.PI, TestUtility.getNumericCellValue(rts, 0, 0), DELTA);
        Cell cell = TestUtility.getCell(rts, 0, 0);
        assertNotNull(cell);
        assertEquals("008000", TestUtility.getFontColorString(workbook,
                workbook.getFontAt(cell.getCellStyle().getFontIndex())));
        assertEquals(Math.PI, TestUtility.getNumericCellValue(rts, 1, 0), DELTA);
        cell = TestUtility.getCell(rts, 1, 0);
        assertNotNull(cell);
        assertEquals("ff0000", TestUtility.getFontColorString(workbook,
                workbook.getFontAt(cell.getCellStyle().getFontIndex())));
        cell = TestUtility.getCell(rts, 2, 0);
        assertNotNull(cell);
        CellStyle cs = TestUtility.getCellStyle(rts, 2, 0);
        assertNotNull(cs);
        Font f = workbook.getFontAt(cs.getFontIndex());
        //assertEquals(FontBoldweight.BOLD.getIndex(), f.getBold());
        assertEquals("000000", TestUtility.getFontColorString(workbook,
                workbook.getFontAt(cell.getCellStyle().getFontIndex())));
        assertEquals(6, TestUtility.getNumericCellValue(rts, 2, 0), DELTA);
    }

    /**
     * This test is a single map test.
     * @return <code>false</code>.
     */
    protected boolean isMultipleBeans()
    {
        return false;
    }

    /**
     * For single beans map tests, return the <code>Map</code> of bean names to
     * bean values.
     * @return A <code>Map</code> of bean names to bean values.
     */
    protected Map<String, Object> getBeansMap()
    {
        Map<String, Object> beans = new HashMap<>();
        beans.put("condTrue", true);
        beans.put("condFalse", false);
        beans.put("pi", Math.PI);
        beans.put("num", 6);

        return beans;
    }
}