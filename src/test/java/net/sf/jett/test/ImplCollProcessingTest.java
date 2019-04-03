package net.sf.jett.test;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the implicit collections processing feature of
 * JETT.
 *
 * @author Randy Gettman
 */
public class ImplCollProcessingTest extends TestCase
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
        return "ImplCollProcessing";
    }

    /**
     * Call certain setup-related methods on the <code>ExcelTransformer</code>
     * before template sheet transformation.
     * @param transformer The <code>ExcelTransformer</code> that will transform
     *    the template worksheet(s).
     */
    @Override
    protected void setupTransformer(ExcelTransformer transformer)
    {
        transformer.turnOffImplicitCollectionProcessing("counties");
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet implicit = workbook.getSheetAt(0);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(implicit, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(implicit, new CellRangeAddress(0, 0, 0, 4)));
        assertEquals("City", TestUtility.getStringCellValue(implicit, 1, 0));
        assertEquals("Celtics", TestUtility.getStringCellValue(implicit, 2, 1));
        assertEquals(37, TestUtility.getNumericCellValue(implicit, 3, 2), DELTA);
        assertEquals(38, TestUtility.getNumericCellValue(implicit, 4, 3), DELTA);
        assertEquals((double) 23 / (23 + 49), TestUtility.getNumericCellValue(implicit, 5, 4), DELTA);
        assertEquals("Toronto", TestUtility.getStringCellValue(implicit, 6, 0));
        assertEquals("Division: Central", TestUtility.getStringCellValue(implicit, 7, 0));
        assertEquals("Lakers", TestUtility.getStringCellValue(implicit, 30, 1));
        assertEquals("Division: Empty",TestUtility. getStringCellValue(implicit, 42, 0));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(implicit, 44, 0));
        assertEquals("Globetrotters", TestUtility.getStringCellValue(implicit, 46, 1));
        assertEquals("After", TestUtility.getStringCellValue(implicit, 47, 0));
        assertEquals(8, implicit.getNumMergedRegions());

        Sheet leftRight = workbook.getSheetAt(1);
        assertEquals("Don't", TestUtility.getStringCellValue(leftRight, 0, 0));
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(leftRight, 0, 1));
        assertEquals("Do", TestUtility.getStringCellValue(leftRight, 0, 6));
        assertEquals("Copy", TestUtility.getStringCellValue(leftRight, 1, 0));
        assertEquals("Wins", TestUtility.getStringCellValue(leftRight, 1, 3));
        assertEquals("not", TestUtility.getStringCellValue(leftRight, 1, 6));
        assertEquals("Me", TestUtility.getStringCellValue(leftRight, 2, 0));
        assertEquals(51, TestUtility.getNumericCellValue(leftRight, 2, 3), DELTA);
        assertEquals("copy", TestUtility.getStringCellValue(leftRight, 2, 6));
        assertEquals("Down!", TestUtility.getStringCellValue(leftRight, 3, 0));
        assertEquals(37, TestUtility.getNumericCellValue(leftRight, 3, 3), DELTA);
        assertEquals("downward!", TestUtility.getStringCellValue(leftRight, 3, 6));
        assertTrue(TestUtility.isCellBlank(leftRight, 4, 0));
        assertEquals(35, TestUtility.getNumericCellValue(leftRight, 4, 3), DELTA);
        assertTrue(TestUtility.isCellBlank(leftRight, 4, 6));
        assertEquals(23, TestUtility.getNumericCellValue(leftRight, 5, 3), DELTA);
        assertEquals(20, TestUtility.getNumericCellValue(leftRight, 6, 3), DELTA);
        assertEquals("Division: Central", TestUtility.getStringCellValue(leftRight, 7, 1));
        assertEquals("Division: Empty", TestUtility.getStringCellValue(leftRight, 42, 1));
        assertTrue(TestUtility.isMergedRegionPresent(leftRight, new CellRangeAddress(42, 42, 1, 5)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(leftRight, 44, 1));
        assertTrue(TestUtility.isMergedRegionPresent(leftRight, new CellRangeAddress(44, 44, 1, 5)));
        assertEquals(21227, TestUtility.getNumericCellValue(leftRight, 46, 3), DELTA);
        assertEquals("After", TestUtility.getStringCellValue(leftRight, 47, 1));
        assertEquals(8, leftRight.getNumMergedRegions());

        Sheet fixedHoriz = workbook.getSheetAt(2);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(fixedHoriz, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(fixedHoriz, new CellRangeAddress(0, 4, 0, 0)));
        assertEquals("Boston", TestUtility.getStringCellValue(fixedHoriz, 0, 2));
        assertEquals("Right", TestUtility.getStringCellValue(fixedHoriz, 0, 7));
        assertEquals("76ers", TestUtility.getStringCellValue(fixedHoriz, 1, 3));
        assertTrue(TestUtility.isCellBlank(fixedHoriz, 1, 7));
        assertEquals(35, TestUtility.getNumericCellValue(fixedHoriz, 2, 4), DELTA);
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(fixedHoriz, 20, 0));
        assertTrue(TestUtility.isMergedRegionPresent(fixedHoriz, new CellRangeAddress(20, 24, 0, 0)));
        assertEquals("Sacramento", TestUtility.getStringCellValue(fixedHoriz, 20, 6));
        assertEquals("Right", TestUtility.getStringCellValue(fixedHoriz, 20, 7));
        assertEquals("Lakers", TestUtility.getStringCellValue(fixedHoriz, 21, 2));
        assertEquals(42, TestUtility.getNumericCellValue(fixedHoriz, 23, 4), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(fixedHoriz, 30, 0));
        assertTrue(TestUtility.isMergedRegionPresent(fixedHoriz, new CellRangeAddress(30, 34, 0, 0)));
        assertTrue(TestUtility.isCellBlank(fixedHoriz, 30, 2));
        assertEquals("Right", TestUtility.getStringCellValue(fixedHoriz, 30, 7));
        assertTrue(TestUtility.isCellBlank(fixedHoriz, 34, 2));
        assertEquals("Right", TestUtility.getStringCellValue(fixedHoriz, 35, 7));
        assertEquals(21227, TestUtility.getNumericCellValue(fixedHoriz, 37, 2), DELTA);
        assertTrue(TestUtility.isCellBlank(fixedHoriz, 38, 3));
        assertEquals("Below", TestUtility.getStringCellValue(fixedHoriz, 40, 0));
        assertEquals(8, fixedHoriz.getNumMergedRegions());

        Sheet turnOff = workbook.getSheetAt(3);
        assertEquals("California", TestUtility.getStringCellValue(turnOff, 0, 1));
        assertEquals(58, TestUtility.getNumericCellValue(turnOff, 0, 3), DELTA);
        assertTrue(TestUtility.isCellBlank(turnOff, 0, 4));
        assertEquals("Nevada", TestUtility.getStringCellValue(turnOff, 1, 1));
        assertEquals(17, TestUtility.getNumericCellValue(turnOff, 1, 3), DELTA);
        assertTrue(TestUtility.isCellBlank(turnOff, 1, 4));
        assertTrue(TestUtility.isCellBlank(turnOff, 2, 0));
        assertTrue(TestUtility.isCellBlank(turnOff, 2, 1));
        assertTrue(TestUtility.isCellBlank(turnOff, 2, 3));

        Sheet noPae = workbook.getSheetAt(4);
        assertEquals("Harlem", TestUtility.getStringCellValue(noPae, 2, 0));
        assertEquals("Lakers", TestUtility.getStringCellValue(noPae, 2, 6));
        assertTrue(TestUtility.isCellBlank(noPae, 3, 0));
        assertEquals("Kings", TestUtility.getStringCellValue(noPae, 6, 6));
        assertTrue(TestUtility.isCellBlank(noPae, 7, 6));
        Cell cNoPae = TestUtility.getCell(noPae, 6, 2);
        assertNotNull(cNoPae);
        CellStyle csNoPae = cNoPae.getCellStyle();
        assertEquals(BorderStyle.THIN, csNoPae.getBorderBottom());
        assertEquals(BorderStyle.THIN, csNoPae.getBorderTop());
        assertEquals(BorderStyle.THIN, csNoPae.getBorderLeft());
        assertEquals(BorderStyle.THIN, csNoPae.getBorderRight());
        assertEquals("c0c0c0", TestUtility.getCellForegroundColorString(noPae, 6, 2));

        Sheet paeClear = workbook.getSheetAt(5);
        assertEquals("Harlem", TestUtility.getStringCellValue(paeClear, 2, 0));
        assertEquals("Lakers", TestUtility.getStringCellValue(paeClear, 2, 6));
        assertTrue(TestUtility.isCellBlank(paeClear, 3, 0));
        assertEquals("Kings", TestUtility.getStringCellValue(paeClear, 6, 6));
        assertTrue(TestUtility.isCellBlank(paeClear, 7, 6));
        Cell cPaeClear = TestUtility.getCell(paeClear, 6, 2);
        assertNotNull(cPaeClear);
        CellStyle csPaeClear = cPaeClear.getCellStyle();
        assertEquals(BorderStyle.THIN, csPaeClear.getBorderBottom());
        assertEquals(BorderStyle.THIN, csPaeClear.getBorderTop());
        assertEquals(BorderStyle.THIN, csPaeClear.getBorderLeft());
        assertEquals(BorderStyle.THIN, csPaeClear.getBorderRight());
        assertEquals("c0c0c0", TestUtility.getCellForegroundColorString(paeClear, 6, 2));

        Sheet paeRemove = workbook.getSheetAt(6);
        assertEquals("Harlem", TestUtility.getStringCellValue(paeRemove, 2, 0));
        assertEquals("Lakers", TestUtility.getStringCellValue(paeRemove, 2, 6));
        assertTrue(TestUtility.isCellBlank(paeRemove, 3, 0));
        assertEquals("Kings",TestUtility. getStringCellValue(paeRemove, 6, 6));
        assertTrue(TestUtility.isCellBlank(paeRemove, 7, 6));
        Cell cPaeRemove = TestUtility.getCell(paeRemove, 6, 2);
        assertNull(cPaeRemove);

        Sheet paeReplaceValue = workbook.getSheetAt(7);
        assertEquals("Globetrotters of Harlem", TestUtility.getStringCellValue(paeReplaceValue, 2, 0));
        assertEquals(21227, TestUtility.getNumericCellValue(paeReplaceValue, 2, 1), DELTA);
        for (int r = 3; r < 7; r++)
        {
            assertEquals("- of -", TestUtility.getStringCellValue(paeReplaceValue, r, 0));
            assertEquals("-", TestUtility.getStringCellValue(paeReplaceValue, r, 1));
            assertEquals("-", TestUtility.getStringCellValue(paeReplaceValue, r, 2));
            assertEquals("-", TestUtility.getStringCellValue(paeReplaceValue, r, 3));

            assertEquals("c0c0c0", TestUtility.getCellForegroundColorString(paeReplaceValue, r, 0));
        }
        assertEquals("Lakers of Los Angeles", TestUtility.getStringCellValue(paeReplaceValue, 2, 4));
        assertEquals(36, TestUtility.getNumericCellValue(paeReplaceValue, 3, 5), DELTA);
        assertEquals((20.0) / (20.0 + 52.0), TestUtility.getNumericCellValue(paeReplaceValue, 6, 7), DELTA);

        // Cannot test for grouping but can test for the collapse side effect.
        Sheet groupDirNone = workbook.getSheetAt(8);
        for (int r = 0; r < 48; r++)
        {
            assertFalse(groupDirNone.getRow(r) != null && groupDirNone.getRow(r).getZeroHeight());
        }
        for (int c = 0; c < 6; c++)
        {
            assertFalse(groupDirNone.isColumnHidden(c));
        }

        Sheet groupDirRows = workbook.getSheetAt(9);
        for (int r = 0; r < 48; r++)
        {
            // These rows are collapsed.
            if (r >= 16 && r <= 20)
            {
                assertTrue(groupDirRows.getRow(r).getZeroHeight());
            }
            else
            {
                assertFalse(groupDirRows.getRow(r) != null && groupDirRows.getRow(r).getZeroHeight());
            }
        }
        for (int c = 0; c < 6; c++)
        {
            assertFalse(groupDirRows.isColumnHidden(c));
        }

        Sheet groupDirCols = workbook.getSheetAt(10);
        for (int r = 0; r < 8; r++)
        {
            assertFalse(groupDirCols.getRow(r) != null && groupDirCols.getRow(r).getZeroHeight());
        }
        for (int c = 0; c < 48; c++)
        {
            // These columns are collapsed.
            if (c >= 13 && c <= 17)
            {
                assertTrue(groupDirCols.isColumnHidden(c));
            }
            else
            {
                assertFalse(groupDirCols.isColumnHidden(c));
            }
        }

        Sheet indexVar = workbook.getSheetAt(11);
        assertEquals("1. Boston", TestUtility.getStringCellValue(indexVar, 2, 0));
        assertEquals("5. Toronto", TestUtility.getStringCellValue(indexVar, 6, 0));
        assertEquals("2. Indiana", TestUtility.getStringCellValue(indexVar, 10, 0));
        assertEquals("3. Milwaukee", TestUtility.getStringCellValue(indexVar, 11, 0));
        assertEquals("4. Detroit", TestUtility.getStringCellValue(indexVar, 12, 0));
        assertEquals("1. Los Angeles", TestUtility.getStringCellValue(indexVar, 30, 0));
        assertEquals("4. Los Angeles", TestUtility.getStringCellValue(indexVar, 33, 0));
        assertEquals("5. Houston", TestUtility.getStringCellValue(indexVar, 41, 0));
        assertEquals("1. Harlem", TestUtility.getStringCellValue(indexVar, 46, 0));

        Sheet limit = workbook.getSheetAt(12);
        assertEquals("Boston", TestUtility.getStringCellValue(limit, 2, 0));
        assertEquals("New York", TestUtility.getStringCellValue(limit, 4, 0));
        assertEquals("Division: Central", TestUtility.getStringCellValue(limit, 5, 0));
        assertEquals("San Antonio", TestUtility.getStringCellValue(limit, 27, 0));
        assertEquals("New Orleans", TestUtility.getStringCellValue(limit, 29, 0));
        assertTrue(TestUtility.isCellBlank(limit, 32, 0));
        assertTrue(TestUtility.isCellBlank(limit, 34, 0));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(limit, 35, 0));
        assertEquals("Harlem", TestUtility.getStringCellValue(limit, 37, 0));
        assertTrue(TestUtility.isCellBlank(limit, 38, 0));
        assertTrue(TestUtility.isCellBlank(limit, 39, 0));

        // Test replacing collection name, but not as part of a large identifier.
        Sheet collNameReplace = workbook.getSheetAt(13);
        assertEquals("Radiator Springs", TestUtility.getStringCellValue(collNameReplace, 2, 4));
        assertEquals("Springfield", TestUtility.getStringCellValue(collNameReplace, 13, 4));
        assertTrue(TestUtility.isCellBlank(collNameReplace, 14, 4));

        // Test specifying "left" only.
        Sheet leftOnly = workbook.getSheetAt(14);
        assertEquals("Don't", TestUtility.getStringCellValue(leftOnly, 0, 0));
        assertEquals("Division:", TestUtility.getStringCellValue(leftOnly, 0, 1));
        assertEquals("Atlantic", TestUtility.getStringCellValue(leftOnly, 0, 6));
        assertEquals("Do", TestUtility.getStringCellValue(leftOnly, 0, 7));
        assertEquals("Copy", TestUtility.getStringCellValue(leftOnly, 1, 0));
        assertEquals("Wins", TestUtility.getStringCellValue(leftOnly, 1, 3));
        assertEquals("not", TestUtility.getStringCellValue(leftOnly, 1, 7));
        assertEquals("Me", TestUtility.getStringCellValue(leftOnly, 2, 0));
        assertEquals(51, TestUtility.getNumericCellValue(leftOnly, 2, 3), DELTA);
        assertEquals(82, TestUtility.getNumericCellValue(leftOnly, 2, 6), DELTA);
        assertEquals("copy", TestUtility.getStringCellValue(leftOnly, 2, 7));
        assertEquals("Down!", TestUtility.getStringCellValue(leftOnly, 3, 0));
        assertEquals(37, TestUtility.getNumericCellValue(leftOnly, 3, 3), DELTA);
        assertEquals("downward!", TestUtility.getStringCellValue(leftOnly, 3, 7));
        assertTrue(TestUtility.isCellBlank(leftOnly, 4, 0));
        assertEquals(35, TestUtility.getNumericCellValue(leftOnly, 4, 3), DELTA);
        assertTrue(TestUtility.isCellBlank(leftOnly, 4, 7));
        assertEquals(23, TestUtility.getNumericCellValue(leftOnly, 5, 3), DELTA);
        assertEquals(20, TestUtility.getNumericCellValue(leftOnly, 6, 3), DELTA);
        assertEquals("Division:", TestUtility.getStringCellValue(leftOnly, 7, 1));
        assertEquals("Central", TestUtility.getStringCellValue(leftOnly, 7, 6));
        assertEquals("Division:", TestUtility.getStringCellValue(leftOnly, 42, 1));
        assertEquals("Empty", TestUtility.getStringCellValue(leftOnly, 42, 6));
        assertTrue(TestUtility.isMergedRegionPresent(leftOnly, new CellRangeAddress(42, 42, 1, 5)));
        assertEquals("Division:", TestUtility.getStringCellValue(leftOnly, 44, 1));
        assertEquals("Of Their Own", TestUtility.getStringCellValue(leftOnly, 44, 6));
        assertTrue(TestUtility.isMergedRegionPresent(leftOnly, new CellRangeAddress(44, 44, 1, 5)));
        assertEquals(21227, TestUtility.getNumericCellValue(leftOnly, 46, 3), DELTA);
        assertEquals("After", TestUtility.getStringCellValue(leftOnly, 47, 1));
        assertEquals(8, leftOnly.getNumMergedRegions());

        // Test specifying "right" only.
        Sheet rightOnly = workbook.getSheetAt(15);
        assertEquals("Don't", TestUtility.getStringCellValue(rightOnly, 0, 0));
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(rightOnly, 0, 1));
        assertEquals("Do", TestUtility.getStringCellValue(rightOnly, 0, 7));
        assertEquals("Copy", TestUtility.getStringCellValue(rightOnly, 1, 0));
        assertEquals("Wins", TestUtility.getStringCellValue(rightOnly, 1, 3));
        assertEquals("not", TestUtility.getStringCellValue(rightOnly, 1, 7));
        assertEquals("Me", TestUtility.getStringCellValue(rightOnly, 2, 0));
        assertEquals(51, TestUtility.getNumericCellValue(rightOnly, 2, 3), DELTA);
        assertEquals(82, TestUtility.getNumericCellValue(rightOnly, 2, 6), DELTA);
        assertEquals("copy", TestUtility.getStringCellValue(rightOnly, 2, 7));
        assertEquals("Down!", TestUtility.getStringCellValue(rightOnly, 3, 0));
        assertEquals(37, TestUtility.getNumericCellValue(rightOnly, 3, 3), DELTA);
        assertEquals("downward!", TestUtility.getStringCellValue(rightOnly, 3, 7));
        assertTrue(TestUtility.isCellBlank(rightOnly, 4, 0));
        assertEquals(35, TestUtility.getNumericCellValue(rightOnly, 4, 3), DELTA);
        assertTrue(TestUtility.isCellBlank(rightOnly, 4, 7));
        assertEquals(23, TestUtility.getNumericCellValue(rightOnly, 5, 3), DELTA);
        assertEquals(20, TestUtility.getNumericCellValue(rightOnly, 6, 3), DELTA);
        assertEquals("Division: Central", TestUtility.getStringCellValue(rightOnly, 7, 1));
        assertEquals("Division: Empty", TestUtility.getStringCellValue(rightOnly, 42, 1));
        assertTrue(TestUtility.isMergedRegionPresent(rightOnly, new CellRangeAddress(42, 42, 1, 6)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(rightOnly, 44, 1));
        assertTrue(TestUtility.isMergedRegionPresent(rightOnly, new CellRangeAddress(44, 44, 1, 6)));
        assertEquals(21227, TestUtility.getNumericCellValue(rightOnly, 46, 3), DELTA);
        assertEquals("After", TestUtility.getStringCellValue(rightOnly, 47, 1));
        assertEquals(8, rightOnly.getNumMergedRegions());

        Sheet varStatus = workbook.getSheetAt(16);
        Map<Integer, String> expInds = new HashMap<Integer, String>();
        expInds.put(1, "0 of 8");  expInds.put(2, "0 of 5");  expInds.put(3, "1 of 5");  expInds.put(4, "2 of 5");  expInds.put(5, "3 of 5");  expInds.put(6, "4 of 5");
        expInds.put(8, "1 of 8");  expInds.put(9, "0 of 5");  expInds.put(10, "1 of 5"); expInds.put(11, "2 of 5"); expInds.put(12, "3 of 5"); expInds.put(13, "4 of 5");
        expInds.put(15, "2 of 8"); expInds.put(16, "0 of 5"); expInds.put(17, "1 of 5"); expInds.put(18, "2 of 5"); expInds.put(19, "3 of 5"); expInds.put(20, "4 of 5");
        expInds.put(22, "3 of 8"); expInds.put(23, "0 of 5"); expInds.put(24, "1 of 5"); expInds.put(25, "2 of 5"); expInds.put(26, "3 of 5"); expInds.put(27, "4 of 5");
        expInds.put(29, "4 of 8"); expInds.put(30, "0 of 5"); expInds.put(31, "1 of 5"); expInds.put(32, "2 of 5"); expInds.put(33, "3 of 5"); expInds.put(34, "4 of 5");
        expInds.put(36, "5 of 8"); expInds.put(37, "0 of 5"); expInds.put(38, "1 of 5"); expInds.put(39, "2 of 5"); expInds.put(40, "3 of 5"); expInds.put(41, "4 of 5");
        expInds.put(43, "6 of 8");
        expInds.put(45, "7 of 8"); expInds.put(46, "0 of 1");
        List<Integer> firstTrues = Arrays.asList(1, 2, 9, 16, 23, 30, 37, 46);
        List<Integer> lastTrues = Arrays.asList(6, 13, 20, 27, 34, 41, 45, 46);

        for (int r = 0; r < 47; r++)
        {
            if (expInds.containsKey(r))
            {
                assertEquals("Row " + r, expInds.get(r), TestUtility.getStringCellValue(varStatus, r, 5));
                assertEquals("Row " + r, firstTrues.contains(r), TestUtility.getBooleanCellValue(varStatus, r, 6));
                assertEquals("Row " + r, lastTrues.contains(r), TestUtility.getBooleanCellValue(varStatus, r, 7));
            }
        }
    }

    /**
     * This test is a single map test.
     * @return <code>false</code>.
     */
    @Override
    protected boolean isMultipleBeans()
    {
        return false;
    }

    /**
     * For single beans map tests, return the <code>Map</code> of bean names to
     * bean values.
     * @return A <code>Map</code> of bean names to bean values.
     */
    @Override
    protected Map<String, Object> getBeansMap()
    {
        Map<String, Object> beans = TestUtility.getDivisionData();
        beans.putAll(TestUtility.getStateData());
        beans.putAll(TestUtility.getSpecificDivisionData(4, "pacific"));
        beans.putAll(TestUtility.getSpecificDivisionData(7, "ofTheirOwn"));
        beans.putAll(TestUtility.getFictionalCountyData());
        return beans;
    }
}
