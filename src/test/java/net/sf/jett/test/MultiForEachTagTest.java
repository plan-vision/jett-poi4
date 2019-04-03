package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "multiForEach" tag.
 *
 * @author Randy Gettman
 */
public class MultiForEachTagTest extends TestCase
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
        return "MultiForEachTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet multi = workbook.getSheetAt(0);
        assertEquals(1, TestUtility.getNumericCellValue(multi, 2, 0), DELTA);
        assertEquals("Los Angeles", TestUtility.getStringCellValue(multi, 2, 1));
        assertEquals(1, TestUtility.getNumericCellValue(multi, 2, 7), DELTA);
        assertEquals("Clark", TestUtility.getStringCellValue(multi, 2, 8));
        assertEquals(3146274, TestUtility.getNumericCellValue(multi, 3, 2), DELTA);
        assertEquals(339486, TestUtility.getNumericCellValue(multi, 3, 9), DELTA);
        assertEquals(2046, TestUtility.getNumericCellValue(multi, 4, 3), DELTA);
        assertEquals(373, TestUtility.getNumericCellValue(multi, 4, 10), DELTA);
        assertEquals(1893, TestUtility.getNumericCellValue(multi, 5, 4), DELTA);
        assertEquals(1869, TestUtility.getNumericCellValue(multi, 5, 11), DELTA);
        assertEquals("San Bernardino", TestUtility.getStringCellValue(multi, 6, 5));
        assertEquals("Minden", TestUtility.getStringCellValue(multi, 6, 12));
        assertEquals("06085", TestUtility.getStringCellValue(multi, 7, 6));
        assertEquals("32019", TestUtility.getStringCellValue(multi, 7, 13));
        assertEquals("Tulare", TestUtility.getStringCellValue(multi, 19, 1));
        assertEquals(18, TestUtility.getNumericCellValue(multi, 19, 7), DELTA);
        assertTrue(TestUtility.isCellBlank(multi, 19, 8));
        assertEquals(19, TestUtility.getNumericCellValue(multi, 20, 0), DELTA);
        assertEquals(19, TestUtility.getNumericCellValue(multi, 20, 7), DELTA);
        assertEquals(20, TestUtility.getNumericCellValue(multi, 21, 0), DELTA);
        assertTrue(TestUtility.isCellBlank(multi, 21, 8));
        assertEquals("Alpine", TestUtility.getStringCellValue(multi, 59, 1));
        assertTrue(TestUtility.isCellBlank(multi, 59, 8));
        assertTrue(TestUtility.isCellBlank(multi, 60, 0));
        assertTrue(TestUtility.isCellBlank(multi, 60, 7));

        Sheet limit = workbook.getSheetAt(1);
        assertEquals(1, TestUtility.getNumericCellValue(limit, 2, 0), DELTA);
        assertEquals("Los Angeles", TestUtility.getStringCellValue(limit, 2, 1));
        assertEquals(1, TestUtility.getNumericCellValue(limit, 2, 7), DELTA);
        assertEquals("Clark", TestUtility.getStringCellValue(limit, 2, 8));
        assertEquals(3146274, TestUtility.getNumericCellValue(limit, 3, 2), DELTA);
        assertEquals(339486, TestUtility.getNumericCellValue(limit, 3, 9), DELTA);
        assertEquals(2046, TestUtility.getNumericCellValue(limit, 4, 3), DELTA);
        assertEquals(373, TestUtility.getNumericCellValue(limit, 4, 10), DELTA);
        assertEquals(1893, TestUtility.getNumericCellValue(limit, 5, 4), DELTA);
        assertEquals(1869, TestUtility.getNumericCellValue(limit, 5, 11), DELTA);
        assertEquals("San Bernardino", TestUtility.getStringCellValue(limit, 6, 5));
        assertEquals("Minden", TestUtility.getStringCellValue(limit, 6, 12));
        assertEquals("06085", TestUtility.getStringCellValue(limit, 7, 6));
        assertEquals("32019", TestUtility.getStringCellValue(limit, 7, 13));
        assertEquals("Tulare", TestUtility.getStringCellValue(limit, 19, 1));
        assertEquals(18, TestUtility.getNumericCellValue(limit, 19, 7), DELTA);
        assertTrue(TestUtility.isCellBlank(limit, 19, 8));
        assertEquals(19, TestUtility.getNumericCellValue(limit, 20, 0), DELTA);
        assertEquals(19, TestUtility.getNumericCellValue(limit, 20, 7), DELTA);
        assertEquals(20, TestUtility.getNumericCellValue(limit, 21, 0), DELTA);
        assertTrue(TestUtility.isCellBlank(limit, 21, 8));
        assertTrue(TestUtility.isCellBlank(limit, 22, 0));
        assertTrue(TestUtility.isCellBlank(limit, 22, 7));
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
        return TestUtility.getStateData();
    }
}