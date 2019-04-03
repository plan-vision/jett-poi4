package net.sf.jett.test;

import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "for" tag in various
 * cases.
 *
 * @author Randy Gettman
 */
public class ForTagTest extends TestCase
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
        return "ForTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet multiplication = workbook.getSheetAt(0);
        for (int c = 1; c < 20; c++)
        {
            assertEquals(c, TestUtility.getNumericCellValue(multiplication, 0, c), DELTA);
        }
        assertTrue(TestUtility.isCellBlank(multiplication, 0, 21));
        for (int r = 1; r < 20; r++)
        {
            assertEquals(r, TestUtility.getNumericCellValue(multiplication, r, 0), DELTA);
            for (int c = 1; c < 20; c++)
            {
                assertEquals(c * r, TestUtility.getNumericCellValue(multiplication, r, c), DELTA);
            }
            assertTrue(TestUtility.isCellBlank(multiplication, 4, 21));
        }
        for (int c = 0; c < 21; c++)
        {
            assertTrue(TestUtility.isCellBlank(multiplication, 21, c));
        }

        Sheet oneOrZero = workbook.getSheetAt(1);
        assertEquals(23, TestUtility.getNumericCellValue(oneOrZero, 0, 0), DELTA);
        assertEquals("is the only element!", TestUtility.getStringCellValue(oneOrZero, 0, 1));
        assertEquals("After", TestUtility.getStringCellValue(oneOrZero, 1, 0));
        assertEquals("After2", TestUtility.getStringCellValue(oneOrZero, 2, 0));

        Sheet end = workbook.getSheetAt(2);
        for (int c = 1; c < 27; c++)
        {
            int x = 100 - 4 * (c - 1);
            assertEquals(x, TestUtility.getNumericCellValue(end, 0, c), DELTA);
            assertEquals(x * x, TestUtility.getNumericCellValue(end, 1, c), DELTA);
        }
        assertTrue(TestUtility.isCellBlank(end, 0, 27));
        for (int c = 1; c < 21; c++)
        {
            int y = 1 + 5 * (c - 1);
            boolean isMultOf3 = (y % 3) == 0;
            assertEquals(y, TestUtility.getNumericCellValue(end, 3, c), DELTA);
            if (isMultOf3)
                assertTrue(TestUtility.getBooleanCellValue(end, 4, c));
            else
                assertFalse(TestUtility.getBooleanCellValue(end, 4, c));
        }
        assertTrue(TestUtility.isCellBlank(end, 3, 21));

        Sheet immaterial = workbook.getSheetAt(3);
        for (int r = 1; r < 7; r++)
        {
            int x = 12 - 2 * r;
            assertEquals(x, TestUtility.getNumericCellValue(immaterial, r, 0), DELTA);
            assertEquals(x * x, TestUtility.getNumericCellValue(immaterial, r, 1), DELTA);
        }
        assertTrue(TestUtility.isCellBlank(immaterial, 7, 0));
        assertTrue(TestUtility.isCellBlank(immaterial, 7, 1));
        assertEquals("ff0000", TestUtility.getCellForegroundColorString(immaterial, 7, 0));
        assertEquals("ff0000", TestUtility.getCellForegroundColorString(immaterial, 7, 1));

        Sheet varStatus = workbook.getSheetAt(4);
        List<Integer> expXVals  = Arrays.asList(1, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 4, 4, 5);
        List<Integer> expYVals  = Arrays.asList(5, 4, 3, 2, 1, 5, 4, 3, 2, 5, 4, 3, 5, 4, 5);
        List<Integer> expStartX = Collections.nCopies(15, 1);
        List<Integer> expEndX   = Collections.nCopies(15, 5);
        List<Integer> expStepX  = Collections.nCopies(15, 1);
        List<Integer> expStartY = Collections.nCopies(15, 5);
        List<Integer> expStepY  = Collections.nCopies(15, -1);

        for (int r = 1; r < 16; r++)
        {
            assertEquals("Row " + r, expXVals.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 0), DELTA);
            assertEquals("Row " + r, expYVals.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 1), DELTA);
            assertEquals("Row " + r, expStartX.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 2), DELTA);
            assertEquals("Row " + r, expEndX.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 3), DELTA);
            assertEquals("Row " + r, expStepX.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 4), DELTA);
            assertEquals("Row " + r, expStartY.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 5), DELTA);
            // End y values are set to x.
            assertEquals("Row " + r, expXVals.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 6), DELTA);
            assertEquals("Row " + r, expStepY.get(r - 1), TestUtility.getNumericCellValue(varStatus, r, 7), DELTA);
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
        return new HashMap<>();
    }
}
