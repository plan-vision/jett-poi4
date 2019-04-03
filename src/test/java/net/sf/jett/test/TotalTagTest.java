package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "total" tag.
 *
 * @author Randy Gettman
 */
public class TotalTagTest extends TestCase
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
        return "TotalTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet total = workbook.getSheetAt(0);
        assertEquals(58, TestUtility.getNumericCellValue(total, 1, 1), DELTA);
        assertEquals(17, TestUtility.getNumericCellValue(total, 1, 2), DELTA);
        assertEquals(38070496, TestUtility.getNumericCellValue(total, 2, 1), DELTA);
        assertEquals(1998257, TestUtility.getNumericCellValue(total, 2, 2), DELTA);
        assertEquals(180979, TestUtility.getNumericCellValue(total, 3, 1), DELTA);
        assertEquals(16106, TestUtility.getNumericCellValue(total, 3, 2), DELTA);
        assertEquals(404224, TestUtility.getNumericCellValue(total, 4, 1), DELTA);
        assertEquals(284401, TestUtility.getNumericCellValue(total, 4, 2), DELTA);
        assertEquals(1222, TestUtility.getNumericCellValue(total, 5, 1), DELTA);
        assertEquals(971, TestUtility.getNumericCellValue(total, 5, 2), DELTA);
        assertEquals(51960, TestUtility.getNumericCellValue(total, 6, 1), DELTA);
        assertEquals(47001, TestUtility.getNumericCellValue(total, 6, 2), DELTA);
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