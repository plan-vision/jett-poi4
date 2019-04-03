package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "set" tag.
 *
 * @author Randy Gettman
 * @since 0.11.0
 */
public class SetTagTest extends TestCase
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
        return "SetTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet setTag = workbook.getSheetAt(0);
        assertEquals("Ah, what's up Doc?", TestUtility.getStringCellValue(setTag, 1, 2));
        assertEquals("Bugs: Ah, what's up Doc?", TestUtility.getStringCellValue(setTag, 1, 3));
        assertEquals("1", TestUtility.getStringCellValue(setTag, 2, 2));
        assertEquals("The answer is \"1\".", TestUtility.getStringCellValue(setTag, 2, 3));
        assertEquals(1.0, TestUtility.getNumericCellValue(setTag, 3, 2), DELTA);
        assertEquals("y + 1 = 2.", TestUtility.getStringCellValue(setTag, 3, 3));
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
        Map<String, Object> beans = new HashMap<>();
        beans.putAll(TestUtility.getEmployeeData());
        return beans;
    }
}
