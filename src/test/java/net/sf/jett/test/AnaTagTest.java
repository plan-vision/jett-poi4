package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "ana" tag.
 *
 * @author Randy Gettman
 * @since 0.9.0
 */
public class AnaTagTest extends TestCase
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
     * @throws java.io.IOException If an I/O error occurs.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the input spreadsheet is invalid.
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
        return "AnaTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet ana = workbook.getSheetAt(0);
        assertEquals("California", TestUtility.getStringCellValue(ana, 2, 0));
        assertEquals("Los Angeles", TestUtility.getStringCellValue(ana, 2, 1));
        assertEquals(10363850.0 / 38070496.0, TestUtility.getNumericCellValue(ana, 2, 3), DELTA);
        assertEquals(10515.0    / 404224.0  , TestUtility.getNumericCellValue(ana, 2, 5), DELTA);

        assertEquals("California", TestUtility.getStringCellValue(ana, 59, 0));
        assertEquals("Alpine", TestUtility.getStringCellValue(ana, 59, 1));
        assertEquals(1222.0     / 38070496.0, TestUtility.getNumericCellValue(ana, 59, 3), DELTA);
        assertEquals(1914.0     / 404224.0  , TestUtility.getNumericCellValue(ana, 59, 5), DELTA);

        assertEquals("Nevada", TestUtility.getStringCellValue(ana, 60, 0));
        assertEquals("Clark", TestUtility.getStringCellValue(ana, 60, 1));
        assertEquals(1375765.0  / 1998257.0 , TestUtility.getNumericCellValue(ana, 60, 3), DELTA);
        assertEquals(20489.0    / 284401.0  , TestUtility.getNumericCellValue(ana, 60, 5), DELTA);

        assertEquals("Nevada", TestUtility.getStringCellValue(ana, 76, 0));
        assertEquals("Esmeralda", TestUtility.getStringCellValue(ana, 76, 1));
        assertEquals(971.0      / 1998257.0 , TestUtility.getNumericCellValue(ana, 76, 3), DELTA);
        assertEquals(9295.0     / 284401.0  , TestUtility.getNumericCellValue(ana, 76, 5), DELTA);
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
        return TestUtility.getCountyData();
    }
}