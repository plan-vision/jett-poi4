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
 * This JUnit Test class tests the exposure of the POI Objects "sheet" and
 * "workbook" as the current <code>Sheet</code> and <code>Workbook</code>,
 * respectively.
 *
 * @author Randy Gettman
 */
public class PoiObjectAccessTest extends TestCase
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
        return "PoiObjectAccess";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet poiObjects = workbook.getSheetAt(0);
        assertEquals("Number of Sheets: 2", TestUtility.getStringCellValue(poiObjects, 0, 0));
        assertTrue(TestUtility.isCellBlank(poiObjects, 1, 0));
        assertTrue(TestUtility.isCellBlank(poiObjects, 2, 0));
        assertEquals("JETT Header - First Sheet", poiObjects.getHeader().getCenter());
        assertEquals("Last Modified: 2012 Jun 20", poiObjects.getFooter().getRight());

        Sheet second = workbook.getSheetAt(1);
        assertEquals("Number of Sheets: 2", TestUtility.getStringCellValue(second, 0, 0));
        assertTrue(TestUtility.isCellBlank(second, 1, 0));
        assertTrue(TestUtility.isCellBlank(second, 2, 0));
        assertEquals("JETT Header - Second Sheet", second.getHeader().getCenter());
        assertEquals("Last Modified: 2012 Jun 20", second.getFooter().getRight());
        assertEquals("This Cell is at row 4 and column 1.", TestUtility.getStringCellValue(second, 4, 1));
        assertEquals("This Cell's text is wrapped.", TestUtility.getStringCellValue(second, 6, 3));
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
