package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import static org.junit.Assert.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * This JUnit Test class tests the evaluation of the "group" tag.
 *
 * @author Randy Gettman
 */
public class GroupTagTest extends TestCase
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
        return "GroupTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        // There is no API for determining Excel groups of rows or columns.
        // But it is possible to test the row heights/hidden columns.
        Sheet sheet = workbook.getSheetAt(0);
        assertFalse(sheet.isColumnHidden(0));
        assertFalse(sheet.isColumnHidden(1));
        assertFalse(sheet.isColumnHidden(2));
        assertFalse(sheet.isColumnHidden(3));
        assertFalse(sheet.isColumnHidden(4));  // Affected by a col group
        assertFalse(sheet.isColumnHidden(5));  // Affected by 2 col groups
        assertFalse(sheet.isColumnHidden(6));  // Affected by 2 col groups
        assertFalse(sheet.isColumnHidden(7));  // Affected by a col group
        assertFalse(sheet.isColumnHidden(8));
        assertTrue(sheet.isColumnHidden(9));  // Affected by a col group; hidden
        assertTrue(sheet.isColumnHidden(10));  // Affected by a col group; hidden
        assertTrue(sheet.isColumnHidden(11));  // Affected by a col group; hidden
        assertFalse(sheet.isColumnHidden(12));

        assertFalse(sheet.getRow(0).getZeroHeight());
        assertFalse(sheet.getRow(1) != null && sheet.getRow(1).getZeroHeight());
        assertFalse(sheet.getRow(2).getZeroHeight());  // Affected by a row group
        assertFalse(sheet.getRow(3).getZeroHeight());  // Affected by 2 row groups
        assertFalse(sheet.getRow(4).getZeroHeight());  // Affected by 2 row groups
        assertFalse(sheet.getRow(5).getZeroHeight());  // Affected by a row group
        assertFalse(sheet.getRow(6) != null && sheet.getRow(6).getZeroHeight());
        assertFalse(sheet.getRow(7).getZeroHeight());  // Affected by a row group
        assertFalse(sheet.getRow(8).getZeroHeight());  // Affected by a row group
        assertFalse(sheet.getRow(9).getZeroHeight());  // Affected by a row group
        assertFalse(sheet.getRow(10) != null && sheet.getRow(10).getZeroHeight());
        assertTrue(sheet.getRow(11).getZeroHeight());  // Affected by a row group; hidden
        assertTrue(sheet.getRow(12).getZeroHeight());  // Affected by a row group; hidden
        assertTrue(sheet.getRow(13).getZeroHeight());  // Affected by a row group; hidden
        assertFalse(sheet.getRow(14) != null && sheet.getRow(14).getZeroHeight());
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
        beans.put("pi", Math.PI);
        return beans;
    }
}
