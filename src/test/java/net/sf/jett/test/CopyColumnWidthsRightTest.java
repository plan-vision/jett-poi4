package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import static org.junit.Assert.assertEquals;

/**
 * This JUnit Test class tests the performance of a template that repeatedly
 * calls <code>SheetUtil.copyColumnWidthsRight</code>, which is suspected of
 * copying column widths far to the right of what is necessary.
 *
 * @author Randy Gettman
 * @since 0.11.0
 */
public class CopyColumnWidthsRightTest extends TestCase
{
    /**
     * Tests the .xls template spreadsheet.
     * @throws IOException If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
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
    protected String getExcelNameBase() { return "CopyColumnWidthsRight"; }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet crhd = workbook.getSheetAt(0);
        // The copied row is 289, the shifted row is 1080.
        assertEquals(1080, crhd.getRow(251).getHeight());

        Sheet ccwr = workbook.getSheetAt(1);
        // The copied column is 1645, the shifted column is 4681.
        assertEquals(4681, ccwr.getColumnWidth(251));
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
