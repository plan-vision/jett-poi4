package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the performance of merged region sheet
 * manipulation in JETT.  This specifically affects Ticket #29.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class MergedRegionLoadTest extends TestCase
{
    private long startTime;

    /**
     * Tests the .xls template spreadsheet.
     * @throws java.io.IOException If an I/O error occurs.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the input spreadsheet is invalid.
     */
    @Override
    @Test
    public void testXls() throws IOException, InvalidFormatException
    {
        startTime = System.nanoTime();
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
        startTime = System.nanoTime();
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
        return "MergedRegionLoad";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        long endTime = System.nanoTime();
        double seconds = (endTime - startTime) / 1000000000.0;
        // Testing on my PC: 4 trials on .xls averages 2.439 s; 4 trials on .xlsx averages 5.275 s.
        // Allow for variations.  I had to kill the baseline .xlsx test (without
        // merged region caching in the TagContext) after about 2-3 minutes.
        if (seconds > 15.0)
            fail("Merged Region Load Test took longer than 15 seconds: " + seconds);
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
        return TestUtility.getDummyDivisionsData();
    }
}