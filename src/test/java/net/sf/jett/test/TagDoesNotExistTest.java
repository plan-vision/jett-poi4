package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;
import java.util.HashMap;

import static org.junit.Assert.*;
import org.junit.Test;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.exception.TagParseException;

/**
 * This JUnit Test class tests the error message of a tag that doesn't exist.
 *
 * @author Randy Gettman
 * @since 0.9.0
 */
public class TagDoesNotExistTest extends TestCase
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
        try
        {
            super.testXls();
            fail();
        }
        catch(TagParseException e)
        {
            testExceptionMessage(e);
        }
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
        try
        {
            super.testXlsx();
            fail();
        }
        catch(TagParseException e)
        {
            testExceptionMessage(e);
        }
    }

    /**
     * Returns the Excel name base for the template and resultant spreadsheets
     * for this test.
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "TagDoesNotExist";
    }

    /**
     * Validate the exception message.
     * @param e The exception.
     */
    private void testExceptionMessage(TagParseException e)
    {
        String newline = System.getProperty("line.separator");
        String expected =
                "Invalid tag: <jt:doesnotexist/> at DNE!B47 (originally located at DNE!B3)" + newline +
                        "  inside tag \"if\" (net.sf.jett.tag.IfTag), at DNE!B47 (originally at DNE!B3)" + newline +
                        "  inside tag \"forEach\" (net.sf.jett.tag.ForEachTag), at DNE!A47 (originally at DNE!A3)" + newline +
                        "  inside tag \"forEach\" (net.sf.jett.tag.ForEachTag), at DNE!A1 (originally at DNE!A1)";
        assertEquals(expected, e.getMessage());
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        // Error expected.  See the testExceptionMessage method.
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
        beans.putAll(TestUtility.getDivisionData());
        return beans;
    }
}
