package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

/**
 * This JUnit Test class tests the implicit cloning feature of JETT, with
 * "normal" (one beans map overall) transformation.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class ImplCloningNormalTest extends ImplicitCloningTest
{
    /**
     * Tests the .xls template spreadsheet.
     *
     * @throws java.io.IOException                                        If an I/O error occurs.
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
     *
     * @throws IOException            If an I/O error occurs.
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
     *
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "ImplCloningNormal";
    }

    /**
     * This test is a single map test.
     *
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
     *
     * @return A <code>Map</code> of bean names to bean values.
     */
    @Override
    protected Map<String, Object> getBeansMap()
    {
        Map<String, Object> beans = TestUtility.getDivisionData();
        // Space is at a premium in sheet names -- 31 characters.
        beans.put("dvs", beans.get("divisionsList"));
        beans.remove("divisionsList");
        return beans;
    }
}
