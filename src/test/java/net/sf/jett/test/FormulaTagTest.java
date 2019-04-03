package net.sf.jett.test;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "formula" tag (always
 * bodiless).
 *
 * @author Randy Gettman
 * @since 0.4.0
 */
public class FormulaTagTest extends TestCase
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
        return "FormulaTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet bean = workbook.getSheetAt(0);
        assertEquals("IF(ISERROR(51/(51+21)),\"N/A\",51/(51+21))", TestUtility.getFormulaCellValue(bean, 2, 4));
        assertEquals("(30-(-3))/2", TestUtility.getFormulaCellValue(bean, 4, 5));
        assertEquals("32+42", TestUtility.getFormulaCellValue(bean, 10, 6));
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
        Map<String, Object> beans = TestUtility.getDivisionData();
        List<String> dynFormulas = Arrays.asList(
                "${team.wins} / (${team.wins} + ${team.losses})",
                "(${jagg:eval(division.teams, 'Max(numGamesAboveEven)')} - (${team.numGamesAboveEven})) / 2",
                "${team.wins} + ${team.losses}");
        beans.put("dynFormulas", dynFormulas);
        return beans;
    }
}