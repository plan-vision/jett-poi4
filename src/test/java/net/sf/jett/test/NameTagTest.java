package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the behavior of JETT Formulas and Excel named
 * ranges with the "NameTag" tag.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class NameTagTest extends TestCase
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
        return "NameTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Name countyNames = workbook.getName("county_names");
        assertNotNull(countyNames);
        assertEquals("county_names", countyNames.getNameName());
        assertEquals("California!A3:A60", countyNames.getRefersToFormula());
        assertEquals("California", countyNames.getSheetName());

        Name countyPopulations = workbook.getName("county_populations");
        assertNotNull(countyPopulations);
        assertEquals("county_populations", countyPopulations.getNameName());
        assertEquals("California!B3:B60", countyPopulations.getRefersToFormula());
        assertEquals("California", countyPopulations.getSheetName());

        Name employeeNames = workbook.getName("employee_names");
        assertNotNull(employeeNames);
        assertEquals("employee_names", employeeNames.getNameName());
        assertEquals("Employees!A3:A6", employeeNames.getRefersToFormula());
        assertEquals("Employees", employeeNames.getSheetName());

        Name employeeSalaries = workbook.getName("employee_salaries");
        assertNotNull(employeeSalaries);
        assertEquals("employee_salaries", employeeSalaries.getNameName());
        assertEquals("Employees!B3:B6", employeeSalaries.getRefersToFormula());
        assertEquals("Employees", employeeSalaries.getSheetName());
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
        beans.putAll(TestUtility.getSpecificStateData(0, "state"));
        beans.putAll(TestUtility.getEmployeeData());
        return beans;
    }
}