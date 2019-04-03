package net.sf.jett.test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the behavior of JETT Formulas and when sheet
 * names change, either as part of the transformation and cloning process, or
 * as part of the evaluation of an expression in the sheet name.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class FormulasAndClonedSheetsTest extends TestCase
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
        return "FormulasAndClonedSheets";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        // State specific.
        Sheet california = workbook.getSheetAt(0);
        assertEquals("SUM(B3:B60)", TestUtility.getFormulaCellValue(california, 60, 1));
        assertEquals("SUM(C3:C60)", TestUtility.getFormulaCellValue(california, 60, 2));
        assertEquals("\"Counties:\"&COUNTA(E3:E60)", TestUtility.getFormulaCellValue(california, 60, 4));
        assertEquals("SUM(California!B3:B60,Nevada!B3:B19)", TestUtility.getFormulaCellValue(california, 61, 1));
        assertEquals("SUM(California!C3:C60,Nevada!C3:C19)", TestUtility.getFormulaCellValue(california, 61, 2));
        assertEquals("\"Counties:\"&COUNTA(California!E3:E60,Nevada!E3:E19)", TestUtility.getFormulaCellValue(california, 61, 4));

        Sheet nevada = workbook.getSheetAt(1);
        assertEquals("SUM(B3:B19)", TestUtility.getFormulaCellValue(nevada, 19, 1));
        assertEquals("SUM(C3:C19)", TestUtility.getFormulaCellValue(nevada, 19, 2));
        assertEquals("\"Counties:\"&COUNTA(E3:E19)", TestUtility.getFormulaCellValue(nevada, 19, 4));
        assertEquals("SUM(California!B3:B60,Nevada!B3:B19)", TestUtility.getFormulaCellValue(nevada, 20, 1));
        assertEquals("SUM(California!C3:C60,Nevada!C3:C19)", TestUtility.getFormulaCellValue(nevada, 20, 2));
        assertEquals("\"Counties:\"&COUNTA(California!E3:E60,Nevada!E3:E19)", TestUtility.getFormulaCellValue(nevada, 20, 4));

        // Common to all states.
        for (int s = 0; s < 2; s++)
        {
            Sheet state = workbook.getSheetAt(s);
            assertEquals("COUNTA(Atlantic!B3:B7,Central!B3:B7,Northwest!B3:B7,Pacific!B3:B7,Southeast!B3:B7,Southwest!B3:B7)",
                    TestUtility.getFormulaCellValue(state, 2, 7));
            assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7)",
                    TestUtility.getFormulaCellValue(state, 2, 8));
            assertEquals("SUM(Atlantic!D3:D7,Central!D3:D7,Northwest!D3:D7,Pacific!D3:D7,Southeast!D3:D7,Southwest!D3:D7)",
                    TestUtility.getFormulaCellValue(state, 2, 9));
            assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7,Atlantic!D3:D7,Central!D3:D7,Northwest!D3:D7,Pacific!D3:D7,Southeast!D3:D7,Southwest!D3:D7)",
                    TestUtility.getFormulaCellValue(state, 2, 10));
            assertEquals("SUM(I3)/SUM(K3)", TestUtility.getFormulaCellValue(state, 2, 11));
            assertEquals("\"SUM:\"&SUM(Expressions!B2:B11)", TestUtility.getFormulaCellValue(state, 1, 13));
            assertEquals("\"SUM:\"&SUM(Expressions!C2:C11)", TestUtility.getFormulaCellValue(state, 1, 14));
        }

        Sheet expressions = workbook.getSheetAt(2);
        assertEquals("\"SUM:\"&SUM(B2:B11)", TestUtility.getFormulaCellValue(expressions, 11, 1));
        assertEquals("\"SUM:\"&SUM(C2:C11)", TestUtility.getFormulaCellValue(expressions, 11, 2));
        assertEquals("\"SUM:\"&SUM(Expressions!B2:B11)", TestUtility.getFormulaCellValue(expressions, 12, 1));
        assertEquals("\"SUM:\"&SUM(Expressions!C2:C11)", TestUtility.getFormulaCellValue(expressions, 12, 2));
        assertEquals("SUM(California!B3:B60,Nevada!B3:B19)", TestUtility.getFormulaCellValue(expressions, 2, 5));
        assertEquals("SUM(California!C3:C60,Nevada!C3:C19)", TestUtility.getFormulaCellValue(expressions, 2, 6));
        assertEquals("\"Counties:\"&COUNTA(California!E3:E60,Nevada!E3:E19)", TestUtility.getFormulaCellValue(expressions, 2, 7));
        assertEquals("COUNTA(Atlantic!B3:B7,Central!B3:B7,Northwest!B3:B7,Pacific!B3:B7,Southeast!B3:B7,Southwest!B3:B7)",
                TestUtility.getFormulaCellValue(expressions, 2, 9));
        assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7)",
                TestUtility.getFormulaCellValue(expressions, 2, 10));
        assertEquals("SUM(Atlantic!D3:D7,Central!D3:D7,Northwest!D3:D7,Pacific!D3:D7,Southeast!D3:D7,Southwest!D3:D7)",
                TestUtility.getFormulaCellValue(expressions, 2, 11));
        assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7,Atlantic!D3:D7,Central!D3:D7,Northwest!D3:D7,Pacific!D3:D7,Southeast!D3:D7,Southwest!D3:D7)",
                TestUtility.getFormulaCellValue(expressions, 2, 12));
        assertEquals("SUM(K3)/SUM(M3)", TestUtility.getFormulaCellValue(expressions, 2, 13));

        // Divisions
        for (int s = 3; s < 9; s++)
        {
            Sheet division = workbook.getSheetAt(s);
            assertEquals("COUNTA(B3:B7)", TestUtility.getFormulaCellValue(division, 7, 1));
            assertEquals("SUM(C3:C7)", TestUtility.getFormulaCellValue(division, 7, 2));
            assertEquals("SUM(D3:D7)", TestUtility.getFormulaCellValue(division, 7, 3));
            assertEquals("SUM(C3:C7,D3:D7)", TestUtility.getFormulaCellValue(division, 7, 4));
            assertEquals("SUM(C3:C7)/SUM(E3:E7)", TestUtility.getFormulaCellValue(division, 7, 5));
            assertEquals("COUNTA(Atlantic!B3:B7,Central!B3:B7,Northwest!B3:B7,Pacific!B3:B7,Southeast!B3:B7,Southwest!B3:B7)",
                    TestUtility.getFormulaCellValue(division, 8, 1));
            assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7)",
                    TestUtility.getFormulaCellValue(division, 8, 2));
            assertEquals("SUM(Atlantic!D3:D7,Central!D3:D7,Northwest!D3:D7,Pacific!D3:D7,Southeast!D3:D7,Southwest!D3:D7)",
                    TestUtility.getFormulaCellValue(division, 8, 3));
            assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7,Atlantic!D3:D7,Central!D3:D7,Northwest!D3:D7,Pacific!D3:D7,Southeast!D3:D7,Southwest!D3:D7)",
                    TestUtility.getFormulaCellValue(division, 8, 4));
            assertEquals("SUM(Atlantic!C3:C7,Central!C3:C7,Northwest!C3:C7,Pacific!C3:C7,Southeast!C3:C7,Southwest!C3:C7)/SUM(Atlantic!E3:E7,Central!E3:E7,Northwest!E3:E7,Pacific!E3:E7,Southeast!E3:E7,Southwest!E3:E7)",
                    TestUtility.getFormulaCellValue(division, 8, 5));
            assertEquals("\"SUM:\"&SUM(Expressions!B2:B11)", TestUtility.getFormulaCellValue(division, 1, 12));
            assertEquals("\"SUM:\"&SUM(Expressions!C2:C11)", TestUtility.getFormulaCellValue(division, 1, 13));
        }
    }

    /**
     * This test is a single map test.
     * @return <code>false</code>.
     */
    @Override
    protected boolean isMultipleBeans()
    {
        return true;
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of template
     * sheet names.
     * @return A <code>List</code> of template sheet names.
     */
    @Override
    protected List<String> getListOfTemplateSheetNames()
    {
        return Arrays.asList("stateTemplate", "stateTemplate", "expressionTemplate", "divisionTemplate",
                "divisionTemplate", "divisionTemplate", "divisionTemplate", "divisionTemplate", "divisionTemplate");
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of result
     * sheet names.
     * @return A <code>List</code> of result sheet names.
     */
    @Override
    protected List<String> getListOfResultSheetNames()
    {
        return Arrays.asList("California", "Nevada", "${expressionsSheetName}", "Atlantic",
                "Central", "Southeast", "Northwest", "Pacific", "Southwest");
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of beans maps,
     * which map bean names to bean values for each corresponding sheet.
     * @return A <code>List</code> of <code>Maps</code> of bean names to bean
     *    values.
     */
    @Override
    protected List<Map<String, Object>> getListOfBeansMaps()
    {
        List<Map<String, Object>> beansList = new ArrayList<>();
        // For states.
        beansList.add(TestUtility.getSpecificStateData(0, "state"));
        beansList.add(TestUtility.getSpecificStateData(1, "state"));
        // For "Expressions".
        Map<String, Object> expressionsBeans = new HashMap<>();
        expressionsBeans.put("expressionsSheetName", "Expressions");
        beansList.add(expressionsBeans);
        // For divisions.
        for (int i = 0; i < 8; i++)
            beansList.add(TestUtility.getSpecificDivisionData(i));

        return beansList;
    }
}
