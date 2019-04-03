package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.junit.Test;
import static org.junit.Assert.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.test.model.TestFuncs;
import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the pass-through of namespace objects to the
 * JEXL Engine.
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class NamespaceFuncsTest extends TestCase
{
    /**
     * Ensure that one cannot re-register a namespace.
     */
    @Test(expected=IllegalArgumentException.class)
    public void testReRegistration()
    {
        ExcelTransformer transformer = new ExcelTransformer();
        transformer.registerFuncs("jagg", Math.class);
    }

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
        return "NamespaceFuncs";
    }

    /**
     * Register "math" and "test".
     * @param transformer The <code>ExcelTransformer</code>.
     * @since 0.9.0
     */
    @Override
    protected void setupTransformer(ExcelTransformer transformer)
    {
        transformer.registerFuncs("math", Math.class);
        TestFuncs theTestFuncs = new TestFuncs();
        transformer.registerFuncs("test", theTestFuncs);
        transformer.setCache(10);
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet sheet = workbook.getSheetAt(0);
        assertEquals(-1, TestUtility.getNumericCellValue(sheet, 0, 1), DELTA);
        assertEquals(TestFuncs.THE_ANSWER, TestUtility.getNumericCellValue(sheet, 1, 1), DELTA);
        assertEquals(TestFuncs.THE_ANSWER, TestUtility.getNumericCellValue(sheet, 2, 1), DELTA);
        // The JEXL Engine cache is a parse cache, mapping the expression String
        // to an ASTJexlScript.  The interpretation on the ASTJexlScript still
        // occurs, so the method is still called twice, but the script is only
        // parsed once.
        //int numCalls = theTestFuncs.getCalls();
        //theTestFuncs.resetCalls();
        //assertEquals(1, numCalls);
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
