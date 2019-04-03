package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.test.model.CustomTagLibrary;
import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the custom tag library feature of JETT.
 *
 * @author Randy Gettman
 */
public class CustomTagLibraryTest extends TestCase
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
        return "CustomTagLibrary";
    }

    /**
     * Call certain setup-related methods on the <code>ExcelTransformer</code>
     * before template sheet transformation.
     * @param transformer The <code>ExcelTransformer</code> that will transform
     *    the template worksheet(s).
     */
    @Override
    protected void setupTransformer(ExcelTransformer transformer)
    {
        transformer.registerTagLibrary("custom", CustomTagLibrary.getCustomTagLibrary());
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet custom = workbook.getSheetAt(0);
        assertEquals(6, TestUtility.getNumericCellValue(custom, 0, 1), DELTA);
        assertEquals(4.14, TestUtility.getNumericCellValue(custom, 1, 1), DELTA);
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
        beans.put("num", 3.14);
        return beans;
    }
}
