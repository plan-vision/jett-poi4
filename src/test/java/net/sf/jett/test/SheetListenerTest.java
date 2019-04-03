package net.sf.jett.test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.test.model.DemoSheetListener;
import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the <code>SheetListener</code> feature of JETT.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class SheetListenerTest extends TestCase
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
        return "SheetListener";
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
        transformer.addSheetListener(new DemoSheetListener());
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet first = workbook.getSheetAt(0);
        assertEquals("Message changed by a SheetListener!", TestUtility.getStringCellValue(first, 1, 1));
        assertEquals("Sheet Listener Message!", TestUtility.getStringCellValue(first, 2, 1));
        assertEquals("Changed by DemoSheetListener!", TestUtility.getStringCellValue(first, 0, 5));

        Sheet second = workbook.getSheetAt(1);
        assertEquals("${message2}", TestUtility.getStringCellValue(second, 1, 1));
        assertEquals("${message}", TestUtility.getStringCellValue(second, 2, 1));
        assertTrue(TestUtility.isCellBlank(second, 0, 5));
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
        beans.put("message", "Sheet Listener Message!");
        beans.put("message2", "Message changed by a SheetListener!");
        return beans;
    }
}