package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.junit.Test;
import static org.junit.Assert.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.test.model.BoldTagListener;
import net.sf.jett.test.model.ItalicTagListener;

/**
 * This JUnit Test class tests the processing of "onProcessed" on any tags.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class TagListenersTest extends TestCase
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
        return "TagListeners";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        for (int s = 0; s < 4; s++)
        {
            Sheet sheet = workbook.getSheetAt(s);
            for (int r = 1; r <= 8; r++)
            {
                Row row = sheet.getRow(r);
                for (int c = 0; c <= 3; c++)
                {
                    Cell cell = row.getCell(c);
                    Font f = workbook.getFontAt(cell.getCellStyle().getFontIndex());
                    assertEquals(true, f.getBold());
                    boolean shouldBeItalic = ((s == 0 || s == 1) &&
                            (r == 4 || r == 6) &&
                            (c == 3));
                    assertEquals(shouldBeItalic, f.getItalic());
                }
            }
        }

        Sheet before = workbook.getSheetAt(4);
        assertEquals(4, TestUtility.getNumericCellValue(before, 0, 1), TestCase.DELTA);
        assertEquals("The above will be replaced by ${employees.size()}",
                TestUtility.getStringCellValue(before, 1, 1));

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
        Map<String, Object> beans = TestUtility.getEmployeeData();
        beans.put("boldTagListener", new BoldTagListener());
        beans.put("italicTagListener", new ItalicTagListener());
        return beans;
    }
}
