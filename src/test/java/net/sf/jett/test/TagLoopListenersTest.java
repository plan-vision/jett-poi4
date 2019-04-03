package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.test.model.BlockShadingLoopListener;

/**
 * This JUnit Test class tests the processing of "onLoopProcessed" on a looping
 * tag.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class TagLoopListenersTest extends TestCase
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
        return "TagLoopListeners";
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
            for (int c = 0; c < 4; c++)
            {
                for (int r = 1; r < 9; r++)
                {
                    //System.err.println("TagLoopListeners: Testing s: " + s + ", c: " + c + ", r: " + r);
                    if ((r - 1) / 2 % 2 == 0)
                    {
                        // r is 1, 2, 5, 6
                        assertEquals(FillPatternType.NO_FILL, TestUtility.getCellFillPattern(sheet, r, c));
                    }
                    else
                    {
                        // r is 3, 4, 7, 8
                        assertEquals(FillPatternType.SOLID_FOREGROUND, TestUtility.getCellFillPattern(sheet, r, c));
                        CellStyle cs = TestUtility.getCellStyle(sheet, r, c);
                        assertNotNull(cs);
                        assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), cs.getFillForegroundColor());
                    }
                }
            }
        }

        Sheet before = workbook.getSheetAt(4);
        assertEquals("Three!", TestUtility.getStringCellValue(before, 0, 3));
        assertEquals("The above count, using ${x}, should have 3 replaced!", TestUtility.getStringCellValue(before, 1, 1));
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
        beans.put("blockShadingLoopListener", new BlockShadingLoopListener());
        return beans;
    }
}