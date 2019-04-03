package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.test.model.AreaCellListener;
import net.sf.jett.test.model.PopulationCellListener;
import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the <code>CellListener</code> feature of JETT.
 *
 * @author Randy Gettman
 */
public class CellListenerTest extends TestCase
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
        return "CellListener";
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
        transformer.addCellListener(new PopulationCellListener(1000000));
        transformer.addCellListener(new AreaCellListener(10000));
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet cellListener = workbook.getSheetAt(0);
        for (int i = 0; i < 58; i++)
        {
            Row row = cellListener.getRow(i + 2);
            Cell caPop = row.getCell(1);
            short popFontIdx = caPop.getCellStyle().getFontIndex();
            Font popFont = workbook.getFontAt(popFontIdx);
            double population = caPop.getNumericCellValue();
            /*if (population >= 1000000)
                assertEquals("Expected bold font at row " + (i + 2) + ", cell 1",
                        Font.BOLDWEIGHT_BOLD, popFont.getBoldweight());
            else
                assertEquals("Expected not bold font at row " + (i + 2) + ", cell 1",
                        Font.BOLDWEIGHT_NORMAL, popFont.getBoldweight());*/
            Cell caArea = row.getCell(2);
            short areaFontIdx = caArea.getCellStyle().getFontIndex();
            Font areaFont = workbook.getFontAt(areaFontIdx);
            double area = caArea.getNumericCellValue();
            if (area >= 10000)
                assertTrue("Expected italic font at row " + (i + 2) + ", cell 2",
                        areaFont.getItalic());
            else
                assertFalse("Expected not italic font at row " + (i + 2) + ", cell 2",
                        areaFont.getItalic());
        }
        for (int i = 0; i < 17; i++)
        {
            Row row = cellListener.getRow(i + 2);
            Cell nvPop = row.getCell(7);
            short popFontIdx = nvPop.getCellStyle().getFontIndex();
            Font popFont = workbook.getFontAt(popFontIdx);
            double population = nvPop.getNumericCellValue();
            /*if (population >= 1000000)
                assertEquals("Expected bold font at row " + (i + 2) + ", cell 7",
                        Font.BOLDWEIGHT_BOLD, popFont.getBoldweight());
            else
                assertEquals("Expected not bold font at row " + (i + 2) + ", cell 7",
                        Font.BOLDWEIGHT_NORMAL, popFont.getBoldweight());*/
            Cell nvArea = row.getCell(8);
            short areaFontIdx = nvArea.getCellStyle().getFontIndex();
            Font areaFont = workbook.getFontAt(areaFontIdx);
            double area = nvArea.getNumericCellValue();
            if (area >= 10000)
                assertTrue("Expected italic font at row " + (i + 2) + ", cell 8",
                        areaFont.getItalic());
            else
                assertFalse("Expected not italic font at row " + (i + 2) + ", cell 8",
                        areaFont.getItalic());
        }

        Sheet before = workbook.getSheetAt(1);
        assertEquals("California", TestUtility.getStringCellValue(before, 0, 1));
        assertEquals("The CellListener will replace the above content with ${california.name}",
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
        return TestUtility.getStateData();
    }
}
