package net.sf.jett.test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import net.sf.jett.event.CellEvent;
import net.sf.jett.event.CellListener;
import net.sf.jett.transform.ExcelTransformer;
import net.sf.jett.util.SheetUtil;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests whether extraneous blank cells get created in a
 * template full of looping tags.
 *
 * @author Randy Gettman
 * @since 0.11.0
 */
public class ExtraneousBlankCellsTest extends TestCase
{
    private static final int SCALE = 3;

    /**
     * Tests the .xls template spreadsheet.
     * @throws IOException If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
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
    protected String getExcelNameBase() { return "ExtraneousBlankCells"; }

    /**
     * Call certain setup-related methods on the <code>ExcelTransformer</code>
     * before template sheet transformation.
     * @param transformer The <code>ExcelTransformer</code> that will transform
     *    the template worksheet(s).
     */
    @Override
    protected void setupTransformer(ExcelTransformer transformer)
    {
        // Make the phantom blank cells appear!
        transformer.addCellListener(new CellListener()
        {
            @Override
            public boolean beforeCellProcessed(CellEvent event)
            {
                return true;
            }

            @Override
            public void cellProcessed(CellEvent event)
            {
                Cell cell = event.getCell();
                cell.setCellValue(cell.getStringCellValue() + "!!!");
            }
        });
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
        assertEquals(SCALE + 2, SheetUtil.getLastPopulatedColIndex(sheet));
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
        beans.put("it1", listOf("i", SCALE));
        beans.put("it2", listOf("j", SCALE));
        beans.put("it3", listOf("k", SCALE));
        beans.put("it4", listOf("m", SCALE));

        return beans;
    }

    /**
     * Makes quick lists such as {"a1", "a2", ..., "an"}
     * @param prefix The beginning of the string.
     * @param num Controls how many items.
     * @return A <code>List</code> of <code>Strings</code>.
     */
    private static List<String> listOf(String prefix, int num)
    {
        List<String> result = new ArrayList<>(num);
        for (int i = 1; i <= num; i++)
        {
            result.add(String.format("%s%d", prefix, i));
        }
        return result;
    }
}
