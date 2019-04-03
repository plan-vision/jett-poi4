package net.sf.jett.test;

import java.io.IOException;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.transform.ExcelTransformer;

/**
 * This JUnit Test class tests the fixed collection name feature of JETT.
 *
 * @author Randy Gettman
 */
public class FixedCollectionsTest extends TestCase
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
        return "FixedCollections";
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
        transformer.addFixedSizeCollectionName("division.teams");  // covers explicit forEach case
        transformer.addFixedSizeCollectionName("divisionsList.teams");  // covers implicit case
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet explicit = workbook.getSheetAt(0);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(explicit, 0, 0));
        assertEquals("Philadelphia", TestUtility.getStringCellValue(explicit, 3, 0));
        assertEquals("Raptors", TestUtility.getStringCellValue(explicit, 6, 1));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(explicit, 28, 0));
        assertTrue(TestUtility.isMergedRegionPresent(explicit, new CellRangeAddress(28, 28, 0, 4)));
        assertTrue(TestUtility.isCellBlank(explicit, 44, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 45, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 46, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 47, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 48, 0));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(explicit, 49, 0));
        assertEquals("Harlem", TestUtility.getStringCellValue(explicit, 51, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 52, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 53, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 54, 0));
        assertTrue(TestUtility.isCellBlank(explicit, 55, 0));
        assertEquals("After", TestUtility.getStringCellValue(explicit, 56, 0));
        assertEquals(8, explicit.getNumMergedRegions());

        Sheet implicit = workbook.getSheetAt(1);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(implicit, 0, 0));
        assertEquals("Philadelphia", TestUtility.getStringCellValue(implicit, 3, 0));
        assertEquals("Raptors", TestUtility.getStringCellValue(implicit, 6, 1));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(implicit, 28, 0));
        assertTrue(TestUtility.isMergedRegionPresent(implicit, new CellRangeAddress(28, 28, 0, 4)));
        assertTrue(TestUtility.isCellBlank(implicit, 44, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 45, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 46, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 47, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 48, 0));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(implicit, 49, 0));
        assertEquals("Harlem", TestUtility.getStringCellValue(implicit, 51, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 52, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 53, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 54, 0));
        assertTrue(TestUtility.isCellBlank(implicit, 55, 0));
        assertEquals("After", TestUtility.getStringCellValue(implicit, 56, 0));
        assertEquals(8, implicit.getNumMergedRegions());
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
        return TestUtility.getDivisionData();
    }
}
