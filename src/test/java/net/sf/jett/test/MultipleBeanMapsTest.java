package net.sf.jett.test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the ability of <code>ExcelTransformer</code> to
 * supply different bean maps to cloned <code>Sheets</code>.
 *
 * @author Randy Gettman
 */
public class MultipleBeanMapsTest extends TestCase
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
     * Tests the .xls template spreadsheet.
     * @throws java.io.IOException If an I/O error occurs.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If the input spreadsheet is invalid.
     * @since 0.2.0
     */
    @Override
    @Test
    public void testXlsFiles() throws IOException, InvalidFormatException
    {
        super.testXlsFiles();
    }

    /**
     * Tests the .xlsx template spreadsheet.
     * @throws IOException If an I/O error occurs.
     * @throws InvalidFormatException If the input spreadsheet is invalid.
     * @since 0.2.0
     */
    @Override
    @Test
    public void testXlsxFiles() throws IOException, InvalidFormatException
    {
        super.testXlsxFiles();
    }

    /**
     * Returns the Excel name base for the template and resultant spreadsheets
     * for this test.
     * @return The Excel name base for this test.
     */
    @Override
    protected String getExcelNameBase()
    {
        return "MultipleBeanMaps";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet atlantic = workbook.getSheetAt(0);
        assertEquals("Atlantic", atlantic.getSheetName());
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(atlantic, 0, 0));
        assertEquals("Boston", TestUtility.getStringCellValue(atlantic, 2, 0));
        Sheet central = workbook.getSheetAt(1);
        assertEquals("Central", central.getSheetName());
        assertEquals("Division: Central", TestUtility.getStringCellValue(central, 0, 0));
        assertEquals("Chicago", TestUtility.getStringCellValue(central, 2, 0));
        Sheet southeast = workbook.getSheetAt(2);
        assertEquals("Southeast", southeast.getSheetName());
        assertEquals("Division: Southeast", TestUtility.getStringCellValue(southeast, 0, 0));
        assertEquals("Miami", TestUtility.getStringCellValue(southeast, 2, 0));
        Sheet northwest = workbook.getSheetAt(3);
        assertEquals("Northwest", northwest.getSheetName());
        assertEquals("Division: Northwest", TestUtility.getStringCellValue(northwest, 0, 0));
        assertEquals("Oklahoma City", TestUtility.getStringCellValue(northwest, 2, 0));
        Sheet pacific = workbook.getSheetAt(4);
        assertEquals("Pacific", pacific.getSheetName());
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(pacific, 0, 0));
        assertEquals("Los Angeles", TestUtility.getStringCellValue(pacific, 2, 0));
        Sheet southwest = workbook.getSheetAt(5);
        assertEquals("Southwest", southwest.getSheetName());
        assertEquals("Division: Southwest", TestUtility.getStringCellValue(southwest, 0, 0));
        assertEquals("San Antonio", TestUtility.getStringCellValue(southwest, 2, 0));
        Sheet empty = workbook.getSheetAt(6);
        assertEquals("Empty", empty.getSheetName());
        assertEquals("Division: Empty", TestUtility.getStringCellValue(empty, 0, 0));
        assertTrue(TestUtility.isCellBlank(empty, 2, 0));
        Sheet ofTheirOwn = workbook.getSheetAt(7);
        assertEquals("Of Their Own", ofTheirOwn.getSheetName());
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(ofTheirOwn, 0, 0));
        assertEquals("Harlem", TestUtility.getStringCellValue(ofTheirOwn, 2, 0));
    }

    /**
     * This test is a multiple beans map test.
     * @return <code>true</code>.
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
        String[] templateSheetNameArray = new String[8];
        Arrays.fill(templateSheetNameArray, "Division");
        return Arrays.asList(templateSheetNameArray);
    }

    /**
     * For multiple beans map tests, return the <code>List</code> of result
     * sheet names.
     * @return A <code>List</code> of result sheet names.
     */
    @Override
    protected List<String> getListOfResultSheetNames()
    {
        return Arrays.asList("Atlantic", "Central", "Southeast", "Northwest", "Pacific", "Southwest", "Empty", "Of Their Own");
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
        for (int i = 0; i < 8; i++)
            beansList.add(TestUtility.getSpecificDivisionData(i));
        return beansList;
    }
}
