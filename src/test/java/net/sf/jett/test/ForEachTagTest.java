package net.sf.jett.test;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 * This JUnit Test class tests the evaluation of the "forEach" tag in entire
 * rows, block area, and bodiless modes.
 *
 * @author Randy Gettman
 */
public class ForEachTagTest extends TestCase
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
        return "ForEachTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet vertVert = workbook.getSheetAt(0);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(vertVert, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertVert, new CellRangeAddress(0, 0, 0, 4)));
        assertEquals("Boston", TestUtility.getStringCellValue(vertVert, 2, 0));
        assertEquals("Raptors", TestUtility.getStringCellValue(vertVert, 6, 1));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(vertVert, 28, 0));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(vertVert, 29, 4).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(vertVert, 30, 2), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(vertVert, 31, 3), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(vertVert, 32, 4), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(vertVert, 42, 0));
        assertEquals("City", TestUtility.getStringCellValue(vertVert, 43, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertVert, new CellRangeAddress(44, 44, 0, 4)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(vertVert, 44, 0));
        assertEquals("Name", TestUtility.getStringCellValue(vertVert, 45, 1));
        assertEquals("Harlem", TestUtility.getStringCellValue(vertVert, 46, 0));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(vertVert, 46, 1).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(vertVert, 47, 0));
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(vertVert, 48, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertVert, new CellRangeAddress(48, 48, 0, 4)));
        assertEquals("Boston", TestUtility.getStringCellValue(vertVert, 50, 0));
        assertEquals("Raptors", TestUtility.getStringCellValue(vertVert, 54, 1));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(vertVert, 76, 0));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(vertVert, 77, 4).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(vertVert, 78, 2), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(vertVert, 79, 3), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(vertVert, 80, 4), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(vertVert, 90, 0));
        assertEquals("City", TestUtility.getStringCellValue(vertVert, 91, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertVert, new CellRangeAddress(92, 92, 0, 4)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(vertVert, 92, 0));
        assertEquals("Name", TestUtility.getStringCellValue(vertVert, 93, 1));
        assertEquals("Harlem", TestUtility.getStringCellValue(vertVert, 94, 0));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(vertVert, 94, 1).getCellStyle().getFillForegroundColor());
        assertEquals("After2", TestUtility.getStringCellValue(vertVert, 95, 0));
        assertEquals(16, vertVert.getNumMergedRegions());

        Sheet horizVert = workbook.getSheetAt(1);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(horizVert, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horizVert, new CellRangeAddress(0, 0, 0, 4)));
        assertEquals("Boston", TestUtility.getStringCellValue(horizVert, 2, 0));
        assertEquals("Raptors", TestUtility.getStringCellValue(horizVert, 6, 1));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(horizVert, 0, 20));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(horizVert, 1, 24).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(horizVert, 2, 22), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(horizVert, 3, 23), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(horizVert, 4, 24), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(horizVert, 0, 30));
        assertEquals("City", TestUtility.getStringCellValue(horizVert, 1, 30));
        assertTrue(TestUtility.isMergedRegionPresent(horizVert, new CellRangeAddress(0, 0, 35, 39)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(horizVert, 0, 35));
        assertEquals("Name", TestUtility.getStringCellValue(horizVert, 1, 36));
        assertEquals("Harlem", TestUtility.getStringCellValue(horizVert, 2, 35));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(horizVert, 2, 35).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(horizVert, 0, 40));
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(horizVert, 9, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horizVert, new CellRangeAddress(9, 9, 0, 4)));
        assertEquals("Boston", TestUtility.getStringCellValue(horizVert, 11, 0));
        assertEquals("Raptors", TestUtility.getStringCellValue(horizVert, 15, 1));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(horizVert, 9, 20));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(horizVert, 10, 24).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(horizVert, 11, 22), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(horizVert, 12, 23), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(horizVert, 13, 24), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(horizVert, 9, 30));
        assertEquals("City", TestUtility.getStringCellValue(horizVert, 10, 30));
        assertTrue(TestUtility.isMergedRegionPresent(horizVert, new CellRangeAddress(9, 9, 35, 39)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(horizVert, 9, 35));
        assertEquals("Name", TestUtility.getStringCellValue(horizVert, 10, 36));
        assertEquals("Harlem", TestUtility.getStringCellValue(horizVert, 11, 35));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(horizVert, 11, 35).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(horizVert, 9, 40));
        assertEquals(16, horizVert.getNumMergedRegions());

        Sheet vertHoriz = workbook.getSheetAt(2);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(vertHoriz, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertHoriz, new CellRangeAddress(0, 4, 0, 0)));
        assertEquals("Boston", TestUtility.getStringCellValue(vertHoriz, 0, 2));
        assertEquals("Raptors", TestUtility.getStringCellValue(vertHoriz, 1, 6));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(vertHoriz, 20, 0));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(vertHoriz, 24, 1).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(vertHoriz, 22, 2), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(vertHoriz, 23, 3), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(vertHoriz, 24, 4), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(vertHoriz, 30, 0));
        assertEquals("City", TestUtility.getStringCellValue(vertHoriz, 30, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vertHoriz, new CellRangeAddress(35, 39, 0, 0)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(vertHoriz, 35, 0));
        assertEquals("Name", TestUtility.getStringCellValue(vertHoriz, 36, 1));
        assertEquals("Harlem", TestUtility.getStringCellValue(vertHoriz, 35, 2));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(vertHoriz, 35, 2).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(vertHoriz, 40, 0));
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(vertHoriz, 0, 9));
        assertTrue(TestUtility.isMergedRegionPresent(vertHoriz, new CellRangeAddress(0, 4, 9, 9)));
        assertEquals("Boston", TestUtility.getStringCellValue(vertHoriz, 0, 11));
        assertEquals("Raptors", TestUtility.getStringCellValue(vertHoriz, 1, 15));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(vertHoriz, 20, 9));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(vertHoriz, 24, 10).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(vertHoriz, 22, 11), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(vertHoriz, 23, 12), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(vertHoriz, 24, 13), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(vertHoriz, 30, 9));
        assertEquals("City", TestUtility.getStringCellValue(vertHoriz, 30, 10));
        assertTrue(TestUtility.isMergedRegionPresent(vertHoriz, new CellRangeAddress(35, 39, 9, 9)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(vertHoriz, 35, 9));
        assertEquals("Name", TestUtility.getStringCellValue(vertHoriz, 36, 10));
        assertEquals("Harlem", TestUtility.getStringCellValue(vertHoriz, 35, 11));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(vertHoriz, 35, 11).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(vertHoriz, 40, 9));
        assertEquals(16, vertHoriz.getNumMergedRegions());

        Sheet horizHoriz = workbook.getSheetAt(3);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(horizHoriz, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horizHoriz, new CellRangeAddress(0, 4, 0, 0)));
        assertEquals("Boston", TestUtility.getStringCellValue(horizHoriz, 0, 2));
        assertEquals("Raptors", TestUtility.getStringCellValue(horizHoriz, 1, 6));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(horizHoriz, 0, 28));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(horizHoriz, 4, 29).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(horizHoriz, 2, 30), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(horizHoriz, 3, 31), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(horizHoriz, 4, 32), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(horizHoriz, 0, 42));
        assertEquals("City", TestUtility.getStringCellValue(horizHoriz, 0, 43));
        assertTrue(TestUtility.isMergedRegionPresent(horizHoriz, new CellRangeAddress(0, 4, 44, 44)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(horizHoriz, 0, 44));
        assertEquals("Name", TestUtility.getStringCellValue(horizHoriz, 1, 45));
        assertEquals("Harlem", TestUtility.getStringCellValue(horizHoriz, 0, 46));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(horizHoriz, 1, 46).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(horizHoriz, 0, 47));
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(horizHoriz, 0, 48));
        assertTrue(TestUtility.isMergedRegionPresent(horizHoriz, new CellRangeAddress(0, 4, 48, 48)));
        assertEquals("Boston", TestUtility.getStringCellValue(horizHoriz, 0, 50));
        assertEquals("Raptors", TestUtility.getStringCellValue(horizHoriz, 1, 54));
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(horizHoriz, 0, 76));
        // getFillForegroundColor returns zero for XSSFCellStyles!!!
        //assertEquals(IndexedColors.PALE_BLUE.getIndex(), getCell(horizHoriz, 4, 77).getCellStyle().getFillForegroundColor());
        assertEquals(53, TestUtility.getNumericCellValue(horizHoriz, 2, 78), DELTA);
        assertEquals(36, TestUtility.getNumericCellValue(horizHoriz, 3, 79), DELTA);
        assertEquals((double) 32/74, TestUtility.getNumericCellValue(horizHoriz, 4, 80), DELTA);
        assertEquals("Division: Empty", TestUtility.getStringCellValue(horizHoriz, 0, 90));
        assertEquals("City", TestUtility.getStringCellValue(horizHoriz, 0, 91));
        assertTrue(TestUtility.isMergedRegionPresent(horizHoriz, new CellRangeAddress(0, 4, 92, 92)));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(horizHoriz, 0, 92));
        assertEquals("Name", TestUtility.getStringCellValue(horizHoriz, 1, 93));
        assertEquals("Harlem", TestUtility.getStringCellValue(horizHoriz, 0, 94));
        //assertEquals(IndexedColors.GREY_25_PERCENT.getIndex(), getCell(horizHoriz, 1, 94).getCellStyle().getFillForegroundColor());
        assertEquals("After", TestUtility.getStringCellValue(horizHoriz, 0, 95));
        assertEquals(16, horizHoriz.getNumMergedRegions());

        Sheet indexVar = workbook.getSheetAt(4);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(indexVar, 0, 0));
        assertEquals("1.", TestUtility.getStringCellValue(indexVar, 2, 0));
        assertEquals("2.", TestUtility.getStringCellValue(indexVar, 3, 0));
        assertEquals("3.", TestUtility.getStringCellValue(indexVar, 4, 0));
        assertEquals("4.", TestUtility.getStringCellValue(indexVar, 5, 0));
        assertEquals("5.", TestUtility.getStringCellValue(indexVar, 6, 0));
        assertEquals("1.", TestUtility.getStringCellValue(indexVar, 9, 0));
        assertEquals("2.", TestUtility.getStringCellValue(indexVar, 10, 0));
        assertEquals("3.", TestUtility.getStringCellValue(indexVar, 11, 0));
        assertEquals("4.", TestUtility.getStringCellValue(indexVar, 12, 0));
        assertEquals("5.", TestUtility.getStringCellValue(indexVar, 13, 0));
        assertEquals("1.", TestUtility.getStringCellValue(indexVar, 46, 0));

        Sheet where = workbook.getSheetAt(5);
        assertEquals("Boston", TestUtility.getStringCellValue(where, 2, 0));
        assertEquals("Philadelphia", TestUtility.getStringCellValue(where, 3, 0));
        assertEquals("Division: Central - Teams Above 0.500", TestUtility.getStringCellValue(where, 4, 0));
        assertEquals("Chicago", TestUtility.getStringCellValue(where, 6, 0));
        assertEquals("Miami", TestUtility.getStringCellValue(where, 9, 0));
        assertEquals("Atlanta", TestUtility.getStringCellValue(where, 11, 0));
        assertEquals("Oklahoma City", TestUtility.getStringCellValue(where, 14, 0));
        assertEquals("Portland", TestUtility.getStringCellValue(where, 16, 0));
        assertEquals("Lakers", TestUtility.getStringCellValue(where, 19, 1));
        assertEquals("San Antonio", TestUtility.getStringCellValue(where, 22, 0));
        assertEquals("Houston", TestUtility.getStringCellValue(where, 26, 0));
        assertEquals("Division: Of Their Own - Teams Above 0.500", TestUtility.getStringCellValue(where, 27, 0));
        assertEquals("Harlem", TestUtility.getStringCellValue(where, 29, 0));
        assertEquals("After", TestUtility.getStringCellValue(where, 30, 0));

        Sheet limit = workbook.getSheetAt(6);
        assertEquals("Celtics", TestUtility.getStringCellValue(limit, 2, 1));
        assertEquals("Knicks", TestUtility.getStringCellValue(limit, 4, 1));
        assertEquals("Bulls", TestUtility.getStringCellValue(limit, 7, 1));
        assertEquals("Bucks", TestUtility.getStringCellValue(limit, 9, 1));
        assertEquals("Heat", TestUtility.getStringCellValue(limit, 12, 1));
        assertEquals("Hawks", TestUtility.getStringCellValue(limit, 14, 1));
        assertEquals("Thunder", TestUtility.getStringCellValue(limit, 17, 1));
        assertEquals("Trailblazers", TestUtility.getStringCellValue(limit, 19, 1));
        assertEquals("Lakers", TestUtility.getStringCellValue(limit, 22, 1));
        assertEquals("Warriors", TestUtility.getStringCellValue(limit, 24, 1));
        assertEquals("Spurs", TestUtility.getStringCellValue(limit, 27, 1));
        assertEquals("Hornets", TestUtility.getStringCellValue(limit, 29, 1));
        assertTrue(TestUtility.isCellBlank(limit, 32, 1));
        assertTrue(TestUtility.isCellBlank(limit, 33, 1));
        assertTrue(TestUtility.isCellBlank(limit, 34, 1));
        assertEquals("Globetrotters", TestUtility.getStringCellValue(limit, 37, 1));
        assertTrue(TestUtility.isCellBlank(limit, 38, 1));
        assertTrue(TestUtility.isCellBlank(limit, 39, 1));
        assertEquals("After", TestUtility.getStringCellValue(limit, 40, 0));

        Sheet groupRows = workbook.getSheetAt(7);
        for (int r = 0; r < 96; r++)
        {
            // These rows are collapsed.
            if ((r >=  9 && r <= 13) ||
                    (r >= 57 && r <= 61))
            {
                assertTrue(groupRows.getRow(r).getZeroHeight());
            }
            else
            {
                assertFalse(groupRows.getRow(r) != null && groupRows.getRow(r).getZeroHeight());
            }
        }

        Sheet groupCols = workbook.getSheetAt(8);
        for (int c = 0; c < 96; c++)
        {
            // These columns are collapsed.
            if ((c >= 23 && c <= 27) ||
                    (c >= 71 && c <= 75))
            {
                assertTrue(groupCols.isColumnHidden(c));
            }
            else
            {
                assertFalse(groupCols.isColumnHidden(c));
            }
        }

        // Note different order of divisions imposed by the "group by".
        // As of 0.8.0, this now tests dynamic properties from jAgg.
        // The "groupBy" attribute is "division_name", not "divisionName", so
        // that jAgg can still call "get("division_name")" and the JETT "group
        // by" operation still succeeds.
        Sheet groupBy = workbook.getSheetAt(9);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(groupBy, 0, 0));
        assertEquals("Boston", TestUtility.getStringCellValue(groupBy, 2, 0));
        assertEquals("Division: Central", TestUtility.getStringCellValue(groupBy, 7, 0));
        assertEquals("Bulls", TestUtility.getStringCellValue(groupBy, 9, 1));
        assertEquals("Division: Northwest", TestUtility.getStringCellValue(groupBy, 14, 0));
        assertEquals("Timberwolves", TestUtility.getStringCellValue(groupBy, 20, 1));
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(groupBy, 21, 0));
        assertEquals(21227, TestUtility.getNumericCellValue(groupBy, 23, 2), DELTA);
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(groupBy, 24, 0));
        assertEquals(52, TestUtility.getNumericCellValue(groupBy, 30, 3), DELTA);
        assertEquals("Division: Southeast", TestUtility.getStringCellValue(groupBy, 31, 0));
        assertEquals(30, TestUtility.getNumericCellValue(groupBy, 36, 2), DELTA);
        assertEquals("Division: Southwest", TestUtility.getStringCellValue(groupBy, 38, 0));
        assertEquals("Houston", TestUtility.getStringCellValue(groupBy, 44, 0));

        Sheet orderBy = workbook.getSheetAt(10);
        assertEquals("Southwest", TestUtility.getStringCellValue(orderBy, 1, 0));
        assertEquals(38.0 / (38 + 35), TestUtility.getNumericCellValue(orderBy, 1, 5), DELTA);
        assertEquals("Southwest", TestUtility.getStringCellValue(orderBy, 5, 0));
        assertEquals(57.0 / (57 + 16), TestUtility.getNumericCellValue(orderBy, 5, 5), DELTA);
        assertEquals("Southeast", TestUtility.getStringCellValue(orderBy, 6, 0));
        assertEquals(17.0 / (17 + 55), TestUtility.getNumericCellValue(orderBy, 6, 5), DELTA);
        assertEquals("Southeast", TestUtility.getStringCellValue(orderBy, 10, 0));
        assertEquals(51.0 / (51 + 22), TestUtility.getNumericCellValue(orderBy, 10, 5), DELTA);
        assertEquals("Pacific", TestUtility.getStringCellValue(orderBy, 11, 0));
        assertEquals(20.0 / (20 + 52), TestUtility.getNumericCellValue(orderBy, 11, 5), DELTA);
        assertEquals("Pacific", TestUtility.getStringCellValue(orderBy, 15, 0));
        assertEquals(53.0 / (53 + 20), TestUtility.getNumericCellValue(orderBy, 15, 5), DELTA);
        assertEquals("Of Their Own", TestUtility.getStringCellValue(orderBy, 16, 0));
        assertEquals(21227.0 / (21227 + 341), TestUtility.getNumericCellValue(orderBy, 16, 5), DELTA);
        assertEquals("Northwest", TestUtility.getStringCellValue(orderBy, 17, 0));
        assertEquals(17.0 / (17 + 57), TestUtility.getNumericCellValue(orderBy, 17, 5), DELTA);
        assertEquals("Northwest", TestUtility.getStringCellValue(orderBy, 21, 0));
        assertEquals(48.0 / (48 + 24), TestUtility.getNumericCellValue(orderBy, 21, 5), DELTA);
        assertEquals("Central", TestUtility.getStringCellValue(orderBy, 22, 0));
        assertEquals(14.0 / (14 + 58), TestUtility.getNumericCellValue(orderBy, 22, 5), DELTA);
        assertEquals("Central", TestUtility.getStringCellValue(orderBy, 26, 0));
        assertEquals(53.0 / (53 + 19), TestUtility.getNumericCellValue(orderBy, 26, 5), DELTA);
        assertEquals("Atlantic", TestUtility.getStringCellValue(orderBy, 27, 0));
        assertEquals(20.0 / (20 + 53), TestUtility.getNumericCellValue(orderBy, 27, 5), DELTA);
        assertEquals("Atlantic", TestUtility.getStringCellValue(orderBy, 31, 0));
        assertEquals(51.0 / (51 + 21), TestUtility.getNumericCellValue(orderBy, 31, 5), DELTA);

        assertEquals("Division: Southwest", TestUtility.getStringCellValue(orderBy, 0, 8));
        assertEquals("Rockets", TestUtility.getStringCellValue(orderBy, 2, 9));
        assertEquals(38.0 / (38 + 35), TestUtility.getNumericCellValue(orderBy, 2, 12), DELTA);
        assertEquals("Spurs", TestUtility.getStringCellValue(orderBy, 6, 9));
        assertEquals(57.0 / (57 + 16), TestUtility.getNumericCellValue(orderBy, 6, 12), DELTA);
        assertEquals("Division: Southeast", TestUtility.getStringCellValue(orderBy, 7, 8));
        assertEquals("Wizards", TestUtility.getStringCellValue(orderBy, 9, 9));
        assertEquals(17.0 / (17 + 55), TestUtility.getNumericCellValue(orderBy, 9, 12), DELTA);
        assertEquals("Heat", TestUtility.getStringCellValue(orderBy, 13, 9));
        assertEquals(51.0 / (51 + 22), TestUtility.getNumericCellValue(orderBy, 13, 12), DELTA);
        assertEquals("Division: Pacific", TestUtility.getStringCellValue(orderBy, 14, 8));
        assertEquals("Kings", TestUtility.getStringCellValue(orderBy, 16, 9));
        assertEquals(20.0 / (20 + 52), TestUtility.getNumericCellValue(orderBy, 16, 12), DELTA);
        assertEquals("Lakers", TestUtility.getStringCellValue(orderBy, 20, 9));
        assertEquals(53.0 / (53 + 20), TestUtility.getNumericCellValue(orderBy, 20, 12), DELTA);
        assertEquals("Division: Of Their Own", TestUtility.getStringCellValue(orderBy, 21, 8));
        assertEquals("Globetrotters", TestUtility.getStringCellValue(orderBy, 23, 9));
        assertEquals(21227.0 / (21227 + 341), TestUtility.getNumericCellValue(orderBy, 23, 12), DELTA);
        assertEquals("Division: Northwest", TestUtility.getStringCellValue(orderBy, 24, 8));
        assertEquals("Timberwolves", TestUtility.getStringCellValue(orderBy, 26, 9));
        assertEquals(17.0 / (17 + 57), TestUtility.getNumericCellValue(orderBy, 26, 12), DELTA);
        assertEquals("Thunder", TestUtility.getStringCellValue(orderBy, 30, 9));
        assertEquals(48.0 / (48 + 24), TestUtility.getNumericCellValue(orderBy, 30, 12), DELTA);
        assertEquals("Division: Central", TestUtility.getStringCellValue(orderBy, 31, 8));
        assertEquals("Cavaliers", TestUtility.getStringCellValue(orderBy, 33, 9));
        assertEquals(14.0 / (14 + 58), TestUtility.getNumericCellValue(orderBy, 33, 12), DELTA);
        assertEquals("Bulls", TestUtility.getStringCellValue(orderBy, 37, 9));
        assertEquals(53.0 / (53 + 19), TestUtility.getNumericCellValue(orderBy, 37, 12), DELTA);
        assertEquals("Division: Atlantic", TestUtility.getStringCellValue(orderBy, 38, 8));
        assertEquals("Raptors", TestUtility.getStringCellValue(orderBy, 40, 9));
        assertEquals(20.0 / (20 + 53), TestUtility.getNumericCellValue(orderBy, 40, 12), DELTA);
        assertEquals("Celtics", TestUtility.getStringCellValue(orderBy, 44, 9));
        assertEquals(51.0 / (51 + 21), TestUtility.getNumericCellValue(orderBy, 44, 12), DELTA);

        Sheet varStatus = workbook.getSheetAt(11);
        Map<Integer, String> expInds = new HashMap<>();
        expInds.put(1, "0 of 8");  expInds.put(2, "0 of 5");  expInds.put(3, "1 of 5");  expInds.put(4, "2 of 5");  expInds.put(5, "3 of 5");  expInds.put(6, "4 of 5");
        expInds.put(8, "1 of 8");  expInds.put(9, "0 of 5");  expInds.put(10, "1 of 5"); expInds.put(11, "2 of 5"); expInds.put(12, "3 of 5"); expInds.put(13, "4 of 5");
        expInds.put(15, "2 of 8"); expInds.put(16, "0 of 5"); expInds.put(17, "1 of 5"); expInds.put(18, "2 of 5"); expInds.put(19, "3 of 5"); expInds.put(20, "4 of 5");
        expInds.put(22, "3 of 8"); expInds.put(23, "0 of 5"); expInds.put(24, "1 of 5"); expInds.put(25, "2 of 5"); expInds.put(26, "3 of 5"); expInds.put(27, "4 of 5");
        expInds.put(29, "4 of 8"); expInds.put(30, "0 of 5"); expInds.put(31, "1 of 5"); expInds.put(32, "2 of 5"); expInds.put(33, "3 of 5"); expInds.put(34, "4 of 5");
        expInds.put(36, "5 of 8"); expInds.put(37, "0 of 5"); expInds.put(38, "1 of 5"); expInds.put(39, "2 of 5"); expInds.put(40, "3 of 5"); expInds.put(41, "4 of 5");
        expInds.put(43, "6 of 8");
        expInds.put(45, "7 of 8"); expInds.put(46, "0 of 1");
        List<Integer> firstTrues = Arrays.asList(1, 2, 9, 16, 23, 30, 37, 46);
        List<Integer> lastTrues = Arrays.asList(6, 13, 20, 27, 34, 41, 45, 46);

        for (int r = 0; r < 47; r++)
        {
            if (expInds.containsKey(r))
            {
                assertEquals("Row " + r, expInds.get(r), TestUtility.getStringCellValue(varStatus, r, 5));
                assertEquals("Row " + r, firstTrues.contains(r), TestUtility.getBooleanCellValue(varStatus, r, 6));
                assertEquals("Row " + r, lastTrues.contains(r), TestUtility.getBooleanCellValue(varStatus, r, 7));
            }
        }
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
        beans.putAll(TestUtility.getDivisionData());
        beans.putAll(TestUtility.getTeamsData());
        return beans;
    }
}