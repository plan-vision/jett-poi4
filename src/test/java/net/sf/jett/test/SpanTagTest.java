package net.sf.jett.test;

import java.io.IOException;
import java.util.Arrays;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.util.RichTextStringUtil;

/**
 * This JUnit Test class tests the evaluation of the "span" tag (always
 * bodiless).
 *
 * @author Randy Gettman
 */
public class SpanTagTest extends TestCase
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
        return "SpanTag";
    }

    /**
     * Validate the newly created resultant <code>Workbook</code> with JUnit
     * assertions.
     * @param workbook A <code>Workbook</code>.
     */
    @Override
    protected void check(Workbook workbook)
    {
        Sheet vert = workbook.getSheetAt(0);
        assertEquals("Case vert cell factor=3", TestUtility.getStringCellValue(vert, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(0, 2, 0, 0)));
        assertEquals("Case vert row factor=3", TestUtility.getStringCellValue(vert, 0, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(0, 2, 1, 6)));
        assertEquals("Case vert col factor=3", TestUtility.getStringCellValue(vert, 3, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(3, 20, 0, 0)));
        assertEquals("Case vert block factor=3", TestUtility.getStringCellValue(vert, 3, 1));
        RichTextString rts = TestUtility.getRichTextStringCellValue(vert, 3, 1);
        Font font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 0), workbook);
        assertTrue((font == null) || "000000".equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 22), workbook);
        assertTrue((font == null) || "000000".equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rts, 23), workbook);
        assertEquals("008000", TestUtility.getFontColorString(workbook, font));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(3, 20, 1, 6)));
        assertEquals("After1", TestUtility.getStringCellValue(vert, 21, 0));
        assertEquals("After2", TestUtility.getStringCellValue(vert, 21, 6));

        assertEquals("Case vert cell factor=1", TestUtility.getStringCellValue(vert, 22, 0));
        assertFalse(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(22, 22, 0, 0)));
        assertEquals("Case vert row factor=1", TestUtility.getStringCellValue(vert, 22, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(22, 22, 1, 6)));
        assertEquals("Case vert col factor=1", TestUtility.getStringCellValue(vert, 23, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(23, 28, 0, 0)));
        assertEquals("Case vert block factor=1", TestUtility.getStringCellValue(vert, 23, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vert, new CellRangeAddress(23, 28, 1, 6)));
        assertEquals("After3", TestUtility.getStringCellValue(vert, 29, 0));
        assertEquals("After4", TestUtility.getStringCellValue(vert, 29, 6));

        assertEquals("After5", TestUtility.getStringCellValue(vert, 30, 0));
        assertEquals("After6", TestUtility.getStringCellValue(vert, 30, 6));

        assertEquals(7, vert.getNumMergedRegions());

        Sheet horiz = workbook.getSheetAt(1);
        assertEquals("Case horiz cell factor=3", TestUtility.getStringCellValue(horiz, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(0, 0, 0, 2)));
        assertEquals("Case horiz row factor=3", TestUtility.getStringCellValue(horiz, 0, 3));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(0, 0, 3, 20)));
        assertEquals("Case horiz col factor=3", TestUtility.getStringCellValue(horiz, 1, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(1, 6, 0, 2)));
        assertEquals("Case horiz block factor=3", TestUtility.getStringCellValue(horiz, 1, 3));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(1, 6, 3, 20)));
        assertEquals("After1", TestUtility.getStringCellValue(horiz, 0, 21));
        assertEquals("After2", TestUtility.getStringCellValue(horiz, 6, 21));

        assertEquals("Case horiz cell factor=1", TestUtility.getStringCellValue(horiz, 0, 22));
        assertFalse(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(0, 0, 22, 22)));
        assertEquals("Case horiz row factor=1", TestUtility.getStringCellValue(horiz, 0, 23));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(0, 0, 23, 28)));
        assertEquals("Case horiz col factor=1", TestUtility.getStringCellValue(horiz, 1, 22));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(1, 6, 22, 22)));
        assertEquals("Case horiz block factor=1", TestUtility.getStringCellValue(horiz, 1, 23));
        assertTrue(TestUtility.isMergedRegionPresent(horiz, new CellRangeAddress(1, 6, 23, 28)));
        assertEquals("After3", TestUtility.getStringCellValue(horiz, 0, 29));
        assertEquals("After4", TestUtility.getStringCellValue(horiz, 6, 29));

        assertEquals("After5", TestUtility.getStringCellValue(horiz, 0, 30));
        assertEquals("After6", TestUtility.getStringCellValue(horiz, 6, 30));

        assertEquals(7, horiz.getNumMergedRegions());

        Sheet vertAdjust = workbook.getSheetAt(2);
        assertEquals("Case vert cell adjust=1", TestUtility.getStringCellValue(vertAdjust, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(0, 1, 0, 0)));
        assertEquals("Case vert row adjust=1", TestUtility.getStringCellValue(vertAdjust, 0, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(0, 1, 1, 6)));
        assertEquals("Case vert col adjust=1", TestUtility.getStringCellValue(vertAdjust, 2, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(2, 8, 0, 0)));
        assertEquals("Case vert block adjust=1", TestUtility.getStringCellValue(vertAdjust, 2, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(2, 8, 1, 6)));
        assertEquals("After1", TestUtility.getStringCellValue(vertAdjust, 9, 0));
        assertEquals("After2", TestUtility.getStringCellValue(vertAdjust, 9, 6));

        assertEquals("Case vert cell adjust=0", TestUtility.getStringCellValue(vertAdjust, 10, 0));
        assertFalse(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(10, 10, 0, 0)));
        assertEquals("Case vert row adjust=0", TestUtility.getStringCellValue(vertAdjust, 10, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(10, 10, 1, 6)));
        assertEquals("Case vert col adjust=0", TestUtility.getStringCellValue(vertAdjust, 11, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(11, 16, 0, 0)));
        assertEquals("Case vert block adjust=0", TestUtility.getStringCellValue(vertAdjust, 11, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(11, 16, 1, 6)));
        assertEquals("After3", TestUtility.getStringCellValue(vertAdjust, 17, 0));
        assertEquals("After4", TestUtility.getStringCellValue(vertAdjust, 17, 6));

        assertEquals("Case vert col adjust=-1", TestUtility.getStringCellValue(vertAdjust, 18, 0));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(18, 22, 0, 0)));
        assertEquals("Case vert block adjust=-1", TestUtility.getStringCellValue(vertAdjust, 18, 1));
        assertTrue(TestUtility.isMergedRegionPresent(vertAdjust, new CellRangeAddress(18, 22, 1, 6)));
        assertEquals("After5", TestUtility.getStringCellValue(vertAdjust, 23, 0));
        assertEquals("After6", TestUtility.getStringCellValue(vertAdjust, 23, 6));

        assertEquals(9, vertAdjust.getNumMergedRegions());

        Sheet horizAdjust = workbook.getSheetAt(3);
        assertEquals("Case horiz cell adjust=1", TestUtility.getStringCellValue(horizAdjust, 0, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(0, 0, 0, 1)));
        assertEquals("Case horiz row adjust=1", TestUtility.getStringCellValue(horizAdjust, 0, 2));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(0, 0, 2, 8)));
        assertEquals("Case horiz col adjust=1", TestUtility.getStringCellValue(horizAdjust, 1, 0));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(1, 6, 0, 1)));
        assertEquals("Case horiz block adjust=1", TestUtility.getStringCellValue(horizAdjust, 1, 2));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(1, 6, 2, 8)));
        assertEquals("After1", TestUtility.getStringCellValue(horizAdjust, 0, 9));
        assertEquals("After2", TestUtility.getStringCellValue(horizAdjust, 6, 9));

        assertEquals("Case horiz cell adjust=0", TestUtility.getStringCellValue(horizAdjust, 0, 10));
        assertFalse(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(0, 0, 10, 10)));
        assertEquals("Case horiz row adjust=0", TestUtility.getStringCellValue(horizAdjust, 0, 11));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(0, 0, 11, 16)));
        assertEquals("Case horiz col adjust=0", TestUtility.getStringCellValue(horizAdjust, 1, 10));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(1, 6, 10, 10)));
        assertEquals("Case horiz block adjust=0", TestUtility.getStringCellValue(horizAdjust, 1, 11));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(1, 6, 11, 16)));
        assertEquals("After3", TestUtility.getStringCellValue(horizAdjust, 0, 17));
        assertEquals("After4", TestUtility.getStringCellValue(horizAdjust, 6, 17));

        assertEquals("Case horiz row adjust=-1", TestUtility.getStringCellValue(horizAdjust, 0, 18));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(0, 0, 18, 22)));
        assertEquals("Case horiz block adjust=-1", TestUtility.getStringCellValue(horizAdjust, 1, 18));
        assertTrue(TestUtility.isMergedRegionPresent(horizAdjust, new CellRangeAddress(1, 6, 18, 22)));
        assertEquals("After5", TestUtility.getStringCellValue(horizAdjust, 0, 23));
        assertEquals("After6", TestUtility.getStringCellValue(horizAdjust, 6, 23));

        assertEquals(9, horizAdjust.getNumMergedRegions());

        Sheet vertBorder = workbook.getSheetAt(4);
        assertTrue(TestUtility.isMergedRegionPresent(vertBorder, new CellRangeAddress(1, 6, 1, 2)));
        for (int r = 1; r <= 6; r++)
        {
        	BorderStyle borderBottom = (r == 6) ? BorderStyle.THIN : BorderStyle.NONE;
        	BorderStyle borderTop = (r == 1) ? BorderStyle.THIN : BorderStyle.NONE;
            String borderBottomColor= "000000";
            String borderTopColor = "000000";
            for (int c = 1; c <= 2; c++)
            {
            	BorderStyle borderLeft = (c == 1) ? BorderStyle.THIN : BorderStyle.NONE;
            	BorderStyle borderRight = (c == 2) ? BorderStyle.THIN : BorderStyle.NONE;
                String borderLeftColor= "000000";
                String borderRightColor = "000000";

                CellStyle cs = TestUtility.getCellStyle(vertBorder, r, c);
                assertNotNull(cs);
                assertEquals(borderBottom, cs.getBorderBottom());
                assertEquals(borderTop, cs.getBorderTop());
                assertEquals(borderLeft, cs.getBorderLeft());
                assertEquals(borderRight, cs.getBorderRight());
                if (borderBottom != BorderStyle.NONE)
                    assertEquals(borderBottomColor, TestUtility.getCellBottomBorderColorString(vertBorder, r, c));
                if (borderTop != BorderStyle.NONE)
                    assertEquals(borderTopColor, TestUtility.getCellTopBorderColorString(vertBorder, r, c));
                if (borderLeft != BorderStyle.NONE)
                    assertEquals(borderLeftColor, TestUtility.getCellLeftBorderColorString(vertBorder, r, c));
                if (borderRight != BorderStyle.NONE)
                    assertEquals(borderRightColor, TestUtility.getCellRightBorderColorString(vertBorder, r, c));
            }
        }
        assertTrue(TestUtility.isMergedRegionPresent(vertBorder, new CellRangeAddress(1, 5, 4, 5)));
        for (int r = 1; r <= 5; r++)
        {
        	BorderStyle borderBottom = (r == 5) ? BorderStyle.THIN : BorderStyle.NONE;
        	BorderStyle borderTop = (r == 1) ? BorderStyle.THIN : BorderStyle.NONE;
            String borderBottomColor = "ff0000";
            String borderTopColor = "ff0000";
            for (int c = 4; c <= 5; c++)
            {
            	BorderStyle borderLeft = (c == 4) ? BorderStyle.THIN : BorderStyle.NONE;
            	BorderStyle borderRight = (c == 5) ? BorderStyle.THIN : BorderStyle.NONE;
                String borderLeftColor = "ff0000";
                String borderRightColor = "ff0000";

                CellStyle cs = TestUtility.getCellStyle(vertBorder, r, c);
                assertNotNull(cs);
                assertEquals(borderBottom, cs.getBorderBottom());
                assertEquals(borderTop, cs.getBorderTop());
                assertEquals(borderLeft, cs.getBorderLeft());
                assertEquals(borderRight, cs.getBorderRight());
                if (borderBottom != BorderStyle.NONE)
                    assertEquals(borderBottomColor, TestUtility.getCellBottomBorderColorString(vertBorder, r, c));
                if (borderTop != BorderStyle.NONE)
                    assertEquals(borderTopColor, TestUtility.getCellTopBorderColorString(vertBorder, r, c));
                if (borderLeft != BorderStyle.NONE)
                    assertEquals(borderLeftColor, TestUtility.getCellLeftBorderColorString(vertBorder, r, c));
                if (borderRight != BorderStyle.NONE)
                    assertEquals(borderRightColor, TestUtility.getCellRightBorderColorString(vertBorder, r, c));
            }
        }
        // Test that the new style didn't create a "gray50percent" style!
        assertTrue(TestUtility.isMergedRegionPresent(vertBorder, new CellRangeAddress(1, 5, 4, 5)));
        for (int r = 1; r <= 3; r++)
        {
        	BorderStyle borderBottom = (r == 3) ? BorderStyle.THICK : BorderStyle.NONE;
        	BorderStyle borderTop = (r == 1) ? BorderStyle.THICK : BorderStyle.NONE;
            String borderBottomColor = "000000";
            String borderTopColor = "000000";
            for (int c = 7; c <= 9; c++)
            {
            	BorderStyle borderLeft = (c == 7) ? BorderStyle.THICK : BorderStyle.NONE;
            	BorderStyle borderRight = (c == 9) ? BorderStyle.THICK : BorderStyle.NONE;
                String borderLeftColor = "000000";
                String borderRightColor = "000000";

                CellStyle cs = TestUtility.getCellStyle(vertBorder, r, c);
                assertNotNull(cs);
                assertEquals(borderBottom, cs.getBorderBottom());
                assertEquals(borderTop, cs.getBorderTop());
                assertEquals(borderLeft, cs.getBorderLeft());
                assertEquals(borderRight, cs.getBorderRight());
                assertEquals(FillPatternType.NO_FILL, cs.getFillPattern());
                if (borderBottom != BorderStyle.NONE)
                    assertEquals(borderBottomColor, TestUtility.getCellBottomBorderColorString(vertBorder, r, c));
                if (borderTop != BorderStyle.NONE)
                    assertEquals(borderTopColor, TestUtility.getCellTopBorderColorString(vertBorder, r, c));
                if (borderLeft != BorderStyle.NONE)
                    assertEquals(borderLeftColor, TestUtility.getCellLeftBorderColorString(vertBorder, r, c));
                if (borderRight != BorderStyle.NONE)
                    assertEquals(borderRightColor, TestUtility.getCellRightBorderColorString(vertBorder, r, c));
            }
        }
        // End test that the new style didn't create a "gray50percent" style!
        assertEquals(3, vertBorder.getNumMergedRegions());

        Sheet horizBorder = workbook.getSheetAt(5);
        assertTrue(TestUtility.isMergedRegionPresent(horizBorder, new CellRangeAddress(1, 2, 1, 6)));
        for (int r = 1; r <= 2; r++)
        {
        	BorderStyle borderBottom = (r == 2) ? BorderStyle.MEDIUM : BorderStyle.NONE;
        	BorderStyle borderTop = (r == 1) ? BorderStyle.MEDIUM : BorderStyle.NONE;
            String borderBottomColor = "000000";
            String borderTopColor = "000000";
            for (int c = 1; c <= 6; c++)
            {
            	BorderStyle borderLeft = (c == 1) ? BorderStyle.MEDIUM : BorderStyle.NONE;
            	BorderStyle borderRight = (c == 6) ? BorderStyle.MEDIUM : BorderStyle.NONE;
                String borderLeftColor = "000000";
                String borderRightColor = "000000";

                CellStyle cs = TestUtility.getCellStyle(horizBorder, r, c);
                assertNotNull(cs);
                assertEquals(borderBottom, cs.getBorderBottom());
                assertEquals(borderTop, cs.getBorderTop());
                assertEquals(borderLeft, cs.getBorderLeft());
                assertEquals(borderRight, cs.getBorderRight());
                if (borderBottom != BorderStyle.NONE)
                    assertEquals(borderBottomColor, TestUtility.getCellBottomBorderColorString(horizBorder, r, c));
                if (borderTop != BorderStyle.NONE)
                    assertEquals(borderTopColor, TestUtility.getCellTopBorderColorString(horizBorder, r, c));
                if (borderLeft != BorderStyle.NONE)
                    assertEquals(borderLeftColor, TestUtility.getCellLeftBorderColorString(horizBorder, r, c));
                if (borderRight != BorderStyle.NONE)
                    assertEquals(borderRightColor, TestUtility.getCellRightBorderColorString(horizBorder, r, c));
            }
        }
        assertTrue(TestUtility.isMergedRegionPresent(horizBorder, new CellRangeAddress(4, 5, 1, 5)));
        for (int r = 4; r <= 5; r++)
        {
        	BorderStyle borderBottom = (r == 5) ? BorderStyle.MEDIUM : BorderStyle.NONE;
        	BorderStyle borderTop = (r == 4) ? BorderStyle.MEDIUM : BorderStyle.NONE;
            String borderBottomColor = "0000ff";
            String borderTopColor = "0000ff";
            for (int c = 1; c <= 5; c++)
            {
            	BorderStyle borderLeft = (c == 1) ? BorderStyle.MEDIUM : BorderStyle.NONE;
            	BorderStyle borderRight = (c == 5) ? BorderStyle.MEDIUM : BorderStyle.NONE;
                String borderLeftColor = "0000ff";
                String borderRightColor = "0000ff";

                CellStyle cs = TestUtility.getCellStyle(horizBorder, r, c);
                assertNotNull(cs);
                assertEquals(borderBottom, cs.getBorderBottom());
                assertEquals(borderTop, cs.getBorderTop());
                assertEquals(borderLeft, cs.getBorderLeft());
                assertEquals(borderRight, cs.getBorderRight());
                if (borderBottom != BorderStyle.NONE)
                    assertEquals(borderBottomColor, TestUtility.getCellBottomBorderColorString(horizBorder, r, c));
                if (borderTop != BorderStyle.NONE)
                    assertEquals(borderTopColor, TestUtility.getCellTopBorderColorString(horizBorder, r, c));
                if (borderLeft != BorderStyle.NONE)
                    assertEquals(borderLeftColor, TestUtility.getCellLeftBorderColorString(horizBorder, r, c));
                if (borderRight != BorderStyle.NONE)
                    assertEquals(borderRightColor, TestUtility.getCellRightBorderColorString(horizBorder, r, c));
            }
        }
        assertEquals(2, horizBorder.getNumMergedRegions());

        Sheet colorNormal = workbook.getSheetAt(6);
        assertEquals(2, colorNormal.getNumMergedRegions());
        assertTrue(TestUtility.isMergedRegionPresent(colorNormal, new CellRangeAddress(1, 2, 0, 0)));
        assertTrue(TestUtility.isMergedRegionPresent(colorNormal, new CellRangeAddress(3, 5, 0, 0)));
        for (int r = 0; r < 6; r++)
        {
            CellStyle cs = TestUtility.getCellStyle(colorNormal, r, 0);
            assertNotNull(cs);
            Font f = workbook.getFontAt(cs.getFontIndex());
            // XSSFWorkbook apparently won't store a Font if AUTOMATIC is chosen on the template.
            // But JETT translates this to black.
            assertTrue("Row " + r, "000000".equals(TestUtility.getFontColorString(workbook, f)) ||
                    "null".equals(TestUtility.getFontColorString(workbook, f)));
        }

        Sheet fixed = workbook.getSheetAt(7);
        assertEquals(1, fixed.getNumMergedRegions());
        assertTrue(TestUtility.isMergedRegionPresent(fixed, new CellRangeAddress(0, 2, 1, 1)));
        assertEquals("subTitle1", TestUtility.getStringCellValue(fixed, 0, 0));
        assertEquals("TITLE1", TestUtility.getStringCellValue(fixed, 0, 1));
        assertEquals("subTitle2", TestUtility.getStringCellValue(fixed, 1, 0));
        assertEquals("subTitle3", TestUtility.getStringCellValue(fixed, 2, 0));
        assertEquals("subTitle4", TestUtility.getStringCellValue(fixed, 3, 0));
        assertEquals("TITLE2", TestUtility.getStringCellValue(fixed, 3, 1));
        assertEquals("After", TestUtility.getStringCellValue(fixed, 4, 0));
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
        Map<String, Object> beans = TestUtility.getElementData();
        // Used in "factor".
        beans.put("expand", 3);
        beans.put("nothing", 1);
        beans.put("remove", 0);
        // Used in "adjust".
        beans.put("shrink", -1);
        beans.put("same", 0);
        beans.put("grow", 1);
        // Used in "color normal".
        beans.put("heights", Arrays.asList(1, 2, 3));
        return beans;
    }
}
