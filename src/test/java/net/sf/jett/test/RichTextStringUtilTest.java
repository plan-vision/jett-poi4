package net.sf.jett.test;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;

import static org.junit.Assert.*;

import net.sf.jett.model.CellStyleCache;
import net.sf.jett.model.FontCache;
import net.sf.jett.util.RichTextStringUtil;

/**
 * This JUnit Test class tests the <code>RichTextStringUtil</code> class.
 *
 * @author Randy Gettman
 */
public class RichTextStringUtilTest
{
    /**
     * A repeated value referenced often.
     */
    //                                                     10
    //                                            012345678901
    public static final String TRUE_EXPRESSION = "${trueValue}";
    /**
     * A repeated value referenced often.
     */
    //                                                       10
    //                                             012345678901
    public static final String FALSE_EXPRESSION = "${falseValue}";
    /**
     * A repeated value referenced often.
     */
    //                                       01234567
    public static final String TRUE_VALUE = "I'm true";
    /**
     * A repeated value referenced often.
     */
    //                                                  10        20
    //                                         0123456789012345678901234
    public static final String FALSE_VALUE = "I'm long and I'm false!!!";
    /**
     * The color green, in hex.
     */
    public static final String GREEN_HEX_STRING = "008000";
    /**
     * The color red, in hex.
     */
    public static final String RED_HEX_STRING = "ff0000";

    private static InputStream theXlsInputStream;
    private static Workbook theXlsWorkbook;
    private static InputStream theXlsxInputStream;
    private static Workbook theXlsxWorkbook;

    private CellStyleCache myCellStyleCache;
    private FontCache myFontCache;

    /**
     * Before running any of the tests, open a spreadsheet full of test cases.
     *
     * @throws java.io.IOException                                        If there is a problem opening the spreadsheet file.
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException If there is a problem with the spreadsheet
     *                                                                    format.
     */
    @BeforeClass
    public static void setup() throws IOException, InvalidFormatException
    {
        theXlsInputStream = new BufferedInputStream(new FileInputStream("templates/RichTextStringUtilTests.xls"));
        theXlsWorkbook = WorkbookFactory.create(theXlsInputStream);
        theXlsxInputStream = new BufferedInputStream(new FileInputStream("templates/RichTextStringUtilTests.xlsx"));
        theXlsxWorkbook = WorkbookFactory.create(theXlsxInputStream);
    }

    /**
     * Close the <code>InputStream</code> on the spreadsheet.
     *
     * @throws IOException If there is a problem closing the file.
     */
    @AfterClass
    public static void afterTests() throws IOException
    {
        theXlsInputStream.close();
        theXlsxInputStream.close();
    }

    /**
     * Tests Excel 97-2003 RichTextString features.
     */
    @Test
    public void testXls()
    {
        myCellStyleCache = new CellStyleCache(theXlsWorkbook);
        myFontCache = new FontCache(theXlsWorkbook);
        genericTest(theXlsWorkbook);
    }

    /**
     * Tests Excel 2007+ RichTextString features.
     */
    @Test
    public void testXlsx()
    {
        myCellStyleCache = new CellStyleCache(theXlsxWorkbook);
        myFontCache = new FontCache(theXlsxWorkbook);
        genericTest(theXlsxWorkbook);
    }

    /**
     * Tests <code>RichTextStringUtil</code> features.
     *
     * @param workbook A <code>Workbook</code>.
     */
    public void genericTest(Workbook workbook)
    {
        CreationHelper helper = workbook.getCreationHelper();
        Sheet sheet = workbook.getSheetAt(0);

        RichTextString rts, rtsResult;
        String substring;
        Font font;
        Cell cell;
        rts = TestUtility.getRichTextStringCellValue(sheet, 0, 0);

        // Replace All 1
        rtsResult = RichTextStringUtil.replaceAll(rts, helper, TRUE_EXPRESSION, TRUE_VALUE);
        assertEquals("<jt:if test=\"${condition}\" then=\"" + TRUE_VALUE + "\" false=\"" + FALSE_EXPRESSION + "\"/>", rtsResult.getString());
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 32), workbook);
        assertTrue((font == null) || !GREEN_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 33), workbook);
        assertTrue(GREEN_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 40), workbook);
        assertTrue(GREEN_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 41), workbook);
        assertTrue((font == null) || !GREEN_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        // Substring 1
        rtsResult = RichTextStringUtil.substring(rts, helper, 33, 45);
        substring = rtsResult.getString();
        assertEquals(TRUE_EXPRESSION, substring);
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 0), workbook);
        assertEquals(GREEN_HEX_STRING, TestUtility.getFontColorString(workbook, font));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 11), workbook);
        assertEquals(GREEN_HEX_STRING, TestUtility.getFontColorString(workbook, font));
        // Apply Font 1
        cell = TestUtility.getCell(sheet, 1, 0);
        RichTextStringUtil.applyFont(rtsResult, cell, myCellStyleCache, myFontCache);
        assertNotNull(cell);
        font = workbook.getFontAt(cell.getCellStyle().getFontIndex());
        assertEquals(GREEN_HEX_STRING, TestUtility.getFontColorString(workbook, font));

        // Replace All 2
        rtsResult = RichTextStringUtil.replaceAll(rts, helper, FALSE_EXPRESSION, FALSE_VALUE);
        assertEquals("<jt:if test=\"${condition}\" then=\"" + TRUE_EXPRESSION + "\" false=\"" + FALSE_VALUE + "\"/>", rtsResult.getString());
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 53), workbook);
        assertTrue((font == null) || !RED_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 54), workbook);
        assertTrue(RED_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 78), workbook);
        assertTrue(RED_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 79), workbook);
        assertTrue((font == null) || !RED_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        // Substring 2
        rtsResult = RichTextStringUtil.substring(rts, helper, 54, 67);
        substring = rtsResult.getString();
        assertEquals(FALSE_EXPRESSION, substring);
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 0), workbook);
        assertTrue(RED_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        font = TestUtility.convertToFont(RichTextStringUtil.getFontAtIndex(rtsResult, 12), workbook);
        assertTrue(RED_HEX_STRING.equals(TestUtility.getFontColorString(workbook, font)));
        // Apply Font 2
        cell = TestUtility.getCell(sheet, 2, 0);
        RichTextStringUtil.applyFont(rtsResult, cell, myCellStyleCache, myFontCache);
        assertNotNull(cell);
        font = workbook.getFontAt(cell.getCellStyle().getFontIndex());
        assertEquals(RED_HEX_STRING, TestUtility.getFontColorString(workbook, font));
    }
}
