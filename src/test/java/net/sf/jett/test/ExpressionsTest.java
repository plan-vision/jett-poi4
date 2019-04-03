package net.sf.jett.test;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.jdbc.ResultSetRow;
import net.sf.jett.test.model.Team;

/**
 * This JUnit Test class tests the evaluation of expressions and replacement
 * in spreadsheet cells.
 *
 * @author Randy Gettman
 */
public class ExpressionsTest extends TestCase
{
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
        return "ExprTest";
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
        assertEquals("Cell A3", TestUtility.getStringCellValue(sheet, 2, 0));
        assertEquals(3.14, TestUtility.getNumericCellValue(sheet, 2, 2), DELTA);
        assertEquals("Hello World!", TestUtility.getStringCellValue(sheet, 3, 1));
        assertEquals("JETT", TestUtility.getStringCellValue(sheet, 4, 1));
        assertEquals("JETT: Hello World!", TestUtility.getStringCellValue(sheet, 5, 1));
        assertEquals("Springfield Isotopes (38-4)", TestUtility.getStringCellValue(sheet, 6, 2));
        assertEquals("Springfield", TestUtility.getStringCellValue(sheet, 7, 2));
        assertEquals("Isotopes", TestUtility.getStringCellValue(sheet, 8, 2));
        assertEquals(38, TestUtility.getNumericCellValue(sheet, 9, 2), DELTA);
        assertEquals(4, TestUtility.getNumericCellValue(sheet, 10, 2), DELTA);
        assertEquals("[1, 3, 7, 10, 11, 23]", TestUtility.getStringCellValue(sheet, 11, 1));
        double numberListAvg = (double) (1 + 3 + 7 + 10 + 11 + 23) / 6;
        assertEquals(numberListAvg, TestUtility.getNumericCellValue(sheet, 12, 1), DELTA);
        //         10        20        30        40        50        60        70        80        90       100
        //01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
        assertEquals("I can use bold, italic, underline, strikeout, different fonts, superscript, and subscript within one cell!",
                TestUtility.getStringCellValue(sheet, 14, 0));
        RichTextString rts = TestUtility.getRichTextStringCellValue(sheet, 14, 0);
        assertNotNull(rts);
        assertEquals(106, rts.length());

        int formattingRunIndices[] = new int[]
                // Bold  , Italic, FontUnderline, Strikeout, Fonts , Superscript, Subscript, One
                {10, 14, 16, 22, 24, 33   , 35, 44   , 56, 61, 63, 74     , 80, 89   , 97, 100};

        // HSSF (.xls) does not count the initial run as a different formatting
        // run; it is the formatting of the actual Cell.
        int adjust;
        if (sheet instanceof HSSFSheet)
        {
            adjust = 0;
        }
        else
        {
            // XSSFSheet
            adjust = 1;
            assertEquals(0, rts.getIndexOfFormattingRun(0));
        }
        assertEquals(16 + adjust, rts.numFormattingRuns());
        for (int i = 0; i < formattingRunIndices.length; i++)
        {
            assertEquals(formattingRunIndices[i], rts.getIndexOfFormattingRun(i + adjust));
        }

        assertEquals("B17", TestUtility.getStringCellValue(sheet, 16, 1));
        assertEquals("B17:D18", TestUtility.getStringCellValue(sheet, 16, 3));
        assertEquals("JETT supports static method calling!", TestUtility.getStringCellValue(sheet, 17, 1));

        assertEquals(0, TestUtility.getNumericCellValue(sheet, 19, 1), Math.ulp(0));
        assertEquals(42, TestUtility.getNumericCellValue(sheet, 20, 1), Math.ulp(42));
        assertEquals(1.2345678901234567890E39, TestUtility.getNumericCellValue(sheet, 21, 1), Math.ulp(1.2345678901234567890E39));
        assertEquals("359538626972463140000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000",
                TestUtility.getStringCellValue(sheet, 22, 1));

        assertEquals(0, TestUtility.getNumericCellValue(sheet, 23, 1), Math.ulp(0));
        assertEquals(8.6, TestUtility.getNumericCellValue(sheet, 24, 1), Math.ulp(8.6));
        assertEquals(0, TestUtility.getNumericCellValue(sheet, 25, 1), Math.ulp(0));
        assertEquals("359538626972463141629054847463408713596141135051689993197834953606314521560057077521179117265533756343080917907028764928468642653778928365536935093407075033972099821153102564152490980180778657888151737016910267884609166473806445896331617118664246696549595652408289446337476354361838599762500808052368249716736",
                TestUtility.getStringCellValue(sheet, 26, 1));

        assertEquals(42, TestUtility.getNumericCellValue(sheet, 27, 1), Math.ulp(0));
        assertEquals(8.6, TestUtility.getNumericCellValue(sheet, 28, 1), Math.ulp(0));

        assertEquals("${testBean1}Hello World!${testBean2}JETT", TestUtility.getStringCellValue(sheet, 29, 1));

        assertEquals(42, TestUtility.getNumericCellValue(sheet, 30, 1), Math.ulp(0));
        String card = TestUtility.getStringCellValue(sheet, 31, 1);
        assertNotNull(card);
        String[] fields = card.split("\\s+");
        assertEquals(3, fields.length);
        List<String> possibleRanks = Arrays.asList("Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
                "Jack", "Queen", "King", "Ace");
        List<String> possibleSuits = Arrays.asList("Clubs", "Diamonds", "Spades", "Hearts");
        assertTrue(possibleRanks.contains(fields[0]));
        assertEquals("of", fields[1]);
        assertTrue(possibleSuits.contains(fields[2]));

        Header header = sheet.getHeader();
        assertEquals("Header Left: 1", header.getLeft());
        assertEquals("Header Center: 3", header.getCenter());
        assertEquals("Header Right: 7", header.getRight());
        Footer footer = sheet.getFooter();
        assertEquals("Footer Left: 10", footer.getLeft());
        assertEquals("Footer Center: 11", footer.getCenter());
        assertEquals("Footer Right: 23", footer.getRight());
        assertEquals("ExprTest", sheet.getSheetName());
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
        beans.put("testBean1", "Hello World!");
        beans.put("testBean2", "JETT");
        Team team = new Team();
        team.setName("Isotopes");
        team.setCity("Springfield");
        team.setWins(38);
        team.setLosses(4);
        beans.put("team", team);
        List<Integer> numberList = new ArrayList<>();
        numberList.add(1);
        numberList.add(3);
        numberList.add(7);
        numberList.add(10);
        numberList.add(11);
        numberList.add(23);
        beans.put("numberList", numberList);
        beans.put("feat1", "bold");
        beans.put("feat2", "italic");
        beans.put("feat3", "underline");
        beans.put("feat4", "strikeout");
        beans.put("feat5", "fonts");
        beans.put("feat6", "superscript");
        beans.put("feat7", "subscript");
        beans.put("feat8", "one");

        beans.put("biZero", BigInteger.ZERO);
        beans.put("biAnswer", new BigInteger("42"));
        // 40 digits is longer than long.
        beans.put("biBiggerThanLong", new BigInteger("1234567890123456789012345678901234567890"));
        // This should be approximately twice Double.MAX_VALUE.
        beans.put("biBiggerThanDouble", new BigInteger(
                //        10        20        30        40        50        60        70        80        90       100        10        20        30        40        50        60        70        80        90       200        10        20        30        40        50        60        70        80        90       300        10
                //1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
                "359538626972463140000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
        ));

        beans.put("bdZero", BigDecimal.ZERO);
        beans.put("bdAnswer", new BigDecimal(8.6));
        // I would have put "Double.MIN_NORMAL / 2" here, but that constant was
        // created for JDK 1.6.
        beans.put("bdSmallerThanNormal", Double.longBitsToDouble(0x0010000000000000L) / 2);
        beans.put("bdBiggerThanDouble", new BigDecimal(Double.MAX_VALUE).multiply(new BigDecimal(2)));

        ResultSetRow row = new ResultSetRow();
        row.set("answer", 42);
        row.set("IHaveAQuestion", 8.6);

        beans.put("valueHolder", row);

        beans.put("newSheetName", "ExprTest");

        return beans;
    }
}