package net.sf.jett.test;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.parser.TagParser;

/**
 * This JUnit Test class tests the <code>TagParser</code>.
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class TagParserTest
{
    private static InputStream theInputStream;
    private static Workbook theWorkbook;

    /**
     * Before running any of the tests, open a spreadsheet full of test cases.
     * @throws IOException If there is a problem opening the spreadsheet file.
     * @throws InvalidFormatException If there is a problem with the spreadsheet
     *    format.
     */
    @BeforeClass
    public static void setup() throws IOException, InvalidFormatException
    {
        theInputStream = new BufferedInputStream(new FileInputStream("templates/TagParserTests.xlsx"));
        theWorkbook = WorkbookFactory.create(theInputStream);
    }

    /**
     * Close the <code>InputStream</code> on the spreadsheet.
     * @throws IOException If there is a problem closing the file.
     */
    @AfterClass
    public static void afterTests() throws IOException
    {
        theInputStream.close();
    }


    /**
     * Ensure that when the tag text contains a colon first, a
     * <code>TagParseException</code> is thrown because there a namespace name
     * was expected.
     */
    @Test(expected = TagParseException.class)
    public void testNoNamespace()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when the tag text contains no value after a namespace and a
     * colon, a <code>TagParseException</code> is thrown because there is
     * no tag name.
     */
    @Test(expected = TagParseException.class)
    public void testNamespaceButNoTagName()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when a value occurs without an attribute, a
     * <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testValueWithoutAttribute()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(2);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when an equals sign is found before an attribute name is
     * found, a <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testNoAttributeName()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(3);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when a colon is found inside an attribute name, a
     * <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testColonInAttributeName()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(4);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when a begin angle bracket is found within tag text, a
     * <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testNestedBeginTag()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(5);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when a begin angle bracket and slash is found within tag
     * text, a <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testNestedEndTag()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(6);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when the end angle bracket is not found, a
     * <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testNoEndTag()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(7);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when no attribute value is found, a
     * <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testNoAttrValue()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(8);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Ensure that when the tag text ends while inside an attribute value, a
     * <code>TagParseException</code> is thrown.
     */
    @Test(expected = TagParseException.class)
    public void testEoiValue()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(9);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();
    }

    /**
     * Tests the parsing of a bodiless tag.
     */
    @Test
    public void testBodilessTag()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(10);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();

        assertEquals(60, parser.getAfterTagIdx());

        Map<String, RichTextString> attributes = parser.getAttributes();
        assertEquals(3, attributes.size());

        List<String> attrNames = Arrays.asList("attr1", "attr2", "attr3");
        List<String> attrValues = Arrays.asList("value1", "true", "${expression}");
        for (int i = 0; i < 3; i++)
        {
            assertTrue(attributes.containsKey(attrNames.get(i)));
            assertEquals(attrValues.get(i), attributes.get(attrNames.get(i)).getString());
        }

        assertEquals(cell, parser.getCell());
        assertEquals("jt", parser.getNamespace());
        assertEquals("test", parser.getTagName());
        assertEquals("jt:test", parser.getNamespaceAndTagName());
        assertEquals("<jt:test attr1=\"value1\" attr2=\"true\" attr3=\"${expression}\"/>", parser.getTagText());
        assertTrue(parser.isBodiless());
        assertFalse(parser.isEndTag());
        assertTrue(parser.isTag());
    }

    /**
     * Ensure that the parser recognizes text that is not a tag.
     */
    @Test
    public void testNotATag()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(11);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();

        assertFalse(parser.isTag());
    }

    /**
     * Ensure that the parser recognizes a tag with a body.
     */
    @Test
    public void testTagWithBody()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(12);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();

        assertTrue(parser.isTag());
        assertFalse(parser.isEndTag());
        assertFalse(parser.isBodiless());
    }

    /**
     * Ensure that the parser recognizes an end tag.
     */
    @Test
    public void testEndTag()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(13);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();

        assertTrue(parser.isTag());
        assertTrue(parser.isEndTag());
    }

    /**
     * Ensure that all JETT-defined escaped sequences are recognized in
     * attribute values.
     * <ul>
     * <li>\" for double-quote</li>
     * <li>\\ for backslash</li>
     * <li>\' for single-quote</li>
     * <li>\b for backspace</li>
     * <li>\f for form-feed</li>
     * <li>\n for newline</li>
     * <li>\r for carriage return</li>
     * <li>\t for tab</li>
     * </ul>
     * @since 0.5.2
     */
    @Test
    public void testEscapesInAttributeValues()
    {
        Sheet sheet = theWorkbook.getSheetAt(0);
        Row row = sheet.getRow(14);
        Cell cell = row.getCell(0);
        TagParser parser = new TagParser(cell);
        parser.parse();

        Map<String, RichTextString> attributes = parser.getAttributes();
        List<String> attrNames = Arrays.asList(
                "doublequote", "backslash", "singlequote", "backspace", "formfeed", "newline", "carriagereturn", "tab");
        // Note: the double-quote and backslash must still be Java-escaped here, just to make it into the string.
        List<String> attrValues = Arrays.asList("Embedded \"double-quotes\"", "Embedded \\backslash", "Embedded 'single-quotes'",
                "Embedded \bbackspace", "Embedded \fform-feed", "Embedded \nnewline", "Embedded \rcarriage-return", "Embedded \ttab");
        for (int i = 0; i < attrNames.size(); i++)
        {
            assertTrue(attributes.containsKey(attrNames.get(i)));
            assertEquals(attrValues.get(i), attributes.get(attrNames.get(i)).getString());
        }
    }
}
