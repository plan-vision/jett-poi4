package net.sf.jett.test;

import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.model.Style;
import net.sf.jett.parser.StyleParser;
import net.sf.jett.exception.StyleParseException;

/**
 * This JUnit Test class tests the <code>StyleParser</code>.
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class StyleParserTest
{
    /**
     * Test to ensure that empty text is legal and defines no styles.
     */
    @Test
    public void testEmpty()
    {
        StyleParser parser = new StyleParser("");
        parser.parse();
        Map<String, Style> styleMap = parser.getStyleMap();
        assertEquals(0, styleMap.size());
    }

    /**
     * Ensure that a period "." needs to be at the start of a style definition.
     */
    @Test(expected = StyleParseException.class)
    public void testErrorPeriodNotFirst()
    {
        StyleParser parser = new StyleParser("style1 { font-weight: bold; font-italic: true}");
        parser.parse();
    }

    /**
     * Ensure that the style name is mandatory after the period "." character.
     */
    @Test(expected = StyleParseException.class)
    public void testErrorNoStyleName()
    {
        StyleParser parser = new StyleParser(". {font-weight: bold; font-italic: true");
        parser.parse();
    }

    /**
     * Ensure that the begin brace is mandatory after the style name.
     */
    @Test(expected = StyleParseException.class)
    public void testMissingBeginBrace()
    {
        StyleParser parser = new StyleParser(".style1 font-weight: bold; font-italic: true");
        parser.parse();
    }

    /**
     * Ensure that no properties is legal.
     */
    @Test
    public void testNoProperties()
    {
        StyleParser parser = new StyleParser(".style1 { }");
        parser.parse();
        Map<String, Style> styleMap = parser.getStyleMap();
        assertEquals(1, styleMap.size());
        Style style1 = styleMap.get("style1");
        assertNotNull(style1);
        assertFalse(style1.isStyleToApply());
    }

    /**
     * Ensure that there must be a property name.
     */
    @Test(expected = StyleParseException.class)
    public void testMissingPropertyName()
    {
        StyleParser parser = new StyleParser(".style1 { : bold }");
        parser.parse();
    }

    /**
     * Ensure that there must be a colon separating the property name and the
     * value.
     */
    @Test(expected = StyleParseException.class)
    public void testMissingColon()
    {
        StyleParser parser = new StyleParser(".style1 { font-weight bold }");
        parser.parse();
    }

    /**
     * Ensure that there must be a value following the colon.
     */
    @Test(expected = StyleParseException.class)
    public void testMissingValue()
    {
        StyleParser parser = new StyleParser(".style1 { font-weight : }");
        parser.parse();
    }

    /**
     * Test a legal style definition of multiple properties.
     */
    @Test
    public void testMultipleProperties()
    {
        StyleParser parser = new StyleParser(".style1 { font-weight : bold; font-italic: true }");
        parser.parse();
        Map<String, Style> styleMap = parser.getStyleMap();
        assertEquals(1, styleMap.size());
        Style style1 = styleMap.get("style1");
        assertNotNull(style1);
        assertTrue(style1.isStyleToApply());
        assertEquals(true, style1.getFontBoldweight());
        assertTrue(style1.isFontItalic());
    }

    /**
     * Test multiple styles with multiple properties.  One of them has an extra
     * semicolon at the end, which is legal.
     */
    @Test
    public void testMultipleStyles()
    {
        StyleParser parser = new StyleParser(".style1 { font-weight : bold; font-italic: true }\n" +
                ".style2 {alignment: center; border: thin}\n" +
                ".style3 {fill-background-color: blue; fill-foreground-color: #FFFF00; fill-pattern: horizontalstripe; }");
        parser.parse();
        Map<String, Style> styleMap = parser.getStyleMap();
        assertEquals(3, styleMap.size());

        Style style1 = styleMap.get("style1");
        assertNotNull(style1);
        assertTrue(style1.isStyleToApply());
        assertEquals(true, style1.getFontBoldweight());
        assertTrue(style1.isFontItalic());

        Style style2 = styleMap.get("style2");
        assertNotNull(style2);
        assertTrue(style2.isStyleToApply());
        assertEquals(HorizontalAlignment.CENTER, style2.getAlignment());
        assertEquals(BorderStyle.THIN, style2.getBorderBottomType());
        assertEquals(BorderStyle.THIN, style2.getBorderLeftType());
        assertEquals(BorderStyle.THIN, style2.getBorderRightType());
        assertEquals(BorderStyle.THIN, style2.getBorderTopType());

        Style style3 = styleMap.get("style3");
        assertNotNull(style3);
        assertTrue(style3.isStyleToApply());
        assertEquals("BLUE", style3.getFillBackgroundColor());
        assertEquals("#FFFF00", style3.getFillForegroundColor());
        //assertEquals(FillPatternType.HORIZONTALSTRIPE, style3.getFillPatternType());
    }

    /**
     * Test an in-line comment.
     */
    @Test
    public void testComment()
    {
        StyleParser parser = new StyleParser(". style1 { font-name: Arial; /*font-weight: bold; */ font-italic: false}");
        parser.parse();
        Map<String, Style> styleMap = parser.getStyleMap();
        assertEquals(1, styleMap.size());

        Style style1 = styleMap.get("style1");
        assertNotNull(style1);
        assertTrue(style1.isStyleToApply());
        assertEquals("ARIAL", style1.getFontName());
        assertNull(style1.getFontBoldweight());
        assertFalse(style1.isFontItalic());
    }

    /**
     * Test to ensure that an unended comment results in a parse exception.
     */
    @Test(expected = StyleParseException.class)
    public void testBadComment()
    {
        StyleParser parser = new StyleParser(". style1 { font-name: Arial; /*font-weight: bold; font-italic: false}");
        parser.parse();
    }

    /**
     * Test that multiple word values are parsed correctly.
     */
    @Test
    public void testMultipleWordValue()
    {
        StyleParser parser = new StyleParser(". style1 {font-name: Times New Roman }");
        parser.parse();
        Map<String, Style> styleMap = parser.getStyleMap();
        assertEquals(1, styleMap.size());

        Style style1 = styleMap.get("style1");
        assertNotNull(style1);
        assertEquals("TIMES NEW ROMAN", style1.getFontName());
    }
}
