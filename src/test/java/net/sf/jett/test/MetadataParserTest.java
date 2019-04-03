package net.sf.jett.test;

import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.exception.MetadataParseException;
import net.sf.jett.parser.MetadataParser;

/**
 * This JUnit Test class tests the <code>MetadataParser</code>.
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class MetadataParserTest
{
    /**
     * Tests a simple metadata string.
     */
    @Test
    public void testSimple()
    {
        String metadata = "extraRows=1;left=2;right=3;copyRight=true;fixed=true;pastEndAction=clear";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertTrue(parser.isDefiningCols());
        assertEquals("1", parser.getExtraRows());
        assertEquals("2", parser.getColsLeft());
        assertEquals("3", parser.getColsRight());
        assertEquals("clear", parser.getPastEndAction());
        assertEquals("true", parser.getCopyingRight());
        assertEquals("true", parser.getFixed());
    }

    /**
     * Tests that the "clear" past end action value is recognized and legal.
     * Also tests defaults for other keys.
     */
    @Test
    public void testPastEndActionClear()
    {
        String metadata = "pastEndAction=clear";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertFalse(parser.isDefiningCols());
        assertNull(parser.getExtraRows());
        assertNull(parser.getColsLeft());
        assertNull(parser.getColsRight());
        assertEquals("clear", parser.getPastEndAction());
        assertNull(parser.getCopyingRight());
        assertNull(parser.getFixed());
    }

    /**
     * Tests that the "remove" past end action value is recognized and legal.
     */
    @Test
    public void testPastEndActionRemove()
    {
        String metadata = "pastEndAction=remove";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("remove", parser.getPastEndAction());
    }

    /**
     * If there is an unrecognized key, ensure that a
     * <code>MetadataParseException</code> is thrown.
     */
    @Test(expected = MetadataParseException.class)
    public void testUnrecognizedKey()
    {
        String metadata = "badKey=true";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();
    }

    /**
     * If the key is missing, ensure that a
     * <code>MetadataParseException</code> is thrown.
     */
    @Test(expected = MetadataParseException.class)
    public void testKeyMissing()
    {
        String metadata = "=true";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();
    }

    /**
     * If the key value is missing, ensure that a
     * <code>MetadataParseException</code> is thrown.
     */
    @Test(expected = MetadataParseException.class)
    public void testKeyValueMissing()
    {
        String metadata = "extraRows=";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();
    }

    /**
     * If the equals sign is missing, ensure that a
     * <code>MetadataParseException</code> is thrown.
     */
    @Test(expected = MetadataParseException.class)
    public void testEqualsMissing()
    {
        String metadata = "extraRows";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();
    }

    /**
     * Ensure that the parser recognizes that it's defining columns if only the
     * left columns value is defined.
     */
    @Test
    public void testDefiningColsLeftOnly()
    {
        String metadata = "left=1";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertTrue(parser.isDefiningCols());
    }

    /**
     * Ensure that the parser recognizes that it's defining columns if only the
     * right columns value is defined.
     */
    @Test
    public void testDefiningColsRightOnly()
    {
        String metadata = "right=1";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertTrue(parser.isDefiningCols());
    }

    /**
     * Test the "none" group direction value.
     */
    @Test
    public void testGroupDirValueNone()
    {
        String metadata = "groupDir=none";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("none", parser.getGroupDir());
    }

    /**
     * Test the "cols" group direction value.
     */
    @Test
    public void testGroupDirValueCols()
    {
        String metadata = "groupDir=cols";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("cols", parser.getGroupDir());
    }

    /**
     * Test the "rows" group direction value.
     */
    @Test
    public void testGroupDirValueRows()
    {
        String metadata = "groupDir=rows";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("rows", parser.getGroupDir());
    }

    /**
     * Test the "collapse" value with "groupDir" "none".
     */
    @Test
    public void testCollapseGroupDirNone()
    {
        String metadata = "collapse=true;groupDir=none";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("none", parser.getGroupDir());
        assertEquals("true", parser.getCollapsingGroup());
    }

    /**
     * Test the "collapse" value with "groupDir" "rows".
     */
    @Test
    public void testCollapseGroupDirRows()
    {
        String metadata = "collapse=true;groupDir=rows";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("rows", parser.getGroupDir());
        assertEquals("true", parser.getCollapsingGroup());
    }

    /**
     * Test the "collapse" value with "groupDir" "cols".
     */
    @Test
    public void testCollapseGroupDirCols()
    {
        String metadata = "collapse=true;groupDir=cols";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("cols", parser.getGroupDir());
        assertEquals("true", parser.getCollapsingGroup());
    }

    /**
     * Test the "copyRight" value with "left".
     */
    @Test
    public void testCopyRightWithLeft()
    {
        String metadata = "copyRight=true;left=1";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("1", parser.getColsLeft());
        assertEquals("true", parser.getCopyingRight());
    }

    /**
     * Test the "copyRight" value with "right".
     */
    @Test
    public void testCopyRightWithRight()
    {
        String metadata = "copyRight=true;right=1";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("1", parser.getColsRight());
        assertEquals("true", parser.getCopyingRight());
    }

    /**
     * Tests the creation of a <code>TagLoopListener</code> with a class name.
     * @since 0.3.0
     */
    @Test
    public void testLoopListenerClass()
    {
        String metadata = "onLoopProcessed=net.sf.jett.test.model.BlockShadingLoopListener";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("net.sf.jett.test.model.BlockShadingLoopListener", parser.getTagLoopListener());
    }

    /**
     * Tests the retrieval of a <code>TagLoopListener</code> from the beans map.
     * @since 0.3.0
     */
    @Test
    public void testLoopListenerInstance()
    {
        String metadata = "onLoopProcessed=${blockShadingLoopListener}";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("${blockShadingLoopListener}", parser.getTagLoopListener());
    }

    /**
     * Tests the creation of a <code>TagListener</code> with a class name.
     * @since 0.3.0
     */
    @Test
    public void testListenerClass()
    {
        String metadata = "onProcessed=net.sf.jett.test.model.BoldTagListener";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("net.sf.jett.test.model.BoldTagListener", parser.getTagListener());
    }

    /**
     * Tests the retrieval of a <code>TagListener</code> from the beans map.
     * @since 0.3.0
     */
    @Test
    public void testListenerInstance()
    {
        String metadata = "onProcessed=${boldTagListener}";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("${boldTagListener}", parser.getTagListener());
    }

    /**
     * Tests the "looping" variable name.
     * @since 0.3.0
     */
    @Test
    public void testIndexVarName()
    {
        String metadata = "indexVar=index";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("index", parser.getIndexVarName());
    }

    /**
     * Tests the limit variable name.
     * @since 0.3.0
     */
    @Test
    public void testLimit()
    {
        String metadata = "limit=10";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("10", parser.getLimit());
    }

    /**
     * Tests the replacement value.
     * @since 0.7.0
     */
    @Test
    public void testReplacementValue()
    {
        String metadata = "pastEndAction=replaceExpr;replaceValue=\"-\"";

        MetadataParser parser = new MetadataParser(metadata);
        parser.parse();

        assertEquals("-", parser.getReplacementValue());
    }
}
