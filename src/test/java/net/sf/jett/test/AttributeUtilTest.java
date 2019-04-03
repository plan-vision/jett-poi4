package net.sf.jett.test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import static org.junit.Assert.*;

import org.junit.Before;
import org.junit.Test;

import net.sf.jett.exception.AttributeExpressionException;
import net.sf.jett.expression.ExpressionFactory;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.tag.BaseTag;
import net.sf.jett.tag.Tag;
import net.sf.jett.tag.TagContext;
import net.sf.jett.test.model.Division;
import net.sf.jett.test.model.Employee;
import net.sf.jett.util.AttributeUtil;

/**
 * Tests the <code>AttributeUtil</code> class.
 *
 * @author Randy Gettman
 * @since 0.6.0
 */
public class AttributeUtilTest
{
    private Map<String, Object> myBeans;
    private Tag myTag;

    /**
     * We don't have a <code>Workbook</code> here, but we can create our own
     * <code>RichTextString</code>.
     */
    private static class TestCreationHelper implements CreationHelper
    {
        @Override
        public ClientAnchor createClientAnchor()
        {
            return null;
        }

        @Override
        public DataFormat createDataFormat()
        {
            return null;
        }

        @Override
        public FormulaEvaluator createFormulaEvaluator()
        {
            return null;
        }

        @Override
        public Hyperlink createHyperlink(HyperlinkType type)
        {
            return null;
        }

        @Override
        public ExtendedColor createExtendedColor()
        {
            return null;
        }

        @Override
        public RichTextString createRichTextString(String text)
        {
            return new XSSFRichTextString(text);
        }

		@Override
		public AreaReference createAreaReference(String reference) {			
			return null;
		}

		@Override
		public AreaReference createAreaReference(CellReference topLeft, CellReference bottomRight) {
			return null;
		}
    }

    /**
     * Set up by creating the beans map.
     */
    @Before
    public void setup()
    {
        myBeans = new HashMap<String, Object>();
        myBeans.put("t", true);
        myBeans.put("f", false);
        myBeans.put("answer", 42);
        myBeans.put("zero", 0);
        myBeans.put("isquared", -1);
        myBeans.put("question", 8.6);
        myBeans.put("project", "JETT");
        myBeans.put("null", null);
        Employee bugs = new Employee();
        bugs.setFirstName("Bugs");
        bugs.setLastName("Bunny");
        bugs.setSalary(1500);
        myBeans.put("bugs", bugs);
        myBeans.put("acronym", Arrays.asList("Java", "Excel", "Template", "Translator"));
        myBeans.put("integerArray", new Integer[] {4, 8, 15, 16, 23, 42});
        myBeans.put("integerArrayArray", new Integer[][] {new Integer[] {4, 8}, new Integer[] {15, 16, 23}, new Integer[] {42}});

        myTag = new BaseTag()
        {
            public String getName()
            {
                return "AttributeUtilTestTag";
            }

            public boolean process()
            {
                return true;
            }

            /**
             * Returns a dummy <code>TagContext</code> with a dummy <code>Block</code>.
             * @return A dummy <code>TagContext</code>.
             */
            public TagContext getContext()
            {
                TagContext context = new TagContext();
                context.setBlock(new Block(null, 0, 0, 0, 0));
                return context;
            }

            /**
             * Returns a dummy <code>WorkbookContext</code> that contains only a
             * dummy <code>ExpressionFactory</code> and a dummy tag locations map.
             * @return A <code>WorkbookContext</code>.
             */
            public WorkbookContext getWorkbookContext()
            {
                WorkbookContext context = new WorkbookContext();
                context.setExpressionFactory(new ExpressionFactory());
                context.setTagLocationsMap(new HashMap<String, String>());
                return context;
            }
        };
    }

    /**
     * Make sure it's evaluated as <code>true</code>.
     */
    @Test
    public void testBooleanTrue()
    {
        assertTrue(AttributeUtil.evaluateBoolean(myTag, new XSSFRichTextString("${t}"), myBeans, false));
    }

    /**
     * Make sure it's evaluated as <code>false</code>.
     */
    @Test
    public void testBooleanFalse()
    {
        assertFalse(AttributeUtil.evaluateBoolean(myTag, new XSSFRichTextString("${f}"), myBeans, true));
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testBooleanDNE()
    {
        AttributeUtil.evaluateBoolean(myTag, new XSSFRichTextString("${dne}"), myBeans, true);
    }

    /**
     * Tests integer resolution.
     */
    @Test
    public void testEvaluateInt()
    {
        assertEquals(42, AttributeUtil.evaluateInt(myTag, new XSSFRichTextString("${answer}"), myBeans, "attr_name", 0));
    }

    /**
     * Proper exception must be thrown for unparseable integer.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateIntBad()
    {
        AttributeUtil.evaluateInt(myTag, new XSSFRichTextString("${t}"), myBeans, "attr_name", 0);
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateIntDNE()
    {
        AttributeUtil.evaluateInt(myTag, new XSSFRichTextString("${dne}"), myBeans, "attr_name", 0);
    }

    /**
     * Throw positive number at int method testing for being non-negative.
     */
    @Test
    public void testEvaluateNonNegativeIntPositive()
    {
        assertEquals(42, AttributeUtil.evaluateNonNegativeInt(myTag, new XSSFRichTextString("${answer}"), myBeans, "attr_name", -1));
    }

    /**
     * Throw zero at int method testing for being non-negative.
     */
    @Test
    public void testEvaluateNonNegativeIntZero()
    {
        assertEquals(0, AttributeUtil.evaluateNonNegativeInt(myTag, new XSSFRichTextString("${zero}"), myBeans, "attr_name", -1));
    }

    /**
     * Throw negative number at int method testing for being non-negative.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateNonNegativeNegative()
    {
        AttributeUtil.evaluateNonNegativeInt(myTag, new XSSFRichTextString("${isquared}"), myBeans, "attr_name", 0);
    }

    /**
     * Throw positive number at int method testing for being positive.
     */
    @Test
    public void testEvaluatePositiveIntPositive()
    {
        assertEquals(42, AttributeUtil.evaluatePositiveInt(myTag, new XSSFRichTextString("${answer}"), myBeans, "attr_name", -1));
    }

    /**
     * Throw zero at int method testing for being positive.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluatePositiveIntZero()
    {
        AttributeUtil.evaluatePositiveInt(myTag, new XSSFRichTextString("${zero}"), myBeans, "attr_name", -1);
    }

    /**
     * Throw negative number at int method testing for being positive.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluatePositiveIntNegative()
    {
        AttributeUtil.evaluatePositiveInt(myTag, new XSSFRichTextString("${isquared}"), myBeans, "attr_name", 0);
    }

    /**
     * Throw positive number at int method testing for being non-zero.
     */
    @Test
    public void testEvaluateNonZeroIntPositive()
    {
        assertEquals(42, AttributeUtil.evaluateNonZeroInt(myTag, new XSSFRichTextString("${answer}"), myBeans, "attr_name", -1));
    }

    /**
     * Throw zero at int method testing for being non-zero.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateNonZeroIntZero()
    {
        AttributeUtil.evaluateNonZeroInt(myTag, new XSSFRichTextString("${zero}"), myBeans, "attr_name", -1);
    }

    /**
     * Throw negative number at int method testing for being non-zero.
     */
    @Test
    public void testEvaluateNonZeroIntNegative()
    {
        assertEquals(-1, AttributeUtil.evaluateNonZeroInt(myTag, new XSSFRichTextString("${isquared}"), myBeans, "attr_name", 0));
    }

    /**
     * Tests double resolution.
     */
    @Test
    public void testEvaluateDouble()
    {
        assertEquals(8.6, AttributeUtil.evaluateDouble(myTag, new XSSFRichTextString("${question}"), myBeans, "attr_name", 0), 0.0000001);
    }

    /**
     * Proper exception must be thrown for unparseable double.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateDoubleBad()
    {
        AttributeUtil.evaluateDouble(myTag, new XSSFRichTextString("${t}"), myBeans, "attr_name", 0);
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateDoubleDNE()
    {
        AttributeUtil.evaluateDouble(myTag, new XSSFRichTextString("${dne}"), myBeans, "attr_name", 0);
    }

    /**
     * Make sure that a negative <code>double</code> yields an
     * <code>AttributeExpressionException</code>.
     *
     * @since 0.11.0
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluatePositiveDoubleNegative()
    {
        AttributeUtil.evaluatePositiveDouble(myTag, new XSSFRichTextString("${-question}"), myBeans, "attr_name", 1);
    }

    /**
     * Make sure that a zero <code>double</code> yields an
     * <code>AttributeExpressionException</code>.
     *
     * @since 0.11.0
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluatePositiveDoubleZero()
    {
        AttributeUtil.evaluatePositiveDouble(myTag, new XSSFRichTextString("${zero}"), myBeans, "attr_name", 1);
    }

    /**
     * Make sure that a negative <code>double</code> yields an
     * <code>AttributeExpressionException</code>.
     *
     * @since 0.11.0
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateNonNegativeDoubleNegative()
    {
        AttributeUtil.evaluateNonNegativeDouble(myTag, new XSSFRichTextString("${-question}"), myBeans, "attr_name", 0);
    }

    /**
     * Tests RichTextStrings.
     *
     * @since 0.9.0
     */
    @Test
    public void testEvaluateRichTextStringNotNull()
    {
        RichTextString result = (RichTextString) AttributeUtil.evaluateRichTextStringNotNull(myTag,
                new XSSFRichTextString("Name: ${bugs.lastName}, ${bugs.firstName}"),
                new TestCreationHelper(), myBeans, "attr_name", "");
        assertEquals("Name: Bunny, Bugs", result.toString());
    }

    /**
     * Ensures that if a <code>null</code> is passed, the exception is thrown.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateRichTextStringNull()
    {
        AttributeUtil.evaluateRichTextStringNotNull(myTag, new XSSFRichTextString("${null}"),
                new TestCreationHelper(), myBeans, "attr_name", "");
    }

    /**
     * Tests String resolution.
     */
    @Test
    public void testEvaluateString()
    {
        assertEquals("JETT", AttributeUtil.evaluateString(myTag, new XSSFRichTextString("${project}"), myBeans, null));
    }

    /**
     * Tests if a <code>null</code> comes out.
     */
    @Test
    public void testEvaluateStringNull()
    {
        assertNull(AttributeUtil.evaluateString(myTag, new XSSFRichTextString("${null}"), myBeans, "notNullDefault"));
    }

    // Can't have this test, because we have to have null be a valid possible result.
    //@Test(expected = AttributeExpressionException.class)
    //public void testEvaluateStringDNE()
    //{
    //   AttributeUtil.evaluateString(new XSSFRichTextString("${dne}"), myBeans, null);
    //}

    /**
     * Catches the null result.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateStringNotNull()
    {
        AttributeUtil.evaluateStringNotNull(myTag, new XSSFRichTextString("${null}"), myBeans, "attr_name", "notNullDefault");
    }

    /**
     * Tests that a result is contained in a set of specific values.
     */
    @Test
    public void testEvaluateStringSpecificValues()
    {
        assertEquals("JETT", AttributeUtil.evaluateStringSpecificValues(myTag, new XSSFRichTextString("${project}"), myBeans, "attr_name",
                Arrays.asList("Apache POI", "JETT", "jAgg"), null));
    }

    /**
     * Tests that an exception results when a result is not contained in a set
     * of specific values.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateStringSpecificValuesNotFound()
    {
        AttributeUtil.evaluateStringSpecificValues(myTag, new XSSFRichTextString("${project}"), myBeans, "attr_name",
                Arrays.asList("Apache POI", "jAgg"), null);
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateObjectDNE()
    {
        AttributeUtil.evaluateObject(myTag, "${dne}", myBeans, "attr_name", String.class, "notNullDefault");
    }

    /**
     * Test whether we can get an object of a specific class.
     */
    @Test
    public void testEvaluateObject()
    {
        Object obj = AttributeUtil.evaluateObject(myTag, "${bugs}", myBeans, "attr_name", Employee.class, null);
        assertNotNull(obj);
        assertTrue(obj instanceof Employee);
    }

    /**
     * Test whether we can detect the wrong class.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateObjectWrongClass()
    {
        AttributeUtil.evaluateObject(myTag, "${bugs}", myBeans, "attr_name", Division.class, null);
    }

    /**
     * Test whether we can instantiate the correct class.
     */
    @Test
    public void testEvaluateObjectInstantiate()
    {
        Object obj = AttributeUtil.evaluateObject(myTag, "net.sf.jett.test.model.Employee", myBeans, "attr_name", Employee.class, null);
        assertNotNull(obj);
        assertTrue(obj instanceof Employee);
    }

    /**
     * Test whether we can detect the wrongly instantiated class.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateObjectInstantiateWrongClass()
    {
        AttributeUtil.evaluateObject(myTag, "net.sf.jett.test.model.Employee", myBeans, "attr_name", Division.class, null);
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateListDNE()
    {
        AttributeUtil.evaluateList(myTag, new XSSFRichTextString("${dne}"), myBeans, null);
    }

    /**
     * Test whether we can get resolve a <code>List</code>.
     */
    @Test
    public void testEvaluateList()
    {
        Object obj = AttributeUtil.evaluateList(myTag, new XSSFRichTextString("${acronym}"), myBeans, null);
        assertNotNull(obj);
        assertTrue(obj instanceof List);
        List list = (List) obj;
        assertEquals(4, list.size());
    }

    /**
     * Test whether we can create a list from a semicolon-separated string.
     */
    @Test
    //@SuppressWarnings("unchecked")
    public void testEvaluateListSemicolonSeparated()
    {
        Object obj = AttributeUtil.evaluateList(myTag, new XSSFRichTextString("four;eight;fifteen;sixteen;twenty-three;forty-two"), myBeans, null);
        assertNotNull(obj);
        assertTrue(obj instanceof List);
        List list = (List) obj;
        assertEquals(6, list.size());
        List<String> expected = Arrays.asList("four", "eight", "fifteen", "sixteen", "twenty-three", "forty-two");
        for (int i = 0; i < expected.size(); i++)
        {
            assertEquals(expected.get(i), list.get(i).toString());
        }
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateIntegerArrayDNE()
    {
        AttributeUtil.evaluateIntegerArray(myTag, new XSSFRichTextString("${dne}"), myBeans, Arrays.asList(1));
    }

    /**
     * Resolve an <code>int[]</code>.
     */
    @Test
    public void testEvaluateIntegerArray()
    {
        List<Integer> intList = AttributeUtil.evaluateIntegerArray(myTag, new XSSFRichTextString("${[4, 8, 15, 16, 23, 42]}"), myBeans, null);
        assertNotNull(intList);
        assertEquals(6, intList.size());
        List<Integer> expected = Arrays.asList(4, 8, 15, 16, 23, 42);
        for (int i = 0; i < expected.size(); i++)
        {
            assertEquals(expected.get(i), intList.get(i));
        }
    }

    /**
     * Resolve an <code>Integer[]</code>.
     */
    @Test
    public void testEvaluateIntegerArrayIntegerArray()
    {
        List<Integer> intList = AttributeUtil.evaluateIntegerArray(myTag, new XSSFRichTextString("${integerArray}"), myBeans, null);
        assertNotNull(intList);
        assertEquals(6, intList.size());
        List<Integer> expected = Arrays.asList(4, 8, 15, 16, 23, 42);
        for (int i = 0; i < expected.size(); i++)
        {
            assertEquals(expected.get(i), intList.get(i));
        }
    }

    /**
     * Parse an integer array from a semicolon-delimited string.
     */
    @Test
    public void testEvaluateIntegerArrayParse()
    {
        List<Integer> intList = AttributeUtil.evaluateIntegerArray(myTag, new XSSFRichTextString("4;8;15;16;23;42"), myBeans, null);
        assertNotNull(intList);
        assertEquals(6, intList.size());
        List<Integer> expected = Arrays.asList(4, 8, 15, 16, 23, 42);
        for (int i = 0; i < expected.size(); i++)
        {
            assertEquals(expected.get(i), intList.get(i));
        }
    }

    /**
     * Make sure that a bad expression with an undefined variable yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateIntegerArrayArrayDNE()
    {
        List<List<Integer>> def = new ArrayList<>();
        def.add(Arrays.asList(1));
        AttributeUtil.evaluateIntegerArrayArray(myTag, new XSSFRichTextString("${dne}"), myBeans, def);
    }

    /**
     * Resolve an <code>int[][]</code>.
     */
    @Test
    public void testEvaluateIntegerArrayArray()
    {
        List<List<Integer>> intList = AttributeUtil.evaluateIntegerArrayArray(myTag, new XSSFRichTextString("${[[4, 8], [15, 16, 23], [42]]}"), myBeans, null);
        assertNotNull(intList);
        assertEquals(3, intList.size());

        List<List<Integer>> expected = new ArrayList<>();
        expected.add(Arrays.asList(4, 8));
        expected.add(Arrays.asList(15, 16, 23));
        expected.add(Arrays.asList(42));

        for (int i = 0; i < expected.size(); i++)
        {
            List<Integer> expectedInternalList = expected.get(i);
            List<Integer> internalList = intList.get(i);
            assertEquals(expectedInternalList.size(), internalList.size());
            for (int j = 0; j < expectedInternalList.size(); j++)
            {
                assertEquals(expectedInternalList.get(j), internalList.get(j));
            }
        }
    }

    /**
     * Resolve an <code>Integer[][]</code>.
     */
    @Test
    public void testEvaluateIntegerArrayArrayIntegerArrayArray()
    {
        List<List<Integer>> intList = AttributeUtil.evaluateIntegerArrayArray(myTag, new XSSFRichTextString("${integerArrayArray}"), myBeans, null);
        assertNotNull(intList);
        assertEquals(3, intList.size());

        List<List<Integer>> expected = new ArrayList<>();
        expected.add(Arrays.asList(4, 8));
        expected.add(Arrays.asList(15, 16, 23));
        expected.add(Arrays.asList(42));

        for (int i = 0; i < expected.size(); i++)
        {
            List<Integer> expectedInternalList = expected.get(i);
            List<Integer> internalList = intList.get(i);
            assertEquals(expectedInternalList.size(), internalList.size());
            for (int j = 0; j < expectedInternalList.size(); j++)
            {
                assertEquals(expectedInternalList.get(j), internalList.get(j));
            }
        }
    }

    /**
     * Parse an integer array from a semicolon-delimited string.
     */
    @Test
    public void testEvaluateIntegerArrayArrayParse()
    {
        List<List<Integer>> intList = AttributeUtil.evaluateIntegerArrayArray(myTag, new XSSFRichTextString("4,8;15,16,23;42"), myBeans, null);
        assertNotNull(intList);
        assertEquals(3, intList.size());

        List<List<Integer>> expected = new ArrayList<>();
        expected.add(Arrays.asList(4, 8));
        expected.add(Arrays.asList(15, 16, 23));
        expected.add(Arrays.asList(42));

        for (int i = 0; i < expected.size(); i++)
        {
            List<Integer> expectedInternalList = expected.get(i);
            List<Integer> internalList = intList.get(i);
            assertEquals(expectedInternalList.size(), internalList.size());
            for (int j = 0; j < expectedInternalList.size(); j++)
            {
                assertEquals(expectedInternalList.get(j), internalList.get(j));
            }
        }
    }

    /**
     * Parse a valid variable name.
     */
    @Test
    public void testEvaluateStringVarName()
    {
        String varName = AttributeUtil.evaluateStringVarName(myTag, new XSSFRichTextString("_varUpper123"), myBeans, null);
        assertNotNull(varName);
        assertEquals("_varUpper123", varName);
    }

    /**
     * Make sure that a variable name starting with a number yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateStringVarNameStartNumber()
    {
        AttributeUtil.evaluateStringVarName(myTag, new XSSFRichTextString("1a"), myBeans, null);
    }

    /**
     * Make sure that a variable name with punctuation yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateStringVarNamePunctInName()
    {
        AttributeUtil.evaluateStringVarName(myTag, new XSSFRichTextString("abc-def"), myBeans, null);
    }

    /**
     * Make sure that a variable name that is a JEXL reserved word yields an
     * <code>AttributeExpressionException</code>.
     */
    @Test(expected = AttributeExpressionException.class)
    public void testEvaluateStringVarNameReservedWord()
    {
        AttributeUtil.evaluateStringVarName(myTag, new XSSFRichTextString("var"), myBeans, null);
    }
}