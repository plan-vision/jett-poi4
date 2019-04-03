package net.sf.jett.util;

import java.util.Arrays;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.CreationHelper;

import net.sf.jett.exception.AttributeExpressionException;
import net.sf.jett.expression.Expression;
import net.sf.jett.tag.Tag;

/**
 * The <code>AttributeUtil</code> class provides methods for
 * evaluating <code>Expressions</code> that are expected to result in a
 * specific type.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class AttributeUtil
{
    /**
     * Separates expressions in attributes that take multiple values.  This was
     * originally defined as the same value in multiple sub-classes, but was
     * moved to BaseTag/AttributeUtil for 0.3.0.
     *
     * @since 0.3.0
     */
    public static final String SPEC_SEP = ";";
    /**
     * Separates expressions in attributes that take multiple values at a second
     * level.  I.e. this is possible: "0,1;2,3" which would be interpreted as a
     * 2D array: <code>[[0, 1], [2, 3]]</code>.
     *
     * @since 0.4.0
     */
    public static final String SPEC_SEP_2 = ",";
    /**
     * Regex to validate a JEXL
     * <a href="http://commons.apache.org/proper/commons-jexl/reference/syntax.html#Language_Elements">variable name</a>.
     *
     * @since 0.11.0
     */
    private static final String REGEX_JEXL_VARNAME = "[A-Za-z_][A-Za-z0-9_]*";
    /**
     * List of
     * <a href="http://commons.apache.org/proper/commons-jexl/reference/syntax.html#Language_Elements">JEXL reserved words</a>.
     *
     * @since 0.11.0
     */
    private static final List<String> JEXL_RESERVED_WORDS = Arrays.asList(
            "or", "and", "eq", "ne", "lt", "gt", "le", "ge", "div", "mod", "not", "null",
            "true", "false", "new", "var", "return"
    );

    /**
     * Don't allow instances.
     *
     * @since 0.8.0
     */
    private AttributeUtil()
    {
    }

    /**
     * Helper method to throw an <code>AttributeExpressionException</code> with
     * a common message indicating that a null value resulted, or an expected
     * variable was missing when attempting to evaluate an expression inside an
     * attribute value.
     *
     * @param tag        The <code>Tag</code>.
     * @param expression The original expression.
     * @return <code>AttributeExpressionException</code> with a standard message.
     */
    private static AttributeExpressionException nullValueOrExpectedVariableMissing(Tag tag, String expression)
    {
        return attributeValidationFailure(tag, expression,
                "Null value or expected variable missing in expression");
    }

    /**
     * Helper method to throw an <code>AttributeExpressionException</code> with
     * a custom validation message.
     *
     * @param tag        The <code>Tag</code>.
     * @param expression The original expression.
     * @param message    The custom message.
     * @return <code>AttributeExpressionException</code> with a custom message.
     * @since 0.9.0
     */
    private static AttributeExpressionException attributeValidationFailure(Tag tag,
                                                                           String expression, String message)
    {
        return new AttributeExpressionException(message + " \"" +
                expression + "\"." + SheetUtil.getTagLocationWithHierarchy(tag));
    }

    /**
     * Helper method to throw an <code>AttributeExpressionException</code> with
     * a custom validation message.
     *
     * @param tag        The <code>Tag</code>.
     * @param expression The original expression.
     * @param message    The custom message.
     * @param cause      The <code>Exception</code> that caused the validation failure.
     * @return <code>AttributeExpressionException</code> with a custom message.
     * @since 0.9.0
     */
    private static AttributeExpressionException attributeValidationFailure(Tag tag,
                                                                           String expression, String message, Exception cause)
    {
        return new AttributeExpressionException(message + " \"" +
                expression + "\"." + SheetUtil.getTagLocationWithHierarchy(tag), cause);
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a boolean value from
     * the result, calling <code>Boolean.parseBoolean()</code> on the result if
     * necessary.  If the text is null, then the result defaults to the given
     * default boolean value.
     *
     * @param tag   The <code>Tag</code>.
     * @param text  Text which may have embedded <code>Expressions</code>.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @param def   The default value if the text is null.
     * @return The boolean result.
     */
    public static boolean evaluateBoolean(Tag tag,
                                          RichTextString text, Map<String, Object> beans, boolean def)
    {
        boolean result;
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text.toString());
        if (obj instanceof Boolean)
            result = (Boolean) obj;
        else
            result = Boolean.parseBoolean(obj.toString());
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract an integer value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.  If the text is null, then the result defaults to the given
     * default integer value.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The integer result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number.
     */
    public static int evaluateInt(Tag tag,
                                  RichTextString text, Map<String, Object> beans, String attrName, int def)
    {
        int result;
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text.toString());
        if (obj instanceof Number)
        {
            result = ((Number) obj).intValue();
        }
        else
        {
            try
            {
                result = Integer.parseInt(obj.toString());
            }
            catch (NumberFormatException e)
            {
                throw attributeValidationFailure(tag, text.toString(),
                        "The \"" + attrName + "\" attribute must be an integer");
            }
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract an integer value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.  Enforce the result to be non-negative.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The integer result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number, or if the result is negative.
     */
    public static int evaluateNonNegativeInt(Tag tag,
                                             RichTextString text, Map<String, Object> beans, String attrName, int def)
    {
        int result = evaluateInt(tag, text, beans, attrName, def);
        if (result < 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "The \"" + attrName + "\" attribute must be non-negative");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract an integer value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.  Enforce the result to be positive.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The integer result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number, or if the result is not positive.
     */
    public static int evaluatePositiveInt(Tag tag,
                                          RichTextString text, Map<String, Object> beans, String attrName, int def)
    {
        int result = evaluateInt(tag, text, beans, attrName, def);
        if (result <= 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "The \"" + attrName + "\" attribute must be positive");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract an integer value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.  Enforce the result to be not zero.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The integer result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number, or if the result is zero.
     */
    public static int evaluateNonZeroInt(Tag tag,
                                         RichTextString text, Map<String, Object> beans, String attrName, int def)
    {
        int result = evaluateInt(tag, text, beans, attrName, def);
        if (result == 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "The \"" + attrName + "\" attribute must not be zero");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a double value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The double result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number.
     */
    public static double evaluateDouble(Tag tag,
                                        RichTextString text, Map<String, Object> beans, String attrName, double def)
    {
        double result;
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text.toString());
        if (obj instanceof Number)
        {
            result = ((Number) obj).doubleValue();
        }
        else
        {
            try
            {
                result = Double.parseDouble(obj.toString());
            }
            catch (NumberFormatException e)
            {
                throw attributeValidationFailure(tag, text.toString(),
                        "The \"" + attrName + "\" attribute must be a number");
            }
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a double value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.  Enforce the result to be non-negative.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The double result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number, or if the result is negative.
     * @since 0.11.0
     */
    public static double evaluateNonNegativeDouble(Tag tag,
                                                   RichTextString text, Map<String, Object> beans, String attrName, double def)
    {
        double result = evaluateDouble(tag, text, beans, attrName, def);
        if (result < 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "The \"" + attrName + "\" attribute must be non-negative");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a double value from
     * the result, calling <code>toString()</code> on the result and parsing it
     * if necessary.  Enforce the result to be positive.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The double result.
     * @throws AttributeExpressionException If the result of the evaluation of the text is
     *                                      not a number, or if the result is not positive.
     * @since 0.11.0
     */
    public static double evaluatePositiveDouble(Tag tag,
                                                RichTextString text, Map<String, Object> beans, String attrName, double def)
    {
        double result = evaluateDouble(tag, text, beans, attrName, def);
        if (result <= 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "The \"" + attrName + "\" attribute must be positive");
        }
        return result;
    }

    /**
     * Evaluates the given rich text, which may have embedded
     * <code>Expressions</code>, and attempts to extract the result, which may
     * be either a <code>RichTextString</code> or a <code>String</code>.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Rich text which may have embedded <code>Expressions</code>.
     * @param helper   A <code>CreationHelper</code> (for creating
     *                 <code>RichTextStrings</code>).
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The <code>String</code> or <code>RichTextString</code> result.
     * @since 0.9.0
     */
    public static Object evaluateRichTextStringNotNull(Tag tag,
                                                       RichTextString text, CreationHelper helper, Map<String, Object> beans, String attrName, String def)
    {
        if (text == null)
            return def;
        Object result = Expression.evaluateString(text, helper, tag.getWorkbookContext().getExpressionFactory(), beans);
        if (result == null || result.toString().length() == 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "Value for \"" + attrName + "\" must not be null or empty");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>String</code>
     * result, calling <code>toString()</code> on the result.
     *
     * @param tag   The <code>Tag</code>.
     * @param text  Text which may have embedded <code>Expressions</code>.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @param def   The default value if the text is null.
     * @return The <code>String</code> result.
     */
    public static String evaluateString(Tag tag,
                                        RichTextString text, Map<String, Object> beans, String def)
    {
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        return (obj == null) ? null : obj.toString();
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>String</code>
     * result, calling <code>toString()</code> on the result.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @param def      The default value if the text is null.
     * @return The <code>String</code> result.
     */
    public static String evaluateStringNotNull(Tag tag,
                                               RichTextString text, Map<String, Object> beans, String attrName, String def)
    {
        String result = evaluateString(tag, text, beans, def);
        if (result == null || result.length() == 0)
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "Value for \"" + attrName + "\" must not be null or empty");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>String</code>
     * result, calling <code>toString()</code> on the result.  Enforces that the
     * result is one of the given expected values, ignoring case.
     *
     * @param tag         The <code>Tag</code>.
     * @param text        Text which may have embedded <code>Expressions</code>.
     * @param beans       A <code>Map</code> of bean names to bean values.
     * @param attrName    The attribute name.  This is only used when constructing
     *                    an exception message.
     * @param legalValues A <code>List</code> of expected values.
     * @param def         The default value if the text is null.
     * @return The <code>String</code> result.
     * @throws AttributeExpressionException If the result isn't one of the expected legal
     *                                      values.
     */
    public static String evaluateStringSpecificValues(Tag tag,
                                                      RichTextString text, Map<String, Object> beans, String attrName, List<String> legalValues, String def)
    {
        String result = evaluateString(tag, text, beans, def);
        for (String legalValue : legalValues)
        {
            if (legalValue.equalsIgnoreCase(result))
                return result;
        }
        throw attributeValidationFailure(tag, text.toString(),
                "Unknown value for \"" + attrName + "\": " + result +
                        " (expected one of " + legalValues.toString() + ").");
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>String</code>
     * result, calling <code>toString()</code> on the result.  Enforces that the
     * result is a valid JEXL variable name, which contains only the following
     * characters: <code>[A-Z][a-z][0-9][_]</code>, and not starting with a
     * number.
     *
     * @param tag      The <code>Tag</code>.
     * @param text     Text which may have embedded <code>Expressions</code>.
     * @param beans    A <code>Map</code> of bean names to bean values.
     * @param attrName The attribute name.  This is only used when constructing
     *                 an exception message.
     * @return The <code>String</code> result.
     * @throws AttributeExpressionException If the result isn't a legal JEXL
     *                                      variable name.
     * @since 0.11.0
     */
    public static String evaluateStringVarName(Tag tag,
                                               RichTextString text, Map<String, Object> beans, String attrName)
    {
        String result = evaluateStringNotNull(tag, text, beans, attrName, null);

        if (!result.matches(REGEX_JEXL_VARNAME))
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "Not a valid JEXL variable name: " + result);
        }
        if (JEXL_RESERVED_WORDS.contains(result))
        {
            throw attributeValidationFailure(tag, text.toString(),
                    "Can't use a JEXL reserved word: " + result);
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a result, and cast it
     * to the same class as the given expected class.
     *
     * @param tag           The <code>Tag</code>.
     * @param text          Text which may have embedded <code>Expressions</code>.
     * @param beans         A <code>Map</code> of bean names to bean values.
     * @param attrName      The attribute name.  This is only used when constructing
     *                      an exception message.
     * @param expectedClass The result is expected to be of the given class or
     *                      of a subclass.
     * @param def           The default value if the text is null.
     * @param <T>           The <code>Class</code> of the expected return type.
     * @return The result.
     * @throws AttributeExpressionException If the result is not of the expected class or
     *                                      of a subclass.
     */
    @SuppressWarnings("unchecked")
    public static <T> T evaluateObject(Tag tag,
                                       RichTextString text, Map<String, Object> beans, String attrName, Class<T> expectedClass, T def)
    {
        if (text == null)
            return def;

        return evaluateObject(tag, text.toString(), beans, attrName, expectedClass, def);
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a result, and cast it
     * to the same class as the given expected class.
     *
     * @param tag           The <code>Tag</code>.
     * @param text          Text which may have embedded <code>Expressions</code>.
     * @param beans         A <code>Map</code> of bean names to bean values.
     * @param attrName      The attribute name.  This is only used when constructing
     *                      an exception message.
     * @param expectedClass The result is expected to be of the given class or
     *                      of a subclass.
     * @param def           The default value if the text is null.
     * @param <T>           The <code>Class</code> of the expected return type.
     * @return The result.
     * @throws AttributeExpressionException If the result is not of the expected class or
     *                                      of a subclass.
     */
    @SuppressWarnings("unchecked")
    public static <T> T evaluateObject(Tag tag,
                                       String text, Map<String, Object> beans, String attrName, Class<T> expectedClass, T def)
    {
        T result;
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text, tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text);
        Class objClass = obj.getClass();
        if (expectedClass.isAssignableFrom(objClass))
        {
            // Don't expect a ClassCastException after the above test.
            result = expectedClass.cast(obj);
        }
        else if (obj instanceof String)
        {
            String className = (String) obj;
            // Treat as a class name to instantiate.
            try
            {
                Class<T> actualClass = (Class<T>) Class.forName(className);
                result = actualClass.newInstance();
                if (!expectedClass.isInstance(result))
                {
                    throw attributeValidationFailure(tag, text, "Expected a \"" + expectedClass.getName() + "\" for \"" +
                            attrName + "\", but instantiated a \"" + className + "\".");
                }
            }
            catch (ClassNotFoundException e)
            {
                throw attributeValidationFailure(tag, text, "Expected a \"" + expectedClass.getName() + "\" for \"" +
                        attrName + "\", could not find class \"" + className + "\"", e);
            }
            catch (InstantiationException | IllegalAccessException | ClassCastException e)
            {
                throw attributeValidationFailure(tag, text, "Expected a \"" + expectedClass.getName() + "\" for \"" +
                        attrName + "\", could not instantiate class \"" + className + "\": ", e);
            }
        }
        else
        {
            throw attributeValidationFailure(tag, text, "Expected a \"" + expectedClass.getName() + "\" for \"" +
                    attrName + "\", got a \"" + obj.getClass().getName() + "\": ");
        }
        return result;
    }

    /**
     * Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>List</code> out
     * of the result, parsing a delimited list to create a list if necessary.
     *
     * @param tag   The <code>Tag</code>.
     * @param text  Text which may have embedded <code>Expressions</code>.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @param def   The default value if the text is null.
     * @return A <code>List</code>.
     */
    public static List<String> evaluateList(Tag tag,
                                            RichTextString text, Map<String, Object> beans, List<String> def)
    {
        List<String> result;
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text.toString());
        if (obj instanceof List)
        {
            List list = (List) obj;
            result = new ArrayList<>(list.size());
            for (Object item : list)
                result.add(item.toString());
        }
        else
        {
            String[] items = obj.toString().split(SPEC_SEP);
            result = Arrays.asList(items);
        }
        return result;
    }

    /**
     * <p>Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>List</code> of
     * <code>Integers</code> from the result, accepting an <code>int</code>
     * array or a <code>Collection</code> or delimited list of numbers.</p>
     * <p>Examples of proper input:</p>
     * <ul>
     * <li>[0, 1, 2]
     * <li>(ArrayList){0, 1, 2}
     * <li>"0; 1; 2"
     * </ul>
     *
     * @param tag   The <code>Tag</code>.
     * @param text  Text which may have embedded <code>Expressions</code>.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @param def   The default value if the text is null.
     * @return A <code>List</code> of <code>Integers</code>.
     */
    public static List<Integer> evaluateIntegerArray(Tag tag,
                                                     RichTextString text, Map<String, Object> beans, List<Integer> def)
    {
        List<Integer> result = new ArrayList<>();
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text.toString());
        if (obj instanceof int[])
        {
            int[] intArray = (int[]) obj;
            for (int i : intArray)
                result.add(i);
        }
        else if (obj instanceof Integer[])
        {
            Integer[] intArray = (Integer[]) obj;
            result.addAll(Arrays.asList(intArray));
        }
        else if (obj instanceof Collection)
        {
            Collection c = (Collection) obj;

            for (Object o : c)
            {
                if (o instanceof Number)
                {
                    result.add(((Number) o).intValue());
                }
                else
                {
                    try
                    {
                        result.add(Integer.parseInt(o.toString()));
                    }
                    catch (NumberFormatException e)
                    {
                        throw attributeValidationFailure(tag, text.toString(),
                                "Expected an integer, got " + o.toString(), e);
                    }
                }
            }
        }
        else
        {
            String[] items = obj.toString().split(SPEC_SEP);
            for (String item : items)
            {
                try
                {
                    result.add(Integer.parseInt(item));
                }
                catch (NumberFormatException e)
                {
                    throw attributeValidationFailure(tag, text.toString(),
                            "Expected an integer, got " + item, e);
                }
            }
        }

        return result;
    }

    /**
     * <p>Evaluates the given text, which may have embedded
     * <code>Expressions</code>, and attempts to extract a <code>List</code> of
     * <code>Lists</code> of <code>Integers</code> from the result, accepting a
     * 2D <code>int</code> array or a <code>Collection</code> of
     * <code>Collections</code> or delimited list of numbers.</p>
     * <p>Examples of proper input:</p>
     * <ul>
     * <li>[[0, 1], [2]]
     * <li>(ArrayList){(ArrayList){0, 1}, (ArrayList){2}}
     * <li>"0, 1; 2"
     * </ul>
     *
     * @param tag   The <code>Tag</code>.
     * @param text  Text which may have embedded <code>Expressions</code>.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @param def   The default value if the text is null.
     * @return A <code>List</code> of <code>Lists</code> of
     * <code>Integers</code>.
     */
    public static List<List<Integer>> evaluateIntegerArrayArray(Tag tag,
                                                                RichTextString text, Map<String, Object> beans, List<List<Integer>> def)
    {
        List<List<Integer>> result = new ArrayList<>();
        if (text == null)
            return def;
        Object obj = Expression.evaluateString(text.toString(), tag.getWorkbookContext().getExpressionFactory(), beans);
        if (obj == null)
            throw nullValueOrExpectedVariableMissing(tag, text.toString());
        if (obj instanceof int[][])
        {
            int[][] intArray = (int[][]) obj;
            for (int[] array : intArray)
            {
                List<Integer> innerList = new ArrayList<>();
                for (int i : array)
                    innerList.add(i);
                result.add(innerList);
            }
        }
        else if (obj instanceof Integer[][])
        {
            Integer[][] intArray = (Integer[][]) obj;
            for (Integer[] array : intArray)
            {
                List<Integer> innerList = new ArrayList<>();
                innerList.addAll(Arrays.asList(array));
                result.add(innerList);
            }
        }
        else if (obj instanceof Collection)
        {
            Collection c = (Collection) obj;

            for (Object o : c)
            {
                List<Integer> innerList = new ArrayList<>();
                if (o instanceof Collection)
                {
                    Collection inner = (Collection) o;
                    for (Object innerObj : inner)
                    {
                        if (innerObj instanceof Number)
                        {
                            innerList.add(((Number) innerObj).intValue());
                        }
                        else
                        {
                            try
                            {
                                innerList.add(Integer.parseInt(innerObj.toString()));
                            }
                            catch (NumberFormatException e)
                            {
                                throw attributeValidationFailure(tag, text.toString(),
                                        "Expected an integer, got " + o.toString(), e);
                            }
                        }
                    }
                }
                result.add(innerList);
            }
        }
        else
        {
            String[] items = obj.toString().split(SPEC_SEP);
            for (String item : items)
            {
                List<Integer> innerList = new ArrayList<>();
                String[] innerItems = item.split(SPEC_SEP_2);
                for (String innerItem : innerItems)
                {
                    try
                    {
                        innerList.add(Integer.parseInt(innerItem));
                    }
                    catch (NumberFormatException e)
                    {
                        throw attributeValidationFailure(tag, text.toString(),
                                "Expected an integer, got " + item, e);
                    }
                }
                result.add(innerList);
            }
        }

        return result;
    }

    /**
     * Ensures that exactly one of the given attribute values exists.
     *
     * @param tag        The <code>Tag</code>.
     * @param attrValues A <code>List</code> of attribute values.
     * @param attrNames  A <code>List</code> of attribute names.
     * @throws AttributeExpressionException If none of the attribute values is not null, or
     *                                      if more than one attribute value is not null.
     */
    public static void ensureExactlyOneExists(Tag tag,
                                              List<RichTextString> attrValues, List<String> attrNames)
    {
        int exists = 0;
        for (RichTextString text : attrValues)
        {
            if (text != null)
            {
                exists++;
                if (exists > 1)
                {
                    throw attributeValidationFailure(tag, attrNames.toString(),
                            "Exactly one attribute must be specified");
                }
            }
        }
        if (exists != 1)
        {
            throw attributeValidationFailure(tag, attrNames.toString(),
                    "Exactly one attribute must be specified");
        }
    }

    /**
     * Ensures that at most one of the given attribute values exists.
     *
     * @param tag        The <code>Tag</code>.
     * @param attrValues A <code>List</code> of attribute values.
     * @param attrNames  A <code>List</code> of attribute names.
     * @throws AttributeExpressionException If more than one of the attribute values is not
     *                                      null.
     * @since 0.4.0
     */
    public static void ensureAtMostOneExists(Tag tag,
                                             List<RichTextString> attrValues, List<String> attrNames)
    {
        int exists = 0;
        for (RichTextString text : attrValues)
        {
            if (text != null)
            {
                exists++;
                if (exists > 1)
                {
                    throw attributeValidationFailure(tag, attrNames.toString(),
                            "At most one attribute must be specified");
                }
            }
        }
        if (exists != 1 && exists != 0)
        {
            throw attributeValidationFailure(tag, attrNames.toString(),
                    "At most one attribute must be specified");
        }
    }

    /**
     * Ensures that at least one of the given attribute values exists.
     *
     * @param tag        The <code>Tag</code>.
     * @param attrValues A <code>List</code> of attribute values.
     * @param attrNames  A <code>List</code> of attribute names.
     * @throws AttributeExpressionException If all of the attribute values are null.
     * @since 0.4.0
     */
    public static void ensureAtLeastOneExists(Tag tag,
                                              List<RichTextString> attrValues, List<String> attrNames)
    {
        for (RichTextString text : attrValues)
        {
            if (text != null)
                return;
        }
        throw attributeValidationFailure(tag, attrNames.toString(), "At least one attribute must be specified");
    }
}
