package net.sf.jett.expression;

import java.io.StringReader;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.parser.ASTIdentifier;
import org.apache.commons.jexl2.parser.ASTMethodNode;
import org.apache.commons.jexl2.parser.ASTNumberLiteral;
import org.apache.commons.jexl2.parser.ASTReference;
import org.apache.commons.jexl2.parser.ASTSizeMethod;
import org.apache.commons.jexl2.parser.Node;
import org.apache.commons.jexl2.parser.Parser;
import org.apache.commons.jexl2.parser.SimpleNode;
import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jett.exception.ParseException;
import net.sf.jett.formula.Formula;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.util.FormulaUtil;
import net.sf.jett.util.RichTextStringUtil;

/**
 * <p>An <code>Expression</code> object represents a JEXL Expression that can
 * be evaluated given a <code>Map</code> of bean names to values.  Many
 * <code>Expressions</code> may be created in a cell.  Here in JETT,
 * <code>Expressions</code> are built from the text found in between "${" and
 * "}".</p>
 *
 * @author Randy Gettman
 */
public class Expression
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Contains a cache of collection names found in expression texts.  If it is
     * known that there are no collection names, then the value will be an empty
     * String, to distinguish from the case in which the result is not known
     * yet, in which case the result is <code>null</code>.
     */
    private static final Map<String, String> MAP_EXPRESSION_TO_COLL_NAMES = new HashMap<>();

    /**
     * This pattern makes sure that there is no backslash in front of an
     * expression that is due to be replaced with the result of its evaluation.
     * @since 0.8.0
     */
    public static final String NEGATIVE_LOOKBEHIND_BACKSLASH = "(?<![\\\\])";

    /**
     * Determines the start of a JEXL expression.
     */
    public static final String BEGIN_EXPR = "${";
    /**
     * Determines the end of a JEXL expression.
     */
    public static final String END_EXPR = "}";

    private String myExpression;

    /**
     * Create an <code>Expression</code>.
     * @param expression The expression in String form.
     */
    public Expression(String expression)
    {
        myExpression = expression;
    }

    /**
     * Evaluate this <code>Expression</code> using the given <code>Map</code> of
     * beans as a context.
     * @param factory An <code>ExpressionFactory</code>.
     * @param beans A <code>Map</code> mapping strings to objects.
     * @return The result of the evaluation.
     */
    @SuppressWarnings("unchecked")
    public Object evaluate(ExpressionFactory factory, Map<String, Object> beans)
    {
        if (beans != null && !beans.isEmpty())
        {
            JexlContext context = new ClassAwareMapContext(beans);
            return factory.createExpression(myExpression).evaluate(context);
        }
        return myExpression;
    }

    /**
     * Find all <code>ASTReferences</code> in the tree.  Calls itself recursively.
     * @param node The <code>Node</code>.
     * @return A <code>List</code> of <code>ASTReferences</code>.
     */
    private List<ASTReference> findReferences(Node node)
    {
        List<ASTReference> references = new ArrayList<>();
        if (node instanceof ASTReference)
        {
            references.add((ASTReference) node);
        }

        int count = node.jjtGetNumChildren();
        for (int i = 0; i < count; i++)
        {
            references.addAll(findReferences(node.jjtGetChild(i)));
        }
        return references;
    }

    /**
     * Determine if any substring starting at the beginning of the given
     * <code>ASTReference</code> evaluates to a <code>Collection</code>.  If so,
     * then return that substring, which is the name of the
     * <code>Collection</code>.  If there is no such substring, then return
     * <code>null</code>.
     * @param node The <code>ASTReference</code>.
     * @param beans The <code>Map</code> of beans.
     * @param context A <code>WorkbookContext</code>, which refers to an
     *    <code>ExpressionFactory</code> and a <code>List</code> of collection
     *    names. Don't return a collection expression whose collection property
     *    name is found in this <code>List</code>.
     * @return The full reference string to the collection name, or
     *    <code>null</code> if there is no collection.
     */
    private String findCollectionName(ASTReference node, Map<String, Object> beans,
                                      WorkbookContext context)
    {
        ExpressionFactory factory = context.getExpressionFactory();
        List<String> noImplProcCollNames = context.getNoImplicitProcessingCollectionNames();
        int count = node.jjtGetNumChildren();
        String collectionName = null;
        for (int i = 0; i < count; i++)
        {
            Node child = node.jjtGetChild(i);
            if (child instanceof ASTIdentifier)
            {
                ASTIdentifier identifier = (ASTIdentifier) child;

                if (collectionName == null)
                    collectionName = identifier.image;
                else
                    collectionName = collectionName + "." + identifier.image;
                logger.debug("    fCN: Test Expr ({}/{}): \"{}\".",
                        i, count, collectionName);

                // Turn off implicit collection processing on a per-collection name
                // basis.
                if (noImplProcCollNames.contains(identifier.image))
                {
                    logger.trace("    fCN: Skipping because {} has been turned off.", identifier.image);
                    continue;
                }

                Expression expr = new Expression(collectionName);
                Object result = expr.evaluate(factory, beans);
                if (result instanceof Collection)
                {
                    // Continue past a Collection if the next method called is a
                    // method that is:
                    // 1. side-effect free
                    // 2. returns something (not void) that is not itself a
                    //    Collection, array, or Iterator.
                    // This is done by looking at the next child (if it exists) and
                    // determining if it is one of a certain "family" of methods
                    // commonly expected to be called on Collections.
                    if (i < count - 1)
                    {
                        Node nextChild = node.jjtGetChild(i + 1);
                        if (nextChild instanceof ASTMethodNode)
                        {
                            ASTMethodNode methodNode = (ASTMethodNode) nextChild;
                            if (logger.isDebugEnabled())
                            {
                                logger.debug("      fCN: method.image = {}", methodNode.image);
                                logger.debug("      fCN: method.toString = {}", methodNode.toString());
                                int numChildren = methodNode.jjtGetNumChildren();
                                logger.debug("      fCN: method.jjtGetNumChildren = {}", numChildren);
                                for (int j = 0; j < numChildren; j++)
                                {
                                    Node methodChild = methodNode.jjtGetChild(j);
                                    logger.debug("Child {} ({}): {}",
                                            j, methodChild.getClass().getName(), methodChild.toString());
                                    if (methodChild instanceof ASTIdentifier)
                                    {
                                        ASTIdentifier childIdentifier = (ASTIdentifier) methodChild;
                                        logger.debug("  child image = \"{}\".", childIdentifier.image);
                                    }
                                }
                            }
                            // First child should be the identifier (name) of the method.
                            ASTIdentifier childIdentifier = (ASTIdentifier) methodNode.jjtGetChild(0);
                            if (childIdentifier.image != null &&
                                    (childIdentifier.image.startsWith("capacity") ||
                                            childIdentifier.image.startsWith("contains") ||
                                            childIdentifier.image.startsWith("element") ||
                                            childIdentifier.image.startsWith("equals") ||
                                            childIdentifier.image.equals("get") ||  // Don't cover getter methods that may return Collections
                                            childIdentifier.image.startsWith("hashCode") ||
                                            childIdentifier.image.startsWith("indexOf") ||
                                            childIdentifier.image.startsWith("isEmpty") ||
                                            childIdentifier.image.startsWith("lastIndexOf") ||
                                            childIdentifier.image.startsWith("size") ||
                                            childIdentifier.image.startsWith("toString")
                                    )
                                    )
                            {
                                // Continue on to the next child (if any).
                                logger.trace("      fCN: Skipping {} because of child method name {}",
                                        collectionName, childIdentifier.image);
                                continue;
                            }
                        }
                        else if (nextChild instanceof ASTSizeMethod)
                        {
                            // Apparently, "ASTSizeMethod" is a special case for
                            // ".size()", but "ASTSizeMethod" is NOT a
                            // "ASTMethodNode", so this check is needed!
                            logger.trace("      fCN: sizeMethod.image = {}",
                                    ((ASTSizeMethod) nextChild).image);
                            continue;
                        }
                        else if (nextChild instanceof ASTNumberLiteral)
                        {
                            // JEXL allows ".n" to access an element of a List.
                            logger.trace("      fCN: numberLiteral.image = {}",
                                    ((ASTNumberLiteral) nextChild).image);
                            continue;
                        }
                        else
                        {
                            logger.debug("      fCN: Next child is a {}", nextChild.getClass().getName());
                        }
                    }
                    else
                    {
                        // No additional children.  The Expression simply evaluates
                        // to a Collection.
                        logger.trace("      fCN: Just a collection: \"{}\".", collectionName);
                        return null;
                    }

                    logger.debug("      fCN: Found collection: \"{}\".", collectionName);
                    return collectionName;
                }
            }
        }
        return null;
    }

    /**
     * <p>Determines whether this represents implicit Collections access, which
     * would result in an implicit collections processing loop.  If so, then it
     * returns the substring representing the <code>Collection</code>, else it
     * returns <code>null</code>.</p>
     * <p>This method uses JEXL internal parser logic.</p>
     *
     * @param beans A <code>Map</code> mapping strings to objects.
     * @param context A <code>WorkbookContext</code>.  Don't return a collection
     *    expression whose collection property name is found in the <code>List</code>
     *    of such names maintained by this <code>WorkbookContext</code>..
     * @return The string representing the <code>Collection</code>, or
     *    <code>null</code> if it doesn't represent implicit Collections access.
     */
    public String getValueIndicatingImplicitCollection(Map<String, Object> beans,
                                                       WorkbookContext context)
    {
        String expression = myExpression;
        // Try cache first.
        String cachedResult = MAP_EXPRESSION_TO_COLL_NAMES.get(expression);
        if (cachedResult != null)
        {
            return cachedResult;
        }

        Parser parser = new Parser(new StringReader(";"));
        try
        {
            SimpleNode tree = parser.parse(new StringReader(expression), null);
            List<ASTReference> references = findReferences(tree);
            for (ASTReference node : references)
            {
                logger.trace("  Reference...");
                String collectionName = findCollectionName(node, beans, context);
                if (collectionName != null)
                {
                    // Cache this result.
                    MAP_EXPRESSION_TO_COLL_NAMES.put(expression, collectionName);
                    return collectionName;
                }
            }
        }
        catch (org.apache.commons.jexl2.parser.ParseException e)
        {
            throw new ParseException("JEXL ParseException caught on expression \"" + expression + "\": " + e.getMessage(), e);
        }
        // If we get here, then there is no Collection name reference.
        MAP_EXPRESSION_TO_COLL_NAMES.put(expression, "");
        return null;
    }

    /**
     * Clear the <code>Map</code> that is used to cache the fact that a certain
     * collection name may be present in expression text.  Call this method when
     * a new beans <code>Map</code> is being used, which would render the cache
     * useless.  Such a situation arises when supplying multiple bean maps to
     * the <code>transform</code> method on <code>ExcelTransformer</code>, and
     * we are moving to a new <code>Sheet</code>, or if either
     * <code>transform</code> method on <code>ExcelTransformer</code> is called
     * more than once.
     */
    public static void clearExpressionToCollNamesMap()
    {
        MAP_EXPRESSION_TO_COLL_NAMES.clear();
    }

    /**
     * Determines whether a string representing an <code>Expression</code>
     * represents implicit Collections access, which would result in an implicit
     * collections processing loop.  If so, then it returns the substring
     * representing the <code>Collection</code>, else it returns
     * <code>null</code>.
     * @param value The string possibly representing an <code>Expression</code>.
     * @param beans A <code>Map</code> mapping strings to objects.
     * @param context A <code>WorkbookContext</code>, which supplies a
     *    <code>List</code> of collection names to ignore and the
     *    <code>ExpressionFactory</code>.  Don't return a collection expression
     *    whose collection property name is found in this <code>List</code>.
     * @return A <code>List</code> of strings representing the
     *    <code>Collections</code> found, possibly empty if it doesn't represent
     *    implicit Collections access.
     */
    public static List<String> getImplicitCollectionExpr(String value, Map<String, Object> beans,
                                                         WorkbookContext context)
    {
        ExpressionFactory factory = context.getExpressionFactory();

        logger.trace("getImplicitCollectionExpr: \"{}\".", value);
        List<Expression> expressions = getExpressions(value);
        List<String> implicitCollections = new ArrayList<>();

        // Don't report errors for some identifiers that depend on implicit
        // processing to be a legal expression, e.g. a property access on a List
        // meant to be a property access on an element of the List.  Store the
        // current silent/lenient flags for restoration later.
        boolean lenient = factory.isLenient();
        boolean silent = factory.isSilent();
        factory.setLenient(true);
        factory.setSilent(true);

        if (value.startsWith(Expression.BEGIN_EXPR) && value.endsWith(Expression.END_EXPR) && expressions.size() == 1)
        {
            Expression expression = new Expression(value.substring(2, value.length() - 1));
            String implColl = expression.getValueIndicatingImplicitCollection(beans, context);
            if (implColl != null && !"".equals(implColl))
                implicitCollections.add(implColl);
        }
        else if (expressions.size() >= 1)
        {
            for (Expression expression : expressions)
            {
                String implColl = expression.getValueIndicatingImplicitCollection(beans, context);
                if (implColl != null && !"".equals(implColl))
                    implicitCollections.add(implColl);
            }
        }

        if (logger.isTraceEnabled())
        {
            logger.trace("  gICE implicitCollections.size() = {}", implicitCollections.size());
            for (String implColl : implicitCollections)
            {
                logger.trace("  gICE implColl item: {}", implColl);
            }
        }

        // Restore settings.
        factory.setLenient(lenient);
        factory.setSilent(silent);

        return implicitCollections;
    }

    /**
     * Find any <code>Expressions</code> embedded in the given string, evaluate
     * them, and replace the expressions with the resulting values.  If the
     * entire string consists of one <code>Expression</code>, then the returned
     * value may be any <code>Object</code>.
     *
     * @param richTextString The rich text string, with possibly embedded
     * expressions.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param factory An <code>ExpressionFactory</code>.
     * @param beans A <code>Map</code> mapping strings to objects.
     * @return A new string, with any embedded expressions replaced with the
     *    expression string values.
     */
    public static Object evaluateString(RichTextString richTextString,
                                        CreationHelper helper, ExpressionFactory factory, Map<String, Object> beans)
    {
        String value = richTextString.getString();
        List<Expression> expressions = getExpressions(value);
        if (value.startsWith(Expression.BEGIN_EXPR) && value.endsWith(Expression.END_EXPR) && expressions.size() == 1)
        {
            Expression expression = new Expression(value.substring(2, value.length() - 1));
            Object result = expression.evaluate(factory, beans);
            if (result instanceof String)
            {
                return RichTextStringUtil.replaceAll(richTextString, helper, value, (String) result, true);
            }
            else
            {
                return result;
            }
        }
        else
        {
            return replaceExpressions(richTextString, helper, expressions, factory, beans);
        }
    }

    /**
     * Find any <code>Expressions</code> embedded in the given string, evaluate
     * them, and replace the expressions with the resulting values.  If the
     * entire string consists of one <code>Expression</code>, then the returned
     * value may be any <code>Object</code>.
     *
     * @param value The string, with possibly embedded expressions.
     * @param factory An <code>ExpressionFactory</code>.
     * @param beans A <code>Map</code> mapping strings to objects.
     * @return A new string, with any embedded expressions replaced with the
     *    expression string values.
     */
    public static Object evaluateString(String value, ExpressionFactory factory, Map<String, Object> beans)
    {
        List<Expression> expressions = getExpressions(value);
        if (value.startsWith(Expression.BEGIN_EXPR) && value.endsWith(Expression.END_EXPR) && expressions.size() == 1)
        {
            Expression expression = new Expression(value.substring(2, value.length() - 1));
            return expression.evaluate(factory, beans);
        }
        else
        {
            return replaceExpressions(value, expressions, factory, beans);
        }
    }

    /**
     * Extract all <code>Expressions</code> from the given value.
     * @param value The given value.
     * @return A <code>List</code> of <code>Expressions</code>, possibly empty.
     */
    private static List<Expression> getExpressions(String value)
    {
        List<Expression> expressions = new ArrayList<>();
        int beginIdx = value.indexOf(Expression.BEGIN_EXPR);
        int endIdx = findEndOfExpression(value, beginIdx + Expression.BEGIN_EXPR.length());
        logger.debug("  getExprs: beginIdx = {}, endIdx = {}", beginIdx, endIdx);

        while (beginIdx != -1 && endIdx != -1 && endIdx > beginIdx)
        {
            int formulaBeginIdx = value.indexOf(Formula.BEGIN_FORMULA);
            int formulaEndIdx = formulaBeginIdx != -1 ?
                    FormulaUtil.getEndOfJettFormula(value, formulaBeginIdx) :
                    value.indexOf(Formula.END_FORMULA);
            boolean exprFound = true;
            // Skip escaped expressions, e.g. "\${...}".
            if (beginIdx > 0 && value.charAt(beginIdx - 1) == '\\')
            {
                exprFound = false;
            }
            // Also, ignore expressions found inside JETT Formulas, which should
            // refer to the template sheet name.  JETT Formulas should not trigger
            // implicit collections processing.
            if (formulaBeginIdx != -1 && formulaEndIdx != -1 &&
                    formulaBeginIdx < beginIdx && formulaEndIdx > endIdx)
            {
                exprFound = false;
            }
            if (exprFound)
            {
                String strExpr = value.substring(beginIdx + 2, endIdx);
                logger.debug("  Expression Found: {}", strExpr);
                Expression expr = new Expression(strExpr);
                expressions.add(expr);
            }

            beginIdx = value.indexOf(Expression.BEGIN_EXPR, endIdx + 1);
            endIdx = findEndOfExpression(value, beginIdx + Expression.BEGIN_EXPR.length());
            logger.debug("  getExprs: beginIdx = {}, endIdx = {}", beginIdx, endIdx);
        }
        return expressions;
    }

    /**
     * Replace all expressions with their evaluated results.  This attempts to
     * preserve any formatting within the <code>RichTextString</code>.
     * @param value The entire string, with possibly many expressions.
     * @param expressions A <code>List</code> of <code>Expressions</code>.
     * @param factory An <code>ExpressionFactory</code>.
     * @param beans A <code>Map</code> of beans to provide context for the
     *    <code>Expressions</code>.
     * @return A <code>String</code> with all expressions replaced with their
     *    evaluated results.
     */
    private static String replaceExpressions(String value,
                                             List<Expression> expressions, ExpressionFactory factory, Map<String, Object> beans)
    {
        // Replace Expressions with values.
        for (Expression expr : expressions)
        {
            logger.debug("replExprs: Loop for {}", expr.myExpression);
            int beginIdx = value.indexOf(Expression.BEGIN_EXPR);
            int endIdx = beginIdx + Expression.BEGIN_EXPR.length() + expr.myExpression.length();
            if (beginIdx != -1 && endIdx != -1 && endIdx > beginIdx)
            {
                String replaceMe = value.substring(beginIdx, endIdx + 1);
                Object result = expr.evaluate(factory, beans);
                String replaceWith = "";
                if (result != null)
                    replaceWith = expr.evaluate(factory, beans).toString();
                logger.debug("  Replacing \"{}\" with \"{}\".", replaceMe, replaceWith);

                // Don't replace an expression when the $ is escaped, e.g. "\${replaceMe}".
                value = value.replaceFirst(NEGATIVE_LOOKBEHIND_BACKSLASH + Pattern.quote(replaceMe),
                        Matcher.quoteReplacement(replaceWith));
                logger.debug("  value is now \"{}\".", value);
            }
            else
            {
                break;
            }
        }
        // Respect escapes of expressions.  E.g. "\${expr}" => "${expr}", unevaluated.
        value = value.replace("\\" + Expression.BEGIN_EXPR, Expression.BEGIN_EXPR);
        return value;
    }

    /**
     * Replace all expressions with their evaluated results.  This attempts to
     * preserve any formatting within the <code>RichTextString</code>.
     * @param richTextString The entire string, with possibly many expressions
     *    and possibly embedded formatting.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param expressions A <code>List</code> of <code>Expressions</code>.
     * @param factory An <code>ExpressionFactory</code>.
     * @param beans A <code>Map</code> of beans to provide context for the
     *    <code>Expressions</code>.
     * @return A <code>RichTextString</code> with all expressions replaced with
     *    their evaluated results, and formatted preserved as best as possible.
     */
    private static RichTextString replaceExpressions(RichTextString richTextString,
                                                     CreationHelper helper, List<Expression> expressions, ExpressionFactory factory, Map<String, Object> beans)
    {
        ArrayList<String> exprStrings = new ArrayList<>(expressions.size());
        ArrayList<String> exprValues = new ArrayList<>(expressions.size());
        for (Expression expr : expressions)
        {
            logger.debug("replExprsRTS: Loop for {}", expr.myExpression);
            exprStrings.add(BEGIN_EXPR + expr.myExpression + END_EXPR);
            Object result = expr.evaluate(factory, beans);
            if (result != null)
                exprValues.add(result.toString());
            else
                exprValues.add("");
            logger.debug("  replacement of \"{}\" with \"{}\".",
                    expr.myExpression, result);
        }
        return RichTextStringUtil.replaceValues(richTextString, helper, exprStrings, exprValues, false);
    }

    /**
     * Find the end of the expression, accounting for the possible presence of
     * braces inside the expression, which is allowed in JEXL syntax for things
     * like map literals, blocks, and if/for/while blocks.
     * @param value The text with embedded expressions.
     * @param startIdx The 0-based start index on which to start looking.
     * @return The 0-based index on which the expression ends, or -1 if the
     *    expression is not terminated.
     */
    private static int findEndOfExpression(String value, int startIdx)
    {
        logger.trace("    fEOE: \"{}\", startIdx: {}", value, startIdx);
        int begins = 1;
        int ends = 0;
        for (int i = startIdx; i < value.length(); i++)
        {
            char c = value.charAt(i);
            if (c == '{')
                begins++;
            else if (c == '}')
                ends++;

            if (begins == ends)
                return i;
        }
        return -1;
    }
}