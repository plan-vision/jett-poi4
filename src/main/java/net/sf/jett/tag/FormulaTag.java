package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.expression.Expression;
import net.sf.jett.expression.ExpressionFactory;
import net.sf.jett.model.Block;
import net.sf.jett.util.AttributeUtil;

/**
 * <p>A <code>FormulaTag</code> represents a dynamically generated Excel
 * Formula.  The <code>text</code> attribute contains formula text that may
 * contain Expressions that are dynamically evaluated.  JETT does not verify
 * that the dynamically generated expression is a valid Excel Formula.  A
 * <code>FormulaTag</code> must be bodiless.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>text or bean (required): <code>String</code></li>
 * <li>ifError (optional): <code>String</code></li>
 * </ul>
 *
 * <p>Either "text" or "bean" must be specified, but not both.</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 */
public class FormulaTag extends BaseTag
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Attribute that specifies the bean name that contains the Expression
     * string to be evaluated by JETT that results in formula creation in the
     * cell.  I.e. <code>bean="beanName" =&gt; "${beanName}" =&gt; "formula"</code>.
     * Then the value is used as the formula expression, i.e. <code>formula =&gt;
     * "${wins} + ${losses}" =&gt; "A2 + B2"</code>, which is used as the formula
     * text.  Either "bean" or "text" must be specified, but not both.
     */
    public static final String ATTR_BEAN = "bean";
    /**
     * Attribute that specifies the Expression string to be evaluated by JETT
     * that results in formula creation in the cell.  I.e.
     * <code>text="${wins} + ${losses}" =&gt; "A2 + B2"</code>, which is used as
     * the formula text.  Either "bean" or "text" must be specified, but not
     * both.
     */
    public static final String ATTR_TEXT = "text";
    /**
     * Attribute that specifies the Expression string to be evaluated by JETT
     * to be used as alternative text in case an Excel error results.
     */
    public static final String ATTR_IF_ERROR = "ifError";

    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_IF_ERROR, ATTR_TEXT, ATTR_BEAN));

    private String myFormulaExpression;
    private String myIfErrorExpression;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "formula";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        return super.getRequiredAttributes();
    }

    /**
     * Returns a <code>List</code> of optional attribute names.
     * @return A <code>List</code> of optional attribute names.
     */
    @Override
    protected List<String> getOptionalAttributes()
    {
        List<String> optAttrs = new ArrayList<>(super.getOptionalAttributes());
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  This tag must be
     * bodiless.
     */
    @SuppressWarnings("unchecked")
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException("Formula tags must not have a body.  Formula tag with body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        RichTextString formulaBean = attributes.get(ATTR_BEAN);
        RichTextString formulaText = attributes.get(ATTR_TEXT);

        AttributeUtil.ensureExactlyOneExists(this, Arrays.asList(formulaBean, formulaText), Arrays.asList(ATTR_BEAN, ATTR_TEXT));
        if (formulaBean != null)
        {
            myFormulaExpression = Expression.evaluateString(
                    "${" + formulaBean.toString() + "}", getWorkbookContext().getExpressionFactory(), beans)
                    .toString();
        }
        else if (formulaText != null)
        {
            myFormulaExpression = attributes.get(ATTR_TEXT).getString();
        }

        logger.debug("myFormulaExpression = {}", myFormulaExpression);

        RichTextString rtsIfError = attributes.get(ATTR_IF_ERROR);
        myIfErrorExpression = (rtsIfError != null) ? rtsIfError.getString() : null;
    }

    /**
     * <p>Evaluate the "text" attribute, and place the resultant text in an
     * Excel Formula.  If the "ifError" attribute is supplied, then wrap the
     * formula in an "IF(ISERROR(formula), ifError, formula)" formula.</p>
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        ExpressionFactory factory = getWorkbookContext().getExpressionFactory();
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();
        Map<String, Object> beans = context.getBeans();
        int left = block.getLeftColNum();
        int top = block.getTopRowNum();
        // It should exist in this Cell; this Tag was found in it.
        Row row = sheet.getRow(top);
        Cell cell = row.getCell(left);

        logger.debug("myFormulaExpression: {}", myFormulaExpression);
        String formulaText = Expression.evaluateString(myFormulaExpression, factory, beans).toString();
        logger.debug("formulaText: {}", formulaText);
        if (myIfErrorExpression != null)
        {
            Object errorResult = Expression.evaluateString(myIfErrorExpression, factory, beans);
            logger.debug("errorResult: {}", errorResult);
            String newFormulaText = "IF(ISERROR(" + formulaText + "), ";
            // Don't quote numbers!
            if (!(errorResult instanceof Number))
                newFormulaText += "\"";
            newFormulaText += errorResult.toString();
            if (!(errorResult instanceof Number))
                newFormulaText += "\"";

            newFormulaText += ", " + formulaText + ")";

            formulaText = newFormulaText;
        }

        logger.debug("  Formula for row {}, cell {} is {}", top, left, formulaText);
        cell.setCellFormula(formulaText);

        return true;
    }
}
