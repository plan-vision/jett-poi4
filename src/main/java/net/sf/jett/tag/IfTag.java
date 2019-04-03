package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>An <code>IfTag</code> represents a conditionally placed
 * <code>Block</code> of <code>Cells</code>.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>test (required): <code>boolean</code></li>
 * <li>then (optional, bodiless only): <code>RichTextString</code></li>
 * <li>else (optional, bodiless only): <code>RichTextString</code></li>
 * <li>elseAction (optional, body only): <code>String</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class IfTag extends BaseTag
{
    /**
     * Value for the "elseAction" attribute indicating to remove the block by
     * shifting cells up, if the test condition is false.
     */
    public static final String ELSE_ACTION_SHIFT_UP = "shiftup";
    /**
     * Value for the "elseAction" attribute indicating to remove the block by
     * shifting cells left, if the test condition is false.
     */
    public static final String ELSE_ACTION_SHIFT_LEFT = "shiftleft";
    /**
     * Value for the "elseAction" attribute indicating to remove the block by
     * clearing cell contents and not shifting cells, if the test condition is
     * false.
     */
    public static final String ELSE_ACTION_CLEAR = "clear";
    /**
     * Value for the "elseAction" attribute indicating to clear the block by
     * remove the cells, but not shifting other cells, if the test condition is
     * false.
     */
    public static final String ELSE_ACTION_REMOVE = "remove";

    /**
     * Attribute for specifying the <code>boolean</code> test condition.
     */
    public static final String ATTR_TEST = "test";
    /**
     * Attribute for specifying the value of the <code>Cell</code> if the
     * condition is <code>true</code> (bodiless if-tag only).
     */
    public static final String ATTR_THEN = "then";
    /**
     * Attribute for specifying the value of the <code>Cell</code> if the
     * condition is <code>false</code> (bodiless if-tag only).
     */
    public static final String ATTR_ELSE = "else";
    /**
     * Attribute for specifying the action to be taken if the condition is
     * <code>false</code> (if-tags with a body only).
     */
    public static final String ATTR_ELSE_ACTION = "elseAction";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_TEST));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ELSE_ACTION));
    private static final List<String> REQ_ATTRS_BODILESS =
            new ArrayList<>(Arrays.asList(ATTR_TEST, ATTR_THEN));
    private static final List<String> OPT_ATTRS_BODILESS =
            new ArrayList<>(Arrays.asList(ATTR_ELSE));

    private String myElseAction;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "if";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = new ArrayList<>(super.getRequiredAttributes());
        if (isBodiless())
            reqAttrs.addAll(REQ_ATTRS_BODILESS);
        else
            reqAttrs.addAll(REQ_ATTRS);
        return reqAttrs;
    }

    /**
     * Returns a <code>List</code> of optional attribute names.
     * @return A <code>List</code> of optional attribute names.
     */
    @Override
    protected List<String> getOptionalAttributes()
    {
        List<String> optAttrs = new ArrayList<>(super.getOptionalAttributes());
        if (isBodiless())
            optAttrs.addAll(OPT_ATTRS_BODILESS);
        else
            optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  Some optional
     * attributes are only valid for bodiless tags, and others are only valid
     * for tags without bodies.
     */
    @Override
    public void validateAttributes()
    {
        super.validateAttributes();
        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();
        Block block = context.getBlock();

        String elseAction = AttributeUtil.evaluateStringSpecificValues(this, attributes.get(ATTR_ELSE_ACTION), beans,
                ATTR_ELSE_ACTION,
                Arrays.asList(ELSE_ACTION_SHIFT_UP, ELSE_ACTION_SHIFT_LEFT, ELSE_ACTION_CLEAR, ELSE_ACTION_REMOVE),
                ELSE_ACTION_SHIFT_UP);
        if (elseAction != null)
        {
            if (ELSE_ACTION_SHIFT_UP.equalsIgnoreCase(elseAction))
                block.setDirection(Block.Direction.VERTICAL);
            else if (ELSE_ACTION_SHIFT_LEFT.equalsIgnoreCase(elseAction))
                block.setDirection(Block.Direction.HORIZONTAL);
            else if (ELSE_ACTION_CLEAR.equalsIgnoreCase(elseAction) ||
                    ELSE_ACTION_REMOVE.equalsIgnoreCase(elseAction))
                block.setDirection(Block.Direction.NONE);

            myElseAction = elseAction;
        }
    }

    /**
     * <p>Evaluate the condition.</p>
     * <p>With Body: If it's true, transform the block of <code>Cells</code>.
     * If it's false, take the "elseAction", which defaults to removing the
     * block.</p>
     * <p>Bodiless: If it's true, evaluate the "then" condition.  If it's false,
     * evaluate the "else" condition, which defaults to a value of null.</p>
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();
        Map<String, Object> beans = context.getBeans();
        WorkbookContext workbookContext = getWorkbookContext();

        Map<String, RichTextString> attributes = getAttributes();

        boolean condition = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_TEST), beans, true);

        if (isBodiless())
        {
            RichTextString result;
            if (condition)
                result = attributes.get(ATTR_THEN);
            else
                result = attributes.get(ATTR_ELSE);
            // Replace the bodiless tag text with the proper result.
            Row row = sheet.getRow(block.getTopRowNum());
            Cell cell = row.getCell(block.getLeftColNum());
            SheetUtil.setCellValue(workbookContext, cell, result, result);

            BlockTransformer transformer = new BlockTransformer();
            transformer.transform(context, workbookContext);
        }
        else
        {
            if (condition)
            {
                BlockTransformer transformer = new BlockTransformer();
                transformer.transform(context, workbookContext);
            }
            else
            {
                if (ELSE_ACTION_CLEAR.equals(myElseAction))
                    clearBlock();
                else
                    removeBlock();  // Takes care of remove, shiftLeft, and shiftUp.
                return false;
            }
        }
        return true;
    }
}
