package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * <p>A <code>SetTag</code> has no direct effect on the resultant spreadsheet.
 * It stores a new variable-value pair in the current beans map.  The variable
 * name must be a legal variable name in JEXL.  Set the value to be any legal
 * JEXL expression.  This tag must be bodiless.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>var (required): <code>String</code></li>
 * <li>value (required): <code>Object</code></li>
 * <li>display (optional): <code>RichTextString</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.11.0
 */
public class SetTag extends BaseTag
{
    /**
     * Attribute that specifies the name of the variable in the beans map to
     * create or modify.
     */
    public static final String ATTR_VAR = "var";
    /**
     * Attribute that specifies the value to set -- any object that can be
     * specified in a legal JEXL expression.
     */
    public static final String ATTR_VALUE = "value";
    /**
     * Attribute that specifies the display value of the cell after this tag
     * has been processed, if any.
     */
    public static final String ATTR_DISPLAY = "display";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_VAR, ATTR_VALUE));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_DISPLAY));

    private String myVarName;
    private Object myValue;
    private RichTextString myDisplay;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "set";
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
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  The "var" and
     * "value" attributes must be present.
     */
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException("Set tags must not have a body.  Set tag with body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        myVarName = AttributeUtil.evaluateStringVarName(this, attributes.get(ATTR_VAR), beans, ATTR_VAR);
        myValue = AttributeUtil.evaluateObject(this, attributes.get(ATTR_VALUE), beans, ATTR_VALUE, Object.class, null);
        myDisplay = attributes.get(ATTR_DISPLAY);
    }

    /**
     * Sets the given variable name to the given value.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Block block = context.getBlock();
        Sheet sheet = context.getSheet();
        int left = block.getLeftColNum();
        int top = block.getTopRowNum();

        beans.put(myVarName, myValue);

        // Set the display value.
        Row row = sheet.getRow(top);
        Cell cell = row.getCell(left);
        WorkbookContext workbookContext = getWorkbookContext();
        SheetUtil.setCellValue(workbookContext, cell, myDisplay);

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, workbookContext);

        return true;
    }
}
