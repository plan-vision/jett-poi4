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
 * <p>A <code>PageBreakTag</code> turns on or turns off whether a column or a
 * row has a page break on it.  It must be bodiless.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>type (required): <code>String</code>
 *     <ul>
 *         <li><strong>none</strong> - Clear any row and column breaks to the right/bottom of this cell.</li>
 *         <li><strong>rows</strong> - Set a row break and clear any column break to the right/bottom of this cell.</li>
 *         <li><strong>cols</strong> - Clear a row break and set a column break to the right/bottom of this cell.</li>
 *         <li><strong>both</strong> - Set a row break and a column break to the right/bottom of this cell.</li>
 *     </ul>
 * </li>
 * <li>display (optional): <code>RichTextString</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.11.0
 */
public class PageBreakTag extends BaseTag
{
    /**
     * Attribute that specifies which page breaks to clear and/or set at the
     * bottom and to the right of the current cell.
     */
    public static final String ATTR_TYPE = "type";
    /**
     * Attribute that specifies the display value of the cell after this tag
     * has been processed, if any.
     */
    public static final String ATTR_DISPLAY = "display";

    /**
     * Attribute value that specifies to clear both any row break below the
     * current cell and any column break to the right of the current cell.
     */
    public static final String TYPE_NONE = "none";
    /**
     * Attribute value that specifies to set a row break below the
     * current cell and clear any column break to the right of the current cell.
     */
    public static final String TYPE_ROWS = "rows";
    /**
     * Attribute value that specifies to clear any row break below the
     * current cell and set a column break to the right of the current cell.
     */
    public static final String TYPE_COLS = "cols";
    /**
     * Attribute value that specifies to set a row break below the
     * current cell and a column break to the right of the current cell.
     */
    public static final String TYPE_BOTH = "both";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_TYPE));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_DISPLAY));

    private String myType;
    private RichTextString myDisplay;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "pageBreak";
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
     * Validates the attributes for this <code>Tag</code>.  The "type"
     * attribute must be present.
     */
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException("PageBreak tags must not have a body.  PageBreak tag with body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        myType = AttributeUtil.evaluateStringSpecificValues(this, attributes.get(ATTR_TYPE), beans, ATTR_TYPE,
                Arrays.asList(TYPE_NONE, TYPE_ROWS, TYPE_COLS, TYPE_BOTH), null);
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
        Block block = context.getBlock();
        Sheet sheet = context.getSheet();
        int left = block.getLeftColNum();
        int top = block.getTopRowNum();

        if (TYPE_BOTH.equalsIgnoreCase(myType) || TYPE_ROWS.equalsIgnoreCase(myType))
        {
            sheet.setRowBreak(top);
        }
        if (TYPE_BOTH.equalsIgnoreCase(myType) || TYPE_COLS.equalsIgnoreCase(myType))
        {
            sheet.setColumnBreak(left);
        }
        if (TYPE_NONE.equalsIgnoreCase(myType) || TYPE_ROWS.equalsIgnoreCase(myType))
        {
            sheet.removeColumnBreak(left);
        }
        if (TYPE_NONE.equalsIgnoreCase(myType) || TYPE_COLS.equalsIgnoreCase(myType))
        {
            sheet.removeRowBreak(top);
        }

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
