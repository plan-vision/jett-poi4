package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>NullTag</code> does nothing to its <code>Block</code> except mark
 * its Cells as processed.  It can't have any attributes in body mode.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>text (required, bodiless only): <code>RichTextString</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class NullTag extends BaseTag
{
    /**
     * Attribute that specifies the un-process text to display (bodiless only).
     */
    public static final String ATTR_TEXT = "text";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_TEXT));

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "null";
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
        return super.getOptionalAttributes();
    }

    /**
     * No validation.
     */
    @Override
    public void validateAttributes()
    {
        super.validateAttributes();
    }

    /**
     * Just mark all <code>Cells</code> in this <code>Block</code> as processed.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();
        int left = block.getLeftColNum();
        int top = block.getTopRowNum();
        WorkbookContext workbookContext = getWorkbookContext();

        if (isBodiless())
        {
            // It should exist in this Cell; this Tag was found in it.
            Row row = sheet.getRow(top);
            Cell cell = row.getCell(left);
            SheetUtil.setCellValue(workbookContext, cell, getAttributes().get(ATTR_TEXT));
        }
        else
        {
            BlockTransformer transformer = new BlockTransformer();
            transformer.transform(context, workbookContext, false);
        }  // End else of isBodiless
        return true;
    }
}
