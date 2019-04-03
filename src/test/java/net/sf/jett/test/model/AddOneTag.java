package net.sf.jett.test.model;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.tag.BaseTag;
import net.sf.jett.tag.TagContext;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * An <code>AddOneTag</code> is a custom <code>Tag</code> that adds 1 to the
 * numeric "value" attribute.  The main purpose of this <code>Tag</code> is to
 * demonstrate custom tags and custom tag libraries.
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em>
 * <li>value (required): <code>Number</code>
 * </ul>
 *
 * @author Randy Gettman
 */
public class AddOneTag extends BaseTag
{
    /**
     * Attribute for specifying the value.
     */
    public static final String ATTR_VALUE = "value";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_VALUE));

    private double myValue;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "addOne";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = super.getRequiredAttributes();
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
     * Validates the attributes for this <code>Tag</code>.  Some optional
     * attributes are only valid for bodiless tags, and others are only valid
     * for tags without bodies.
     */
    public void validateAttributes()
    {
        super.validateAttributes();
        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        if (!isBodiless())
            throw new TagParseException("AddOne tags must not have a body.  AddOne tag with body found at" + getLocation());

        myValue = AttributeUtil.evaluateDouble(this, attributes.get(ATTR_VALUE), beans, ATTR_VALUE, 0);
    }

    /**
     * Replace the cell's content with the value plus one.
     * @return <code>true</code>, this cell's content was processed.
     */
    public boolean process()
    {
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();

        // Replace the bodiless tag text with the proper result.
        Cell cell = sheet.getRow(block.getTopRowNum()).getCell(block.getLeftColNum());
        SheetUtil.setCellValue(getWorkbookContext(), cell, myValue + 1, getAttributes().get(ATTR_VALUE));

        return true;
    }
}
