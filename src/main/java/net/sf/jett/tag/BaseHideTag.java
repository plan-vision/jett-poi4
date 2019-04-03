package net.sf.jett.tag;

import org.apache.poi.ss.usermodel.RichTextString;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;

/**
 * <p>A <code>BaseHideTag</code> represents something that can be hidden conditionally.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>test (required): <code>boolean</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public abstract class BaseHideTag extends BaseTag
{
    /**
     * Attribute for specifying whether to hide what is to be hidden.
     */
    public static final String ATTR_TEST = "test";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_TEST));

    private boolean amIHiding;

    /**
     * Returns a <code>List</code> of required attribute names.
     *
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = new ArrayList<>(super.getRequiredAttributes());
        reqAttrs.addAll(REQ_ATTRS);
        return reqAttrs;
    }

    /**
     * Returns a <code>List</code> of optional attribute names.
     *
     * @return A <code>List</code> of optional attribute names.
     */
    @Override
    protected List<String> getOptionalAttributes()
    {
        return super.getOptionalAttributes();
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  This tag must have a
     * body.
     */
    @SuppressWarnings("unchecked")
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (isBodiless())
            throw new TagParseException(getName() + " tags must have a body.  Bodiless Group tag found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();
        amIHiding = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_TEST), beans, true);
    }

    /**
     * If the condition is met, hide something in the workbook.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        setHidden(amIHiding);

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(getContext(), getWorkbookContext());

        return true;
    }

    /**
     * This method is called if the condition is <code>true</code>, to hide
     * something in the workbook.
     * @param hide Whether to hide or show.
     */
    public abstract void setHidden(boolean hide);
}
