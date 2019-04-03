package net.sf.jett.tag;

import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;

import java.util.Map;

/**
 * <p>A <code>RickrollTag</code> is a <code>HyperlinkTag</code> that forces the
 * type to be "url" and the address to be a URL that shows Rick Astley's "Never
 * Gonna Give You Up" video.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>value (required): <code>RichTextString</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class RickrollTag extends HyperlinkTag
{
    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "rickroll";
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  This tag must be
     * bodiless.
     */
    @Override
    public void checkAttributes()
    {
        Map<String, RichTextString> attributes = getAttributes();
        CreationHelper helper = getContext().getSheet().getWorkbook().getCreationHelper();
        attributes.put(ATTR_TYPE, helper.createRichTextString(TYPE_URL));
        attributes.put(ATTR_ADDRESS, helper.createRichTextString("http://www.youtube.com/watch?v=dQw4w9WgXcQ"));

        super.checkAttributes();
    }
}
