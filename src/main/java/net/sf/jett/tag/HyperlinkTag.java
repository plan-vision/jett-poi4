package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>HyperlinkTag</code> represents a Cell that needs to have a
 * hyperlink on the cell.  It controls Hyperlink properties such link type, the
 * link address, and the link label.  Because Excel won't allow other text
 * besides the Hyperlink in the Cell, any text in the Cell but outside of the
 * Hyperlink tag will be removed when the Hyperlink is created.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>type (optional): <code>String</code></li>
 * <li>address (required): <code>String</code></li>
 * <li>value (required): <code>RichTextString</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class HyperlinkTag extends BaseTag
{
    /**
     * Attribute that specifies the link type to be created, which could be a
     * web address, an email address, a document reference, or a file reference.
     * This defaults to a URL link.
     * @see #TYPE_URL
     * @see #TYPE_EMAIL
     * @see #TYPE_FILE
     * @see #TYPE_DOC
     */
    public static final String ATTR_TYPE = "type";
    /**
     * Attribute that specifies the address of the link, e.g. a web address, an
     * email address, a document reference ("'Some Sheet'!A1"), or a filename
     * ("test.xlsx").
     */
    public static final String ATTR_ADDRESS = "address";
    /**
     * Attribute that specifies the value of the cell, which is the label for
     * the link.
     */
    public static final String ATTR_VALUE = "value";

    /**
     * The "type" value indicating a web address with a URL.
     */
    public static final String TYPE_URL = "url";
    /**
     * The "type" value indicating an email link with an email address.
     */
    public static final String TYPE_EMAIL = "email";
    /**
     * The "type" value indicating a file link with a pathname.
     */
    public static final String TYPE_FILE = "file";
    /**
     * The "type" value indicating a document link, with a cell reference.
     */
    public static final String TYPE_DOC = "doc";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ADDRESS, ATTR_VALUE));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_TYPE));

    private HyperlinkType myLinkType;
    private String myAddress;
    private RichTextString myValue;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "hyperlink";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
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
     * @return A <code>List</code> of optional attribute names.
     */
    @Override
    protected List<String> getOptionalAttributes()
    {
        List<String> optAttrs =new ArrayList<>(super.getOptionalAttributes());
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  This tag must be
     * bodiless.  The type must be valid.
     */
    @Override
    @SuppressWarnings("unchecked")
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException(getName() + " tags must not have a body.  Body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        String type = AttributeUtil.evaluateStringSpecificValues(this, attributes.get(ATTR_TYPE), beans, ATTR_TYPE,
                Arrays.asList(TYPE_URL, TYPE_EMAIL, TYPE_FILE, TYPE_DOC), TYPE_URL);
        if (TYPE_URL.equals(type))
            myLinkType = HyperlinkType.URL;
        else if (TYPE_EMAIL.equals(type))
            myLinkType = HyperlinkType.EMAIL;
        else if (TYPE_FILE.equals(type))
            myLinkType = HyperlinkType.FILE;
        else if (TYPE_DOC.equals(type))
            myLinkType = HyperlinkType.DOCUMENT;

        myAddress = AttributeUtil.evaluateStringNotNull(this, attributes.get(ATTR_ADDRESS), beans, ATTR_ADDRESS, null);

        myValue = attributes.get(ATTR_VALUE);
    }

    /**
     * <p>Place the Hyperlink in the Cell, which replaces any other value left
     * behind in the Cell.</p>
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
        // It should exist in this Cell; this Tag was found in it.
        Row row = sheet.getRow(top);
        Cell cell = row.getCell(left);
        WorkbookContext workbookContext = getWorkbookContext();
        SheetUtil.setCellValue(workbookContext, cell, myValue);

        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        Hyperlink hyperlink = helper.createHyperlink(myLinkType);
        hyperlink.setAddress(myAddress);
        cell.setHyperlink(hyperlink);

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, workbookContext);

        return true;
    }
}
