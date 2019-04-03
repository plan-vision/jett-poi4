package net.sf.jett.tag;

import java.io.FileInputStream;
import java.io.InputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.util.AttributeUtil;

/**
 * <p>An <code>ImageTag</code> represents an image to be placed on the sheet.
 * The <code>rows</code> and <code>cols</code> attributes, if specified,
 * control how many Excel rows tall and columns wide the image is.  If not
 * specified, the image is sized according to its natural dimensions.  The
 * required <code>pathname</code> attribute is the name of the image file, to
 * be loaded relative to the current working directory.  The optional
 * <code>type</code> attribute gives the image type, which defaults to "png".</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>rows (optional): <code>int</code></li>
 * <li>cols (optional): <code>int</code></li>
 * <li>pathname (required): <code>String</code></li>
 * <li>type (optional): <code>String</code>
 *     <ul>
 *     <li><em>png</em> The image is a PNG. (default)</li>
 *     <li><em>jpeg</em> The image is a JPG.</li>
 *     <li><em>dib</em> The image is a device-independent bitmap (or a .bmp).</li>
 *     <li><em>pict</em> The image is a Mac PICT.</li>
 *     <li><em>wmf</em> The image is a Windows Metafile.</li>
 *     <li><em>emf</em> The image is an enhanced Windows Metafile.</li>
 *     </ul>
 *     </li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.10.0
 */
public class ImageTag extends BaseTag
{
    /**
     * Attribute for specifying the number of rows tall the image will be.  If
     * neither <code>rows</code> nor <code>cols</code> is specified, then the
     * image's dimensions will be preserved from the original image file.
     */
    public static final String ATTR_ROWS = "rows";
    /**
     * Attribute for specifying the number of columns wide the image will be.  If
     * neither <code>rows</code> nor <code>cols</code> is specified, then the
     * image's dimensions will be preserved from the original image file.
     */
    public static final String ATTR_COLS = "cols";
    /**
     * Attribute for specifying the pathname of the image file, relative to the
     * current working directory.
     */
    public static final String ATTR_PATHNAME = "pathname";
    /**
     * Attribute for specifying the type of the image file.  If not specified,
     * this defaults to "png".
     */
    public static final String ATTR_TYPE = "type";

    /**
     * The type for a device independent image (or .bmp).
     */
    public static final String TYPE_DIB = "dib";
    /**
     * The type for a Enhanced Windows Metafile image.
     */
    public static final String TYPE_EMF = "emf";
    /**
     * The type for a JPEG image.
     */
    public static final String TYPE_JPEG = "jpeg";
    /**
     * The type for a Mac PICT image.
     */
    public static final String TYPE_PICT = "pict";
    /**
     * The type for a PNG image.
     */
    public static final String TYPE_PNG = "png";
    /**
     * The type for a Windows Metafile image.
     */
    public static final String TYPE_WMF = "wmf";

    /**
     * The default image type, "png".
     */
    public static final String DEF_TYPE = TYPE_PNG;

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_PATHNAME));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(
                    ATTR_ROWS, ATTR_COLS, ATTR_TYPE));

    private String myPathname;
    private int myType;
    private boolean amISizing;
    private int myRows;
    private int myCols;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "image";
    }

    /**
     * Returns the <code>List</code> of required attribute names.
     * @return The <code>List</code> of required attribute names.
     */
    @Override
    public List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = new ArrayList<>(super.getRequiredAttributes());
        reqAttrs.addAll(REQ_ATTRS);
        return reqAttrs;
    }

    /**
     * Returns the <code>List</code> of optional attribute names.
     * @return The <code>List</code> of optional attribute names.
     */
    @Override
    public List<String> getOptionalAttributes()
    {
        List<String> optAttrs = new ArrayList<>(super.getOptionalAttributes());
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  The "rows" and
     * "cols" attributes, if present, must be positive integers.
     */
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException("Image tags must not have a body.  Image tag with body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        RichTextString rtsRows = attributes.get(ATTR_ROWS);
        RichTextString rtsCols = attributes.get(ATTR_COLS);
        if (rtsRows == null && rtsCols == null)
        {
            amISizing = false;
        }
        else
        {
            if (rtsRows != null)
            {
                myRows = AttributeUtil.evaluatePositiveInt(this, rtsRows, beans, ATTR_ROWS, 1);
            }
            if (rtsCols != null)
            {
                myCols = AttributeUtil.evaluatePositiveInt(this, rtsCols, beans, ATTR_COLS, 1);
            }
            amISizing = true;
        }

        myPathname = AttributeUtil.evaluateStringNotNull(this, attributes.get(ATTR_PATHNAME), beans, ATTR_PATHNAME, "");

        RichTextString rtsType = attributes.get(ATTR_TYPE);
        String type;
        if (rtsType != null)
        {
            type = AttributeUtil.evaluateStringSpecificValues(this, rtsType, beans, ATTR_TYPE,
                    Arrays.asList(TYPE_DIB, TYPE_EMF, TYPE_JPEG, TYPE_PICT, TYPE_PNG, TYPE_WMF), TYPE_PNG);
        }
        else
        {
            type = DEF_TYPE;
        }

        if (TYPE_DIB.equalsIgnoreCase(type))
            myType = Workbook.PICTURE_TYPE_DIB;
        else if (TYPE_EMF.equalsIgnoreCase(type))
            myType = Workbook.PICTURE_TYPE_EMF;
        else if (TYPE_JPEG.equalsIgnoreCase(type))
            myType = Workbook.PICTURE_TYPE_JPEG;
        else if (TYPE_PICT.equalsIgnoreCase(type))
            myType = Workbook.PICTURE_TYPE_PICT;
        else if (TYPE_PNG.equalsIgnoreCase(type))
            myType = Workbook.PICTURE_TYPE_PNG;
        else if (TYPE_WMF.equalsIgnoreCase(type))
            myType = Workbook.PICTURE_TYPE_WMF;
    }

    /**
     * Loads the image data and scales it if necessary, with the top-left
     * corner of the image at the top-left corner of this cell.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Workbook workbook = sheet.getWorkbook();
        Block block = context.getBlock();
        int left = block.getLeftColNum();
        int top = block.getTopRowNum();

        int pictIdx;
        try (InputStream is = new FileInputStream(myPathname))
        {
            byte[] imageData = IOUtils.toByteArray(is);
            pictIdx = workbook.addPicture(imageData, myType);
        }
        catch (IOException e)
        {
            throw new TagParseException("Read of pathname \"" + myPathname + "\" failed.", e);
        }

        Drawing drawing = context.getOrCreateDrawing();
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(left);
        anchor.setRow1(top);
        if (amISizing)
        {
            anchor.setCol2(left + myCols);
            anchor.setRow2(top + myRows);
        }
        Picture pict = drawing.createPicture(anchor, pictIdx);
        if (!amISizing)
        {
            pict.resize();
        }

        // Clear the cell of text; the image would presumably cover it up.
        clearBlock();
        return true;
    }
}
