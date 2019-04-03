package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>CommentTag</code> represents a Cell that needs to have an Excel
 * Comment attached to it.  It controls Comment properties such author, the
 * Rich Text string comment, and whether the Comment is initially visible.</p>
 *
 * <p>This tag uses the POI method <code>createDrawingPatriarch</code> to
 * create an Excel comment on the <code>Cell</code> on which the tag is
 * located.  The POI documentation warns of corrupting other "drawings" such as
 * charts and "complex" drawings when calling <code>getDrawingPatriarch</code>
 * (HSSF code).  When testing both .xls and .xlsx template spreadsheets, it
 * appears that drawings and charts do get corrupted in .xls spreadsheets, but
 * they do NOT get corrupted in .xlsx spreadsheets.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>value (required): <code>RichTextString</code></li>
 * <li>author (required): <code>String</code></li>
 * <li>comment (required): <code>RichTextString</code></li>
 * <li>visible (optional): <code>boolean</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class CommentTag extends BaseTag
{
    /**
     * Attribute that specifies the value of the Cell itself after
     * transformation.
     */
    public static final String ATTR_VALUE = "value";
    /**
     * Attribute that specifies the author of the Comment to be created.
     */
    public static final String ATTR_AUTHOR = "author";
    /**
     * Attribute that specifies the comment text.
     */
    public static final String ATTR_COMMENT = "comment";
    /**
     * Attribute that specifies whether the comment is initially visible.  Even
     * if not initially visible, the user can mouseover a cell with a comment to
     * view the comment text as a pop-up.
     */
    public static final String ATTR_VISIBLE = "visible";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_VALUE, ATTR_AUTHOR, ATTR_COMMENT));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_VISIBLE));

    private RichTextString myValue;
    private RichTextString myAuthor;
    private RichTextString myComment;
    private boolean amIVisible;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "comment";
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
            throw new TagParseException("Comment tags must not have a body.  Comment tag with body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        myValue = attributes.get(ATTR_VALUE);
        myAuthor = attributes.get(ATTR_AUTHOR);
        myComment = attributes.get(ATTR_COMMENT);

        amIVisible = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_VISIBLE), beans, false);
    }

    /**
     * <p>Place the "value" attribute in the cell, and the rest of the
     * attributes control the creation of a cell comment.</p>
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
        int left = block.getLeftColNum();
        int top = block.getTopRowNum();
        // It should exist in this Cell; this Tag was found in it.
        Row row = sheet.getRow(top);
        Cell cell = row.getCell(left);
        WorkbookContext workbookContext = getWorkbookContext();
        SheetUtil.setCellValue(workbookContext, cell, myValue);

        Drawing drawing = context.getOrCreateDrawing();
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();

        Object commentString = AttributeUtil.evaluateRichTextStringNotNull(this, myComment, helper, beans, ATTR_COMMENT, "");
        String author = AttributeUtil.evaluateStringNotNull(this, myAuthor, beans, ATTR_AUTHOR, "");
        int commentLength;
        if (commentString instanceof RichTextString)
        {
            commentLength = ((RichTextString) commentString).length();
        }
        else
        {
            commentLength = commentString.toString().length();
        }

        // Calculate number of rows/cols to fit the comment text.  Try adding a
        // column, then adding a row, repeatedly, until we are sure that a box of
        // that size will hold the comment text.
        int rows = 0;
        int cols = 0;
        double width = 0;
        double height = 0;
        double fontHeightPoints = (cell instanceof HSSFCell) ?
                ((HSSFCell) cell).getCellStyle().getFont(sheet.getWorkbook()).getFontHeightInPoints() :
                ((XSSFCell) cell).getCellStyle().getFont().getFontHeightInPoints();
        while (width * height < commentLength)
        {
            cols++;
            width += sheet.getColumnWidth(left + cols) / 256;
            if (width * height >= commentLength)
                break;

            Row r = sheet.getRow(top + rows);
            if (r == null)
                r = sheet.createRow(top + rows);
            height += r.getHeightInPoints() / fontHeightPoints;
            rows++;
        }

        // HSSFComments seem to need an extra row's worth of height.
        if (cell instanceof HSSFCell)
            rows++;

        anchor.setCol1(left);
        anchor.setCol2(left + cols);
        anchor.setRow1(top);
        anchor.setRow2(top + rows);

        Comment comment = drawing.createCellComment(anchor);
        comment.setAuthor(author);
        if (commentString instanceof RichTextString)
        {
            comment.setString((RichTextString) commentString);
        }
        else
        {
            comment.setString(helper.createRichTextString(commentString.toString()));
        }
        comment.setVisible(amIVisible);
        cell.setCellComment(comment);

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, workbookContext);

        return true;
    }
}
