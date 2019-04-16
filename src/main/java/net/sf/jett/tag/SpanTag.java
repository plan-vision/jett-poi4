package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.CellStyleCache;
import net.sf.jett.model.ExcelColor;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>SpanTag</code> represents a cell or merged region that will span
 * extra rows and/or extra columns, depending on growth and/or adjustment
 * factors.  If this tag is applied to a cell that is not part of a merged
 * region, then it may result in the creation of a merged region.  If this tag
 * is applied to a cell that is part of a merged region, then it may result in
 * the removal of the merged region.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>factor (optional): <code>int</code></li>
 * <li>adjust (optional): <code>int</code></li>
 * <li>value (required): <code>RichTextString</code></li>
 * <li>expandRight (optional): <code>boolean</code></li>
 * <li>fixed (optional): <code>boolean</code></li>
 * </ul>
 *
 * <p>Either one or both of the <code>factor</code> and the <code>adjust</code>
 * attributes must be specified.</p>
 *
 * @author Randy Gettman
 */
public class SpanTag extends BaseTag
{
    private static final Logger logger = LoggerFactory.getLogger(SpanTag.class);

    /**
     * Attribute for specifying the growth factor.
     */
    public static final String ATTR_FACTOR = "factor";
    /**
     * Attribute for specifying an adjustment to the size of the merged region.
     * @since 0.4.0
     */
    public static final String ATTR_ADJUST = "adjust";
    /**
     * Attribute for forcing "expand right" behavior.  (Default is expand down.)
     */
    public static final String ATTR_EXPAND_RIGHT = "expandRight";
    /**
     * Attribute that specifies the value of the cell/merged region.
     */
    public static final String ATTR_VALUE = "value";
    /**
     * Attribute that specifies the value of the cell/merged region.
     * @since 0.9.1
     */
    public static final String ATTR_FIXED = "fixed";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_VALUE));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_EXPAND_RIGHT, ATTR_FACTOR, ATTR_ADJUST, ATTR_FIXED));

    private int myFactor = 1;
    private int myAdjust = 0;
    private RichTextString myValue;
    private boolean amIFixed;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "span";
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

        if (!isBodiless())
            throw new TagParseException("SpanTag: Must be bodiless.  SpanTag with body found" + getLocation());

        myValue = attributes.get(ATTR_VALUE);

        List<RichTextString> atLeastOne = Arrays.asList(attributes.get(ATTR_FACTOR), attributes.get(ATTR_ADJUST));
        AttributeUtil.ensureAtLeastOneExists(this, atLeastOne, Arrays.asList(ATTR_FACTOR, ATTR_ADJUST));
        myFactor = AttributeUtil.evaluateNonNegativeInt(this, attributes.get(ATTR_FACTOR), beans, ATTR_FACTOR, 1);
        myAdjust = AttributeUtil.evaluateInt(this, attributes.get(ATTR_ADJUST), beans, ATTR_ADJUST, 0);

        boolean explicitlyExpandingRight = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_EXPAND_RIGHT), beans, false);
        if (explicitlyExpandingRight)
            block.setDirection(Block.Direction.HORIZONTAL);
        else
            block.setDirection(Block.Direction.VERTICAL);

        amIFixed = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_FIXED), beans, false);
    }

    /**
     * <p>If not already part of a merged region, and one of the factors is
     * greater than 1, then create a merged region.  Else, replace the current
     * merged region with a new merged region.</p>
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
//      long start = System.nanoTime();
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();

        logger.debug("SpanTag.process: factor={}, block direction is {}", myFactor, block.getDirection());

        int left = block.getLeftColNum();
        int right = left;
        int top = block.getTopRowNum();
        int bottom = top;
        // Assume a "merged region" of 1 X 1 for now.
        int height = 1;
        int width = 1;

        List<CellRangeAddress> sheetMergedRegions = context.getMergedRegions();
        int index = findMergedRegionAtCell(sheetMergedRegions, left, top);
        if (index != -1)
        {
            // Get the height/width and remove the old merged region.
            CellRangeAddress remove = sheetMergedRegions.get(index);
            right = remove.getLastColumn();
            bottom = remove.getLastRow();
            height = remove.getLastRow() - remove.getFirstRow() + 1;
            width = remove.getLastColumn() - remove.getFirstColumn() + 1;
            logger.debug("  Removing region: {}, height={}, width={}", remove, height, width);
            sheetMergedRegions.remove(index);
        }

        BorderStyle borderBottomType = BorderStyle.NONE;
        BorderStyle borderLeftType = BorderStyle.NONE;
        BorderStyle borderRightType = BorderStyle.NONE;
        BorderStyle borderTopType = BorderStyle.NONE;
        Color borderBottomColor = null;
        Color borderLeftColor = null;
        Color borderRightColor = null;
        Color borderTopColor = null;
        // Get current borders and border colors.
        Row rTop = sheet.getRow(top);
        if (rTop != null)
        {
            Cell cLeft = rTop.getCell(left);
            if (cLeft != null)
            {
                CellStyle cs = cLeft.getCellStyle();
                borderLeftType = cs.getBorderLeft();
                borderTopType = cs.getBorderTop();
                // Border colors need instanceof check.
                if (cs instanceof HSSFCellStyle)
                {
                    borderLeftColor = ExcelColor.getHssfColorByIndex(cs.getLeftBorderColor());
                    borderTopColor = ExcelColor.getHssfColorByIndex(cs.getTopBorderColor());
                }
                else
                {
                    // XSSFCellStyle
                    XSSFCellStyle xcs = (XSSFCellStyle) cs;
                    borderLeftColor = xcs.getLeftBorderXSSFColor();
                    borderTopColor = xcs.getTopBorderXSSFColor();
                }
            }
        }
        Row rBottom = sheet.getRow(bottom);
        if (rBottom != null)
        {
            Cell cRight = rBottom.getCell(right);
            if (cRight != null)
            {
                CellStyle cs = cRight.getCellStyle();
                borderRightType = cs.getBorderRight();
                borderBottomType = cs.getBorderBottom();
                // Border colors need instanceof check.
                if (cs instanceof HSSFCellStyle)
                {
                    borderRightColor = ExcelColor.getHssfColorByIndex(cs.getRightBorderColor());
                    borderBottomColor = ExcelColor.getHssfColorByIndex(cs.getBottomBorderColor());
                }
                else
                {
                    // XSSFCellStyle
                    XSSFCellStyle xcs = (XSSFCellStyle) cs;
                    borderRightColor = xcs.getRightBorderXSSFColor();
                    borderBottomColor = xcs.getBottomBorderXSSFColor();
                }
            }
        }
        if (borderTopType != BorderStyle.NONE || borderBottomType != BorderStyle.NONE ||
                borderRightType != BorderStyle.NONE || borderLeftType != BorderStyle.NONE)
        {
            removeBorders(sheet, left, right, top, bottom);
        }

        // The block for which to shift content out of the way or to remove is
        // actually the old merged region.
        Block mergedBlock = new Block(block.getParent(), left, right, top, bottom);
        mergedBlock.setDirection(block.getDirection());

        // Determine new height or width, plus new bottom or right.
        int change;
        if (block.getDirection() == Block.Direction.VERTICAL)
        {
            change = height * (myFactor - 1) + myAdjust;
            bottom += change;
            height = bottom - top + 1;
        }
        else
        {
            change = width * (myFactor - 1) + myAdjust;
            right += change;
            width = right - left + 1;
        }

        // Remove.
        if (height <= 0 || width <= 0)
        {
            logger.debug("  Calling removeBlock on block: {}", mergedBlock);
            SheetUtil.removeBlock(sheet, context, mergedBlock, getWorkbookContext());
            return false;
        }
        // Shrink.
        if (change < 0)
        {
            Block remove;
            if (block.getDirection() == Block.Direction.VERTICAL)
                remove = new Block(block.getParent(), left, right, bottom + 1, bottom - change);
            else
                remove = new Block(block.getParent(), right + 1, right - change, top, bottom);
            remove.setDirection(block.getDirection());
            logger.debug("  Calling removeBlock on fabricated block: {} (change {})", remove, change);
            if (amIFixed)
            {
                SheetUtil.clearBlock(sheet, remove, getWorkbookContext());
            }
            else
            {
                SheetUtil.removeBlock(sheet, context, remove, getWorkbookContext());
            }
        }
        // Expand.
        if (change > 0 && !amIFixed)
        {
            Block expand;
            if (block.getDirection() == Block.Direction.VERTICAL)
                expand = new Block(block.getParent(), left, right, bottom - change, bottom - change);
            else
                expand = new Block(block.getParent(), right - change, right - change, top, bottom);
            expand.setDirection(block.getDirection());
            logger.debug("  Calling shiftForBlock on fabricated block: {} with change {}", expand, change + 1);
            SheetUtil.shiftForBlock(sheet, context, expand, getWorkbookContext(), change + 1);
        }
        if (borderTopType != BorderStyle.NONE || borderBottomType != BorderStyle.NONE ||
                borderRightType != BorderStyle.NONE || borderLeftType != BorderStyle.NONE)
        {
            putBackBorders(sheet, left, right, top, bottom,
                    borderLeftType, borderRightType, borderTopType, borderBottomType,
                    borderLeftColor, borderRightColor, borderTopColor, borderBottomColor);
        }

        // Set the value.
        Row row = sheet.getRow(top);
        Cell cell = row.getCell(left);
        WorkbookContext workbookContext = getWorkbookContext();
        SheetUtil.setCellValue(workbookContext, cell, myValue);

        // Create the replacement merged region, or the new merged region if it
        // didn't exist before.
        if (height > 1 || width > 1)
        {
            CellRangeAddress create = new CellRangeAddress(top, bottom, left, right);
            logger.debug("  Adding region: {}", create);
            sheetMergedRegions.add(create);
        }

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, workbookContext);
//
//      long end = System.nanoTime();
//      System.out.println("findMergedRegionAtCell: " + (end - start) + " ns");

        return true;
    }

    /**
     * Identify the merged region on the given <code>Sheet</code> whose top-left
     * corner is at the specified column and row indexes.
     * @param sheetMergedRegions A <code>List</code> of
     *    <code>CellRangeAddress</code>es.
     * @param col The 0-based column index of the top-left corner.
     * @param row The 0-based row index of the top-left corner.
     * @return A 0-based index into the <code>Sheet's</code> list of merged
     *    regions, or -1 if not found.
     */
    private int findMergedRegionAtCell(List<CellRangeAddress> sheetMergedRegions, int col, int row)
    {
        int numMergedRegions = sheetMergedRegions.size();
        for (int i = 0; i < numMergedRegions; i++)
        {
            CellRangeAddress candidate = sheetMergedRegions.get(i);
            if (candidate.getFirstRow() == row && candidate.getFirstColumn() == col)
                return i;
        }
        return -1;
    }

    /**
     * Remove all borders from all cells in the region described by the left,
     * right, top, and bottom bounds.
     * @param sheet The <code>Sheet</code>.
     * @param left The 0-based index indicating the left-most part of the region.
     * @param right The 0-based index indicating the right-most part of the region.
     * @param top The 0-based index indicating the top-most part of the region.
     * @param bottom The 0-based index indicating the bottom-most part of the region.
     */
    private void removeBorders(Sheet sheet, int left, int right, int top, int bottom)
    {
        logger.debug("removeBorders: {}, {}, {}, {}", left, right, top, bottom);
        CellStyleCache csCache = getWorkbookContext().getCellStyleCache();
        for (int r = top; r <= bottom; r++)
        {
            Row row = sheet.getRow(r);
            for (int c = left; c <= right; c++)
            {
                Cell cell = row.getCell(c);
                if (cell != null)
                {
                    CellStyle cs = cell.getCellStyle();
                    Font f = sheet.getWorkbook().getFontAt(cs.getFontIndex());
                    Color fontColor;
                    if (cs instanceof HSSFCellStyle)
                    {
                        fontColor = ExcelColor.getHssfColorByIndex(f.getColor());
                    }
                    else
                    {
                        fontColor = ((XSSFFont) f).getXSSFColor();
                    }
                    // At this point, we have all of the desired CellStyle and Font
                    // characteristics.  Find a CellStyle if it exists.
                    CellStyle foundStyle = csCache.retrieveCellStyle(f.getBold(), f.getItalic(), fontColor,
                            f.getFontName(), f.getFontHeightInPoints(), cs.getAlignment(), BorderStyle.NONE,
                            BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, cs.getDataFormatString(),
                            f.getUnderline(), f.getStrikeout(), cs.getWrapText(), cs.getFillBackgroundColorColor(),
                            cs.getFillForegroundColorColor(), cs.getFillPattern(), cs.getVerticalAlignment(), cs.getIndention(),
                            cs.getRotation(), null, null, null, null,
                            f.getCharSet(), f.getTypeOffset(), cs.getLocked(), cs.getHidden());

                    if (foundStyle == null)
                    {
                        foundStyle = SheetUtil.createCellStyle(sheet.getWorkbook(), cs.getAlignment(), BorderStyle.NONE,
                        		BorderStyle.NONE, BorderStyle.NONE, BorderStyle.NONE, cs.getDataFormatString(),
                                cs.getWrapText(), cs.getFillBackgroundColorColor(), cs.getFillForegroundColorColor(),
                                cs.getFillPattern(), cs.getVerticalAlignment(), cs.getIndention(), cs.getRotation(),
                                null, null, null, null, cs.getLocked(), cs.getHidden());
                        foundStyle.setFont(f);
                        csCache.cacheCellStyle(foundStyle);
                    }
                    cell.setCellStyle(foundStyle);
                }
            }
        }
    }

    /**
     * Puts back borders for the newly sized merged region.
     * @param sheet The <code>Sheet</code>.
     * @param left The 0-based index indicating the left-most part of the region.
     * @param right The 0-based index indicating the right-most part of the region.
     * @param top The 0-based index indicating the top-most part of the region.
     * @param bottom The 0-based index indicating the bottom-most part of the region.
     * @param borderLeft The left border type.
     * @param borderRight The right border type.
     * @param borderTop The top border type.
     * @param borderBottom The bottom border type.
     * @param borderLeftColor The left border color.
     * @param borderRightColor The right border color.
     * @param borderTopColor The top border color.
     * @param borderBottomColor The bottom border color.
     */
    private void putBackBorders(Sheet sheet, int left, int right, int top, int bottom,
    							BorderStyle borderLeft, BorderStyle  borderRight, BorderStyle  borderTop, BorderStyle  borderBottom,
                                Color borderLeftColor, Color borderRightColor, Color borderTopColor, Color borderBottomColor)
    {
        logger.debug("putBackBorders: {}, {}, {}, {}", left, right, top, bottom);
        CellStyleCache csCache = getWorkbookContext().getCellStyleCache();
        for (int r = top; r <= bottom; r++)
        {
            Row row = sheet.getRow(r);
            if (row == null)
                row = sheet.createRow(r);
            for (int c = left; c <= right; c++)
            {
                Cell cell = row.getCell(c);
                if (cell == null)
                    cell = row.createCell(c);

                CellStyle cs = cell.getCellStyle();
                Font f = sheet.getWorkbook().getFontAt(cs.getFontIndex());
                Color fontColor;
                if (cs instanceof HSSFCellStyle)
                {
                    fontColor = ExcelColor.getHssfColorByIndex(f.getColor());
                }
                else
                {
                    fontColor = ((XSSFFont) f).getXSSFColor();
                }
                BorderStyle newBorderBottom = (r == bottom) ? borderBottom : BorderStyle.NONE;
                BorderStyle newBorderLeft = (c == left) ? borderLeft : BorderStyle.NONE;
                BorderStyle newBorderRight = (c == right) ? borderRight : BorderStyle.NONE;
                BorderStyle newBorderTop = (r == top) ? borderTop : BorderStyle.NONE;
                Color newBorderBottomColor = (r == bottom) ? borderBottomColor : null;
                Color newBorderLeftColor = (c == left) ? borderLeftColor : null;
                Color newBorderRightColor = (c == right) ? borderRightColor : null;
                Color newBorderTopColor = (r == top) ? borderTopColor : null;
                // At this point, we have all of the desired CellStyle and Font
                // characteristics.  Find a CellStyle if it exists.
                CellStyle foundStyle = csCache.retrieveCellStyle(f.getBold(), f.getItalic(), fontColor,
                        f.getFontName(), f.getFontHeightInPoints(), cs.getAlignment(),
                        newBorderBottom, newBorderLeft, newBorderRight, newBorderTop, cs.getDataFormatString(),
                        f.getUnderline(), f.getStrikeout(), cs.getWrapText(), cs.getFillBackgroundColorColor(),
                        cs.getFillForegroundColorColor(), cs.getFillPattern(), cs.getVerticalAlignment(), cs.getIndention(),
                        cs.getRotation(), newBorderBottomColor, newBorderLeftColor, newBorderRightColor, newBorderTopColor,
                        f.getCharSet(), f.getTypeOffset(), cs.getLocked(), cs.getHidden());

                if (foundStyle == null)
                {
                    foundStyle = SheetUtil.createCellStyle(sheet.getWorkbook(), cs.getAlignment(), newBorderBottom,
                            newBorderLeft, newBorderRight, newBorderTop, cs.getDataFormatString(),
                            cs.getWrapText(), cs.getFillBackgroundColorColor(), cs.getFillForegroundColorColor(),
                            cs.getFillPattern(), cs.getVerticalAlignment(), cs.getIndention(), cs.getRotation(),
                            newBorderBottomColor, newBorderLeftColor, newBorderRightColor, newBorderTopColor,
                            cs.getLocked(), cs.getHidden());
                    foundStyle.setFont(f);
                    csCache.cacheCellStyle(foundStyle);
                }
                cell.setCellStyle(foundStyle);
            }
        }
    }
}
