package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.CellStyleCache;
import net.sf.jett.model.ExcelColor;
import net.sf.jett.model.FontCache;
import net.sf.jett.model.Style;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.parser.StyleParser;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>StyleTag</code> represents a dynamically determined style for a
 * <code>Cell</code>.  A <code>StyleTag</code> must have a body.</p>
 * <p>The <code>style</code> attribute works like the HTML "style" attribute,
 * in that one can specify one or more style elements in a
 * <code>property: value; property: value</code> style.  If a property is
 * specified, then it will override whatever value is already present in the
 * <code>Cell</code>.  If a property value is an empty string or the property
 * is not present, then it will be ignored and it will not override whatever
 * value is already present in the <code>Cell</code>.  Unrecognized property
 * names and unrecognized values for a property are ignored and do not override
 * whatever value is already present in the <code>Cell</code>.  Property names
 * and values may be specified in a case insensitive-fashion, i.e. "CENTER" =
 * "Center" = "center"..</p>
 * <p>The <code>class</code> attributes works like the HTML "class" attribute,
 * in that one can specify one or more CSS-like style "classes" in a
 * semicolon-delimited list.  Register CSS-like files and/or CSS-like text with
 * the <code>ExcelTransformer</code> prior to transformation.  Subsequent class
 * names override previous class names, and the <code>style</code> attribute
 * overrides the <code>class</code> attribute.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>style (optional): <code>String</code></li>
 * <li>class (optional): <code>String</code></li>
 * </ul>
 *
 * <p>For supported property names and values (plus explanations), please see
 * {@link net.sf.jett.parser.TagParser}.</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 * @see net.sf.jett.transform.ExcelTransformer
 */
public class StyleTag extends BaseTag
{
    private static final Logger logger = LoggerFactory.getLogger(StyleTag.class);

    /**
     * Attribute that specifies the desired style property(ies) to change in the
     * current <code>Cell</code>.  Properties are specified in a string with the
     * following format: <code>property1: value1; property2: value2; ...</code>
     */
    public static final String ATTR_STYLE = "style";
    /**
     * Attribute that specifies the desired pre-defined style class to apply to
     * the current <code>Cell</code>.  Pre-defined styles are defined by
     * registering styles with the <code>ExcelTransformer</code> prior to
     * transformation.
     * @since 0.5.0
     * @see net.sf.jett.transform.ExcelTransformer#addCssFile(String)
     * @see net.sf.jett.transform.ExcelTransformer#addCssText(String)
     */
    public static final String ATTR_CLASS = "class";

    // Used so that a user can escape the separator.
    // This matches SPEC_SEP but not the concatenation "\" + SPEC_SEP.
    private static final String SPLIT_SPEC = "(?<!\\\\)" + SPEC_SEP;

    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_CLASS, ATTR_STYLE));

    private Style myStyle;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "style";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        return super.getRequiredAttributes();
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
     * Validates the attributes for this <code>Tag</code>.  This tag must have a
     * body.
     */
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (isBodiless())
            throw new TagParseException("Style tags must have a body.  Bodiless style tag found" + getLocation());

        TagContext context = getContext();
        WorkbookContext wc = getWorkbookContext();
        Map<String, Style> styleMap = wc.getStyleMap();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        myStyle = new Style();

        List<String> styleClasses = AttributeUtil.evaluateList(this, attributes.get(ATTR_CLASS), beans, null);
        if (styleClasses != null)
        {
            for (String styleClass : styleClasses)
            {
                Style style = styleMap.get(styleClass.trim());
                if (style != null)
                    myStyle.apply(style);
            }
        }

        String line = AttributeUtil.evaluateString(this, attributes.get(ATTR_STYLE), beans, null);
        if (line != null)
        {
            String[] styles = line.split(SPLIT_SPEC);
            for (String strStyle : styles)
            {
                String property;
                String value;
                // Replace escaped separators with the normal character for further
                // processing.
                String[] parts = strStyle.replace("\\" + SPEC_SEP, SPEC_SEP).split(":", 2);
                if (parts.length < 2)
                {
                    continue;
                }
                property = parts[0].trim();
                value = parts[1].trim();

                if (value.length() >= 1)
                {
                    StyleParser.addStyle(myStyle, property, value);
                }
            }
        }
    }

    /**
     * <p>Override the cells' current styles with any non-null style property
     * values.</p>
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
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();

        if (myStyle.isStyleToApply())
        {
            // Loop through Rows and Cells, and apply the style to each one in
            // turn.
            for (int r = top; r <= bottom; r++)
            {
                Row row = sheet.getRow(r);
                if (row != null)
                {
                    for (int c = left; c <= right; c++)
                    {
                        Cell cell = row.getCell(c);
                        if (cell != null)
                        {
                            examineAndApplyStyle(workbook, cell);
                        }
                    }
                }
            }
        }

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, getWorkbookContext());

        return true;
    }

    /**
     * Examine the given <code>Cell's</code> current <code>CellStyle</code>.  If
     * necessary, replace its <code>CellStyle</code> and/or <code>Font</code,
     * guided by the property values retrieved earlier from the "style"
     * attribute.
     * @param workbook The <code>Workbook</code> that maintains all
     *    <code>CellStyles</code> and <code>Fonts</code>.
     * @param cell The <code>Cell</code> to examine.
     */
    private void examineAndApplyStyle(Workbook workbook, Cell cell)
    {
        WorkbookContext wc = getWorkbookContext();
        CellStyleCache csCache = wc.getCellStyleCache();
        FontCache fCache = wc.getFontCache();

        CellStyle cs = cell.getCellStyle();
        Font f = workbook.getFontAt(cs.getFontIndex());

        logger.debug("eAAS: cell at ({}, {})", cell.getRowIndex(), cell.getColumnIndex());

        HorizontalAlignment alignment = (myStyle.getAlignment() != null) ? myStyle.getAlignment() : cs.getAlignment();
        BorderStyle borderBottom = (myStyle.getBorderBottomType() != null) ? myStyle.getBorderBottomType() : cs.getBorderBottom();        
        BorderStyle borderLeft = (myStyle.getBorderLeftType() != null) ? myStyle.getBorderLeftType() : cs.getBorderLeft();
        BorderStyle borderRight = (myStyle.getBorderRightType() != null) ? myStyle.getBorderRightType() : cs.getBorderRight();
        BorderStyle borderTop = (myStyle.getBorderTopType() != null) ? myStyle.getBorderTopType() : cs.getBorderTop();
        String dataFormat = (myStyle.getDataFormat() != null) ? myStyle.getDataFormat() : cs.getDataFormatString();
        Color fillBackgroundColor = (myStyle.getFillBackgroundColor() != null) ?
                SheetUtil.getColor(workbook, myStyle.getFillBackgroundColor()) : cs.getFillBackgroundColorColor();
        Color fillForegroundColor = (myStyle.getFillForegroundColor() != null) ?
                SheetUtil.getColor(workbook, myStyle.getFillForegroundColor()): cs.getFillForegroundColorColor();
        FillPatternType fillPattern = (myStyle.getFillPatternType() != null) ? myStyle.getFillPatternType() : cs.getFillPattern();
        boolean hidden = (myStyle.isHidden() != null) ? myStyle.isHidden() : cs.getHidden();
        short indention = (myStyle.getIndention() != null) ? myStyle.getIndention() : cs.getIndention();
        boolean locked = (myStyle.isLocked() != null) ? myStyle.isLocked() : cs.getLocked();
        VerticalAlignment verticalAlignment = (myStyle.getVerticalAlignment() != null) ? myStyle.getVerticalAlignment() : cs.getVerticalAlignment();
        boolean wrapText = (myStyle.isWrappingText() != null) ? myStyle.isWrappingText() : cs.getWrapText();
        boolean fontBoldweight = (myStyle.getFontBoldweight() != null) ? myStyle.getFontBoldweight() : f.getBold();
        int fontCharset = (myStyle.getFontCharset() != null) ? myStyle.getFontCharset().getIndex() : f.getCharSet();
        short fontHeightInPoints = (myStyle.getFontHeightInPoints() != null) ? myStyle.getFontHeightInPoints() : f.getFontHeightInPoints();
        String fontName = (myStyle.getFontName() != null) ? myStyle.getFontName() : f.getFontName();
        boolean fontItalic = (myStyle.isFontItalic() != null) ? myStyle.isFontItalic() : f.getItalic();
        boolean fontStrikeout = (myStyle.isFontStrikeout() != null) ? myStyle.isFontStrikeout() : f.getStrikeout();
        short fontTypeOffset = (myStyle.getFontTypeOffset() != null) ? myStyle.getFontTypeOffset().getIndex() : f.getTypeOffset();
        byte fontUnderline = (myStyle.getFontUnderline() != null) ? myStyle.getFontUnderline().getIndex() : f.getUnderline();
        // Certain properties need a type of workbook check.
        Color bottomBorderColor = null;
        Color leftBorderColor = null;
        Color rightBorderColor = null;
        Color topBorderColor = null;
        Color fontColor;
        short rotationDegrees;
        if (workbook instanceof HSSFWorkbook)
        {
            short hssfBottomBorderColor = (myStyle.getBorderBottomColor() != null) ?
                    ((HSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderBottomColor())).getIndex() : cs.getBottomBorderColor();
            short hssfLeftBorderColor = (myStyle.getBorderLeftColor() != null) ?
                    ((HSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderLeftColor())).getIndex() : cs.getLeftBorderColor();
            short hssfRightBorderColor = (myStyle.getBorderRightColor() != null) ?
                    ((HSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderRightColor())).getIndex() : cs.getRightBorderColor();
            short hssfTopBorderColor = (myStyle.getBorderTopColor() != null) ?
                    ((HSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderTopColor())).getIndex() : cs.getTopBorderColor();
            short hssfFontColor = (myStyle.getFontColor() != null) ?
                    ((HSSFColor) SheetUtil.getColor(workbook, myStyle.getFontColor())).getIndex() : f.getColor();
            if (hssfBottomBorderColor != 0)
                bottomBorderColor = ExcelColor.getHssfColorByIndex(hssfBottomBorderColor);
            if (hssfLeftBorderColor != 0)
                leftBorderColor = ExcelColor.getHssfColorByIndex(hssfLeftBorderColor);
            if (hssfRightBorderColor != 0)
                rightBorderColor = ExcelColor.getHssfColorByIndex(hssfRightBorderColor);
            if (hssfTopBorderColor != 0)
                topBorderColor = ExcelColor.getHssfColorByIndex(hssfTopBorderColor);
            fontColor = ExcelColor.getHssfColorByIndex(hssfFontColor);

            rotationDegrees = (myStyle.getRotationDegrees() != null) ? myStyle.getRotationDegrees() : cs.getRotation();
        }
        else
        {
            // XSSFWorkbook
            XSSFCellStyle xcs = (XSSFCellStyle) cs;
            bottomBorderColor = (myStyle.getBorderBottomColor() != null) ?
                    ((XSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderBottomColor())) : xcs.getBottomBorderXSSFColor();
            leftBorderColor = (myStyle.getBorderLeftColor() != null) ?
                    ((XSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderLeftColor())) : xcs.getLeftBorderXSSFColor();
            rightBorderColor = (myStyle.getBorderRightColor() != null) ?
                    ((XSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderRightColor())) : xcs.getRightBorderXSSFColor();
            topBorderColor = (myStyle.getBorderTopColor() != null) ?
                    ((XSSFColor) SheetUtil.getColor(workbook, myStyle.getBorderTopColor())) : xcs.getTopBorderXSSFColor();
            fontColor = (myStyle.getFontColor() != null) ?
                    ((XSSFColor) SheetUtil.getColor(workbook, myStyle.getFontColor())) : ((XSSFFont) f).getXSSFColor();

            // XSSF: Negative rotation values don't make as much sense as in HSSF.
            // From 0-90, they coincide.
            // But HSSF -1  => XSSF 91 , HSSF -15 => XSSF 105,
            //     HSSF -90 => XSSF 180.
            rotationDegrees = (myStyle.getRotationDegrees() != null) ? myStyle.getRotationDegrees() : cs.getRotation();
            if (rotationDegrees < 0)
            {
                rotationDegrees = (short) (90 - rotationDegrees);
            }
        }

        // Process row height/column width separately.
        if (myStyle.getRowHeight() != null)
        {
            cell.getRow().setHeight(myStyle.getRowHeight());
        }
        if (myStyle.getColumnWidth() != null)
        {
            cell.getSheet().setColumnWidth(cell.getColumnIndex(), myStyle.getColumnWidth());
        }

        // At this point, we have all of the desired CellStyle and Font
        // characteristics.  Find a CellStyle if it exists.
        CellStyle foundStyle = csCache.retrieveCellStyle(fontBoldweight, fontItalic, fontColor, fontName,
                fontHeightInPoints, alignment, borderBottom, borderLeft, borderRight, borderTop, dataFormat, fontUnderline,
                fontStrikeout, wrapText, fillBackgroundColor, fillForegroundColor, fillPattern, verticalAlignment, indention,
                rotationDegrees, bottomBorderColor, leftBorderColor, rightBorderColor, topBorderColor, fontCharset,
                fontTypeOffset, locked, hidden);

        // Find the Font if not already found.
        if (foundStyle == null)
        {
            //short numFonts = workbook.getNumberOfFonts();
            //long start = System.nanoTime();
            Font foundFont = fCache.retrieveFont(fontBoldweight, fontItalic, fontColor, fontName,
                    fontHeightInPoints, fontUnderline, fontStrikeout, fontCharset, fontTypeOffset);
            //long end = System.nanoTime();
            //System.err.println("Find Font: " + (end - start) + " ns");

            // If Font still not found, then create it.
            if (foundFont == null)
            {
                //start = System.nanoTime();
                foundFont = SheetUtil.createFont(workbook, fontBoldweight, fontItalic, fontColor, fontName,
                        fontHeightInPoints, fontUnderline, fontStrikeout, fontCharset, fontTypeOffset);
                //end = System.nanoTime();
                //System.err.println("Create Font: " + (end - start) + " ns");
                fCache.cacheFont(foundFont);
                logger.trace("  Font created.");
            }

            // Create the new CellStyle.
            //start = System.nanoTime();
            foundStyle = SheetUtil.createCellStyle(workbook, alignment, borderBottom, borderLeft,
                    borderRight, borderTop, dataFormat, wrapText, fillBackgroundColor, fillForegroundColor,
                    fillPattern, verticalAlignment, indention, rotationDegrees, bottomBorderColor,
                    leftBorderColor, rightBorderColor, topBorderColor, locked, hidden);
            foundStyle.setFont(foundFont);
            //end = System.nanoTime();
            //System.err.println("Create CS: " + (end - start) + " ns");

            csCache.cacheCellStyle(foundStyle);
            logger.trace("  Created new style.");
        }

        cell.setCellStyle(foundStyle);
    }
}