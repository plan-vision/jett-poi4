package net.sf.jett.model;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>CellStyleCache</code> is used internally to keep track of
 * <code>CellStyles</code>.  It defines and uses a string format for declaring
 * all possible style values that can be defined in a <code>CellStyle</code>,
 * even <code>Font</code> characteristics.  Its lifetime is meant to last only
 * during a single transformation.  When created, it reads in all pre-existing
 * <code>CellStyle</code> information and caches it for later reference.</p>
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class CellStyleCache
{
    private static final Logger logger = LoggerFactory.getLogger(CellStyleCache.class);

    private static final String PROP_SEP = "|";

    private Workbook myWorkbook;
    private Map<String, CellStyle> myCellStyleMap;

    /**
     * Constructs a <code>CellStyleCache</code> on a <code>Workbook</code>.
     * Caches all <code>CellStyles</code> found within.
     *
     * @param workbook A <code>Workbook</code>.
     */
    public CellStyleCache(Workbook workbook)
    {
        myWorkbook = workbook;
        myCellStyleMap = new HashMap<>();
        cachePreExistingCellStyles();
    }

    /**
     * Cache all <code>CellStyles</code> found within the workbook.
     */
    private void cachePreExistingCellStyles()
    {
        int numCellStyles = myWorkbook.getNumCellStyles();
        logger.trace("Caching {} pre-existing cell styles.", numCellStyles);
        for (int i = 0; i < numCellStyles; i++)
        {
            cacheCellStyle(myWorkbook.getCellStyleAt(i));
        }
        logger.trace("Done caching pre-existing styles.");
    }

    /**
     * Returns the number of entries in this cache.
     *
     * @return The number of entries in this cache.
     */
    public int getNumEntries()
    {
        return myCellStyleMap.size();
    }

    /**
     * Retrieve a <code>CellStyle</code> from the cache with the given
     * properties.
     *
     * @param fontBoldweight The font boldweight.
     * @param fontItalic Whether the font is italic.
     * @param fontColor The font color.
     * @param fontName The font name.
     * @param fontHeightInPoints The font height in points.
     * @param alignment The horizontal alignment.
     * @param borderBottom The bottom border type.
     * @param borderLeft The left border type.
     * @param borderRight The right border type.
     * @param borderTop The top border type.
     * @param dataFormat The data format string.
     * @param fontUnderline The font underline.
     * @param fontStrikeout Whether the font is in strikeout.
     * @param wrapText Whether text is wrapped.
     * @param fillBackgroundColor The fill background color.
     * @param fillForegroundColor The fill foreground color.
     * @param fillPattern The fill pattern.
     * @param verticalAlignment The vertical alignment.
     * @param indention How many characters the text is indented.
     * @param rotation How many degrees the text is rotated.
     * @param bottomBorderColor The bottom border color.
     * @param leftBorderColor The left border color.
     * @param rightBorderColor The right border color.
     * @param topBorderColor The top border color.
     * @param fontCharset The font charset.
     * @param fontTypeOffset The font type offset.
     * @param locked Whether the cell is "locked".
     * @param hidden Whether the cell is "hidden".
     * @return A <code>CellStyle</code> that matches all given properties, or
     * <code>null</code> if it doesn't exist.
     */
    public CellStyle retrieveCellStyle(boolean fontBoldweight, boolean fontItalic, Color fontColor, String fontName,
            short fontHeightInPoints, HorizontalAlignment alignment, BorderStyle borderBottom, BorderStyle borderLeft, BorderStyle borderRight,
            BorderStyle borderTop, String dataFormat, byte fontUnderline, boolean fontStrikeout, boolean wrapText,
            Color fillBackgroundColor, Color fillForegroundColor, FillPatternType fillPattern, VerticalAlignment verticalAlignment,
            short indention, short rotation, Color bottomBorderColor, Color leftBorderColor, Color rightBorderColor,
            Color topBorderColor, int fontCharset, short fontTypeOffset, boolean locked, boolean hidden)
    {
        String representation = getRepresentation(fontBoldweight, fontItalic, fontColor, fontName, fontHeightInPoints,
                alignment, borderBottom, borderLeft, borderRight, borderTop, dataFormat, fontUnderline, fontStrikeout,
                wrapText, fillBackgroundColor, fillForegroundColor, fillPattern, verticalAlignment, indention, rotation,
                bottomBorderColor, leftBorderColor, rightBorderColor, topBorderColor, fontCharset, fontTypeOffset, locked,
                hidden
        );
        CellStyle cs = myCellStyleMap.get(representation);
        if (logger.isTraceEnabled())
        {
            if (cs != null)
                logger.trace("CSCache hit  : {}", representation);
            else
                logger.trace("CSCache miss!: {}", representation);
        }
        return cs;
    }

    /**
     * Caches the given <code>CellStyle</code>.
     *
     * @param cs A <code>CellStyle</code>.
     */
    public void cacheCellStyle(CellStyle cs)
    {
        String representation = getRepresentation(cs);
        logger.trace("Caching cs   : {}", representation);
        myCellStyleMap.put(representation, cs);
    }

    /**
     * Finds the given <code>CellStyle</code>, but with the font characteristics
     * of the given <code>Font</code>, not its own <code>Font</code>.
     * @param cs The <code>CellStyle</code>.  Cell style characteristics are
     *    used, but the font characteristics are not used.
     * @param f The <code>Font</code>.  These font characteristics are used
     *    instead of the font characteristics on the <code>CellStyle</code>.
     * @return The <code>CellStyle</code> with the cell style characteristics of
     *    <code>cs</code> and the font characteristics of <code>f</code>, if
     *    found, else <code>null</code>.
     * @since 0.10.0
     */
    public CellStyle findCellStyleWithFont(CellStyle cs, Font f)
    {
        String representation = getRepresentation(cs, f);
        return myCellStyleMap.get(representation);
    }

    /**
     * Gets the string representation of the given <code>CellStyle</code>, using
     * its own font characteristics.
     * @param cs A <code>CellStyle</code>.
     * @return The string representation.
     */
    private String getRepresentation(CellStyle cs)
    {
        return getRepresentation(cs, myWorkbook.getFontAt(cs.getFontIndex()));
    }

    /**
     * Gets the string representation of the given <code>CellStyle</code>, using
     * the cell style characteristics of the <code>CellStyle</code> and the font
     * characteristics of the given <code>Font</code>.
     * @param cs The <code>CellStyle</code>.  Cell style characteristics are
     *    used, but the font characteristics are not used.
     * @param f The <code>Font</code>.  These font characteristics are used
     *    instead of the font characteristics on the <code>CellStyle</code>.
     * @return The string representation.
     * @since 0.10.0
     */
    private String getRepresentation(CellStyle cs, Font f)
    {
        // Colors that need an instanceof check
        Color fontColor;
        Color bottomColor = null;
        Color leftColor = null;
        Color rightColor = null;
        Color topColor = null;
        if (cs instanceof HSSFCellStyle)
        {
            HSSFFont hf = (HSSFFont) f;
            fontColor = hf.getHSSFColor((HSSFWorkbook) myWorkbook);
            // HSSF only stores border colors if the borders aren't "NONE".
            if (cs.getBorderBottom() != BorderStyle.NONE)
                bottomColor = ExcelColor.getHssfColorByIndex(cs.getBottomBorderColor());
            if (cs.getBorderLeft() != BorderStyle.NONE)
                leftColor = ExcelColor.getHssfColorByIndex(cs.getLeftBorderColor());
            if (cs.getBorderRight() != BorderStyle.NONE)
                rightColor = ExcelColor.getHssfColorByIndex(cs.getRightBorderColor());
            if (cs.getBorderTop() != BorderStyle.NONE)
                topColor = ExcelColor.getHssfColorByIndex(cs.getTopBorderColor());
        }
        else if (cs instanceof XSSFCellStyle)
        {
            XSSFFont xf = (XSSFFont) f;
            fontColor = xf.getXSSFColor();
            XSSFCellStyle xcs = (XSSFCellStyle) cs;
            bottomColor = xcs.getBottomBorderXSSFColor();
            leftColor = xcs.getLeftBorderXSSFColor();
            rightColor = xcs.getRightBorderXSSFColor();
            topColor = xcs.getTopBorderXSSFColor();
        }
        else
            throw new IllegalArgumentException("Bad CellStyle type: " + cs.getClass().getName());

        return getRepresentation(f.getBold(), f.getItalic(), fontColor, f.getFontName(),
                f.getFontHeightInPoints(), cs.getAlignment(), cs.getBorderBottom(), cs.getBorderLeft(), cs.getBorderRight(),
                cs.getBorderTop(), cs.getDataFormatString(), f.getUnderline(), f.getStrikeout(), cs.getWrapText(),
                cs.getFillBackgroundColorColor(), cs.getFillForegroundColorColor(), cs.getFillPattern(), cs.getVerticalAlignment() ,
                cs.getIndention(), cs.getRotation(), bottomColor, leftColor, rightColor,
                topColor, f.getCharSet(), f.getTypeOffset(), cs.getLocked(), cs.getHidden());
    }

 
	/**
     * Return the string representation of a <code>CellStyle</code> with the
     * given properties.
     * @param fontBoldweight The font boldweight.
     * @param fontItalic Whether the font is italic.
     * @param fontColor The font color.
     * @param fontName The font name.
     * @param fontHeightInPoints The font height in points.
     * @param alignment The horizontal alignment.
     * @param borderBottom The bottom border type.
     * @param borderLeft The left border type.
     * @param borderRight The right border type.
     * @param borderTop The top border type.
     * @param dataFormat The data format string.
     * @param fontUnderline The font underline.
     * @param fontStrikeout Whether the font is in strikeout.
     * @param wrapText Whether text is wrapped.
     * @param fillBackgroundColor The fill background color.
     * @param fillForegroundColor The fill foreground color.
     * @param fillPattern The fill pattern.
     * @param verticalAlignment The vertical alignment.
     * @param indention How many characters the text is indented.
     * @param rotation How many degrees the text is rotated.
     * @param bottomBorderColor The bottom border color.
     * @param leftBorderColor The left border color.
     * @param rightBorderColor The right border color.
     * @param topBorderColor The top border color.
     * @param fontCharset The font charset.
     * @param fontTypeOffset The font type offset.
     * @param locked Whether the cell is "locked".
     * @param hidden Whether the cell is "hidden".
     * @return The string representation.
     */
    private String getRepresentation(boolean fontBoldweight, boolean fontItalic, Color fontColor, String fontName,
    		short fontHeightInPoints, HorizontalAlignment alignment, BorderStyle borderBottom, BorderStyle borderLeft, BorderStyle borderRight,
    		BorderStyle borderTop, String dataFormat, byte fontUnderline, boolean fontStrikeout, boolean wrapText,
    		Color fillBackgroundColor, Color fillForegroundColor, FillPatternType fillPattern, VerticalAlignment verticalAlignment,
    		short indention, short rotation, Color bottomBorderColor, Color leftBorderColor, Color rightBorderColor,
    		Color topBorderColor, int fontCharset, short fontTypeOffset, boolean locked, boolean hidden)	
    {
        StringBuilder buf = new StringBuilder();

        buf.append(fontBoldweight).append(PROP_SEP);
        // Font italic
        buf.append(fontItalic).append(PROP_SEP);
        // Font color
        buf.append(SheetUtil.getColorHexString(fontColor));
        // Font name
        buf.append(PROP_SEP).append(fontName);
        // Font height in points
        buf.append(PROP_SEP).append(fontHeightInPoints);
        // Alignment
        buf.append(PROP_SEP).append(alignment);
        // Borders: Bottom, Left, Right, Top
        buf.append(PROP_SEP).append(borderBottom);
        buf.append(PROP_SEP).append(borderLeft);
        buf.append(PROP_SEP).append(borderRight);
        buf.append(PROP_SEP).append(borderTop);
        // Data format
        buf.append(PROP_SEP).append(dataFormat);
        // Font underline
        buf.append(PROP_SEP).append(fontUnderline);
        // Font strikeout
        buf.append(PROP_SEP).append(fontStrikeout);
        // Wrap text
        buf.append(PROP_SEP).append(wrapText);
        // Fill bg/fg color
        buf.append(PROP_SEP).append(SheetUtil.getColorHexString(fillBackgroundColor));
        buf.append(PROP_SEP).append(SheetUtil.getColorHexString(fillForegroundColor));
        // Fill pattern
        buf.append(PROP_SEP).append(fillPattern);
        // Vertical alignment
        buf.append(PROP_SEP).append(verticalAlignment);
        // Indention
        buf.append(PROP_SEP).append(indention);
        // Rotation
        buf.append(PROP_SEP).append(rotation);
        // DO NOT DO Column width in chars
        // DO NOT DO row height in points
        // Border Colors: Bottom, Left, Right, Top
        buf.append(PROP_SEP).append(SheetUtil.getColorHexString(bottomBorderColor));
        buf.append(PROP_SEP).append(SheetUtil.getColorHexString(leftBorderColor));
        buf.append(PROP_SEP).append(SheetUtil.getColorHexString(rightBorderColor));
        buf.append(PROP_SEP).append(SheetUtil.getColorHexString(topBorderColor));
        // Font charset
        buf.append(PROP_SEP).append(fontCharset);
        // Font type offset
        buf.append(PROP_SEP).append(fontTypeOffset);
        // Locked
        buf.append(PROP_SEP).append(locked);
        // Hidden
        buf.append(PROP_SEP).append(hidden);

        return buf.toString();
    }
}
