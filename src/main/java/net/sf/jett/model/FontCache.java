package net.sf.jett.model;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>FontCache</code> is used internally to keep track of
 * <code>Fonts</code>.  It defines and uses a string format for declaring
 * all possible font values that can be defined in a <code>Font</code>.  Its
 * lifetime is meant to last only during a single transformation.  When
 * created, it reads in all pre-existing <code>Font</code> information and
 * caches it for later reference.</p>
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class FontCache
{
    private static final Logger logger = LoggerFactory.getLogger(FontCache.class);

    private static final String PROP_SEP = "|";

    private Workbook myWorkbook;
    private Map<String, Font> myFontMap;

    /**
     * Constructs a <code>FontCache</code> on a <code>Workbook</code>.
     * Caches all <code>Fonts</code> found within.
     * @param workbook A <code>Workbook</code>.
     */
    public FontCache(Workbook workbook)
    {
        myWorkbook = workbook;
        myFontMap = new HashMap<>();
        cachePreExistingFonts();
    }

    /**
     * Cache all <code>Fonts</code> found within the workbook.
     */
    private void cachePreExistingFonts()
    {
        short numFonts = myWorkbook.getNumberOfFonts();
        logger.trace("Caching {} pre-existing cell fonts.", numFonts);
        for (short i = 0; i < numFonts; i++)
        {
            cacheFont(myWorkbook.getFontAt(i));
        }
        logger.trace("Done caching pre-existing fonts.");
    }

    /**
     * Returns the number of entries in this cache.
     * @return The number of entries in this cache.
     */
    public int getNumEntries()
    {
        return myFontMap.size();
    }

    /**
     * Retrieve a <code>Font</code> from the cache with the given
     * properties.
     * @param fontBoldweight The font boldweight.
     * @param fontItalic Whether the font is italic.
     * @param fontColor The font color.
     * @param fontName The font name.
     * @param fontHeightInPoints The font height in points.
     * @param fontUnderline The font underline.
     * @param fontStrikeout Whether the font is in strikeout.
     * @param fontCharset The font charset.
     * @param fontTypeOffset The font type offset.
     * @return A <code>Font</code> that matches all given properties, or
     *    <code>null</code> if it doesn't exist.
     */
    public Font retrieveFont(boolean fontBoldweight, boolean fontItalic, Color fontColor, String fontName,
                             short fontHeightInPoints, byte fontUnderline, boolean fontStrikeout, int fontCharset, short fontTypeOffset)
    {
        String representation = getRepresentation(fontBoldweight, fontItalic, fontColor, fontName, fontHeightInPoints,
                fontUnderline, fontStrikeout, fontCharset, fontTypeOffset
        );
        Font f = myFontMap.get(representation);
        if (logger.isTraceEnabled())
        {
            if (f != null)
                logger.trace("FCache hit   : {}", representation);
            else
                logger.trace("FCache miss! : {}", representation);
        }
        return f;
    }

    /**
     * Caches the given <code>Font</code>.
     * @param f A <code>Font</code>.
     */
    public void cacheFont(Font f)
    {
        String representation = getRepresentation(f);
        logger.trace("Caching  f   : {}", representation);
        myFontMap.put(representation, f);
    }

    /**
     * Finds the given cached <code>Font</code> with the font characteristics of
     * the given <code>Font</code>.
     * @param f A <code>Font</code> which may not be in the cache.
     * @return The <code>Font</code> in the cache that matches <code>f</code>'s
     *    font characteristics, if it exists, else <code>null</code>.
     * @since 0.10.0
     */
    public Font findFont(Font f)
    {
        String representation = getRepresentation(f);
        return myFontMap.get(representation);
    }

    /**
     * Gets the string representation of the given <code>Font</code>.
     * @param f A <code>Font</code>.
     * @return The string representation.
     */
    private String getRepresentation(Font f)
    {
        // Colors that need an instanceof check
        Color fontColor;
        if (f instanceof HSSFFont)
        {
            HSSFFont hf = (HSSFFont) f;
            fontColor = hf.getHSSFColor((HSSFWorkbook) myWorkbook);
        }
        else if (f instanceof XSSFFont)
        {
            XSSFFont xf = (XSSFFont) f;
            fontColor = xf.getXSSFColor();
        }
        else
            throw new IllegalArgumentException("Bad Font type: " + f.getClass().getName());

        return getRepresentation(f.getBold(), f.getItalic(), fontColor, f.getFontName(),
                f.getFontHeightInPoints(), f.getUnderline(), f.getStrikeout(), f.getCharSet(), f.getTypeOffset());
    }

    /**
     * Return the string representation of a <code>Font</code> with the
     * given properties.
     * @param fontBoldweight The font boldweight.
     * @param fontItalic Whether the font is italic.
     * @param fontColor The font color.
     * @param fontName The font name.
     * @param fontHeightInPoints The font height in points.
     * @param fontUnderline The font underline.
     * @param fontStrikeout Whether the font is in strikeout.
     * @param fontCharset The font charset.
     * @param fontTypeOffset The font type offset.
     * @return The string representation.
     */
    private String getRepresentation(boolean fontBoldweight, boolean fontItalic, Color fontColor, String fontName,
                                     short fontHeightInPoints, byte fontUnderline, boolean fontStrikeout, int fontCharset, short fontTypeOffset)
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
        // Font underline
        buf.append(PROP_SEP).append(fontUnderline);
        // Font strikeout
        buf.append(PROP_SEP).append(fontStrikeout);
        // Font charset
        buf.append(PROP_SEP).append(fontCharset);
        // Font type offset
        buf.append(PROP_SEP).append(fontTypeOffset);

        return buf.toString();
    }
}