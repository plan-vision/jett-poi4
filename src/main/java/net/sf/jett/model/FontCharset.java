package net.sf.jett.model;

/**
 * <p><code>Charsets</code> represent the built-in type charset names that
 * correspond with Excel's charset scheme.  These are used in conjunction with
 * the charset property in the style tag.</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 * @see net.sf.jett.tag.StyleTag
 * @see net.sf.jett.parser.StyleParser#PROPERTY_FONT_CHARSET
 */
public enum FontCharset
{
    ANSI       (org.apache.poi.ss.usermodel.FontCharset.ANSI.getValue()),
    DEFAULT    (org.apache.poi.ss.usermodel.FontCharset.DEFAULT.getValue()),
    SYMBOL     (org.apache.poi.ss.usermodel.FontCharset.SYMBOL.getValue()),
    MAC        (org.apache.poi.ss.usermodel.FontCharset.MAC.getValue()),
    SHIFTJIS   (org.apache.poi.ss.usermodel.FontCharset.SHIFTJIS.getValue()),
    HANGEUL    (org.apache.poi.ss.usermodel.FontCharset.HANGEUL.getValue()),
    JOHAB      (org.apache.poi.ss.usermodel.FontCharset.JOHAB.getValue()),
    GB2312     (org.apache.poi.ss.usermodel.FontCharset.GB2312.getValue()),
    CHINESEBIG5(org.apache.poi.ss.usermodel.FontCharset.CHINESEBIG5.getValue()),
    GREEK      (org.apache.poi.ss.usermodel.FontCharset.GREEK.getValue()),
    TURKISH    (org.apache.poi.ss.usermodel.FontCharset.TURKISH.getValue()),
    VIETNAMESE (org.apache.poi.ss.usermodel.FontCharset.VIETNAMESE.getValue()),
    HEBREW     (org.apache.poi.ss.usermodel.FontCharset.HEBREW.getValue()),
    ARABIC     (org.apache.poi.ss.usermodel.FontCharset.ARABIC.getValue()),
    BALTIC     (org.apache.poi.ss.usermodel.FontCharset.BALTIC.getValue()),
    RUSSIAN    (org.apache.poi.ss.usermodel.FontCharset.RUSSIAN.getValue()),
    THAI       (org.apache.poi.ss.usermodel.FontCharset.THAI.getValue()),
    EASTEUROPE (org.apache.poi.ss.usermodel.FontCharset.EASTEUROPE.getValue()),
    OEM        (org.apache.poi.ss.usermodel.FontCharset.OEM.getValue());

    private int myIndex;

    /**
     * Constructs a <code>Charset</code>.
     * @param index The index.
     */
    FontCharset(int index)
    {
        myIndex = index;
    }

    /**
     * Returns the index.
     * @return The index.
     */
    public int getIndex()
    {
        return myIndex;
    }

    /**
     * Returns the name, in all lowercase, no underscores or spaces.
     * @return The name, in all lowercase, no underscores or spaces.
     */
    public String toString()
    {
        return name().trim().toLowerCase().replace("_", "");
    }
}