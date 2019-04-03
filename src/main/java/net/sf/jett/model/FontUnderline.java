package net.sf.jett.model;

/**
 * <p><code>FontUnderlines</code> represent the built-in underline names that
 * correspond with Excel's underlining scheme.  These are used in conjunction
 * with the underline property in the style tag.  Legal values are the names of
 * the enumeration objects, without underscores, case insensitive, e.g.
 * "single" == "Single" == "SINGLE".</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 * @see net.sf.jett.tag.StyleTag
 * @see net.sf.jett.parser.StyleParser#PROPERTY_FONT_UNDERLINE
 */
public enum FontUnderline
{
    SINGLE          (org.apache.poi.ss.usermodel.FontUnderline.SINGLE.getByteValue()),
    DOUBLE          (org.apache.poi.ss.usermodel.FontUnderline.DOUBLE.getByteValue()),
    SINGLEACCOUNTING(org.apache.poi.ss.usermodel.FontUnderline.SINGLE_ACCOUNTING.getByteValue()),
    DOUBLEACCOUNTING(org.apache.poi.ss.usermodel.FontUnderline.DOUBLE_ACCOUNTING.getByteValue()),
    NONE            (org.apache.poi.ss.usermodel.FontUnderline.NONE.getByteValue());

    private byte myIndex;

    /**
     * Constructs an <code>FontUnderline</code>.
     * @param index The index.
     */
    FontUnderline(byte index)
    {
        myIndex = index;
    }

    /**
     * Returns the index.
     * @return The index.
     */
    public byte getIndex()
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