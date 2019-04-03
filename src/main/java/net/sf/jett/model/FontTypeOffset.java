package net.sf.jett.model;

import org.apache.poi.ss.usermodel.Font;

/**
 * <p><code>FontTypeOffsets</code> represent the built-in type offset names
 * that correspond with Excel's type offset scheme.  These are used in
 * conjunction with the type offset property in the style tag.  Legal values
 * are the names of the enumeration objects, without underscores, case
 * insensitive, e.g. "sub" == "Sub" == "SUB".</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 * @see net.sf.jett.tag.StyleTag
 * @see net.sf.jett.parser.StyleParser#PROPERTY_FONT_TYPE_OFFSET
 */
public enum FontTypeOffset
{
    NONE (Font.SS_NONE),
    SUB  (Font.SS_SUB),
    SUPER(Font.SS_SUPER);

    private short myIndex;

    /**
     * Constructs a <code>FontTypeOffset</code>.
     * @param index The index.
     */
    FontTypeOffset(short index)
    {
        myIndex = index;
    }

    /**
     * Returns the index.
     * @return The index.
     */
    public short getIndex()
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