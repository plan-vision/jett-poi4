package net.sf.jett.model;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * <p><code>BorderTypes</code> represent the built-in border type names that
 * correspond with Excel's border type scheme.  These are used in
 * conjunction with the border-related properties in the style tag.  Legal
 * values are the names of the enumeration objects, without underscores, case
 * insensitive, e.g. "thin" == "Thin" == "THIN".</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 * @see net.sf.jett.tag.StyleTag
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BORDER
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BORDER_BOTTOM
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BORDER_LEFT
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BORDER_RIGHT
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BORDER_TOP
 */
public enum BorderType
{
    NONE            ((short) BorderStyle.NONE.ordinal()),
    THIN            ((short) BorderStyle.THIN.ordinal()),
    MEDIUM          ((short) BorderStyle.MEDIUM.ordinal()),
    DASHED          ((short) BorderStyle.DASHED.ordinal()),
    HAIR            ((short) BorderStyle.HAIR.ordinal()),
    THICK           ((short) BorderStyle.THICK.ordinal()),
    DOUBLE          ((short) BorderStyle.DOUBLE.ordinal()),
    DOTTED          ((short) BorderStyle.DOTTED.ordinal()),
    MEDIUMDASHED    ((short) BorderStyle.MEDIUM_DASHED.ordinal()),
    DASHDOT         ((short) BorderStyle.DASH_DOT.ordinal()),
    MEDIUMDASHDOT   ((short) BorderStyle.MEDIUM_DASH_DOT.ordinal()),
    DASHDOTDOT      ((short) BorderStyle.DASH_DOT_DOT.ordinal()),
    MEDIUMDASHDOTDOT((short) BorderStyle.MEDIUM_DASH_DOT_DOT.ordinal()),  // DOTC: [sic]
    SLANTEDDASHDOT  ((short) BorderStyle.SLANTED_DASH_DOT.ordinal());

    private short myIndex;

    /**
     * Constructs a <code>BorderType</code>.
     * @param index The index.
     */
    BorderType(short index)
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
