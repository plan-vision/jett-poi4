package net.sf.jett.model;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
/**
 * <p><code>Colors</code> represent the built-in color names that correspond
 * with Excel's indexed color scheme.  These are used in conjunction with
 * several property names defined for the style tag.  These color names do NOT
 * necessarily correspond with HTML/CSS standard color names.  Legal values are
 * the names of the enumeration objects, without underscores, case insensitive,
 * e.g. "center" == "Center" == "CENTER".</p>
 *
 * @author Randy Gettman
 * @since 0.4.0
 * @see net.sf.jett.tag.StyleTag
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BORDER_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_BOTTOM_BORDER_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_LEFT_BORDER_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_RIGHT_BORDER_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_TOP_BORDER_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_FILL_BACKGROUND_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_FILL_FOREGROUND_COLOR
 * @see net.sf.jett.parser.StyleParser#PROPERTY_FONT_COLOR
 */



public enum ExcelColor
{
    AQUA               (HSSFColor.HSSFColorPredefined.AQUA   , IndexedColors.AQUA                 , 51, 204, 204),
    AUTOMATIC          (HSSFColor.HSSFColorPredefined.AUTOMATIC            , IndexedColors.AUTOMATIC            , 0, 0, 0),
    BLACK              (HSSFColor.HSSFColorPredefined.BLACK                , IndexedColors.BLACK                , 0, 0, 0),
    BLUE               (HSSFColor.HSSFColorPredefined.BLUE                 , IndexedColors.BLUE                 , 0, 0, 255),
    BLUEGREY           (HSSFColor.HSSFColorPredefined.BLUE_GREY            , IndexedColors.BLUE_GREY            , 102, 102, 153),
    BRIGHTGREEN        (HSSFColor.HSSFColorPredefined.BRIGHT_GREEN         , IndexedColors.BRIGHT_GREEN         , 0, 255, 0),
    BROWN              (HSSFColor.HSSFColorPredefined.BROWN                , IndexedColors.BROWN                , 153, 51, 0),
    CORAL              (HSSFColor.HSSFColorPredefined.CORAL                , IndexedColors.CORAL                , 255, 128, 128),
    CORNFLOWERBLUE     (HSSFColor.HSSFColorPredefined.CORNFLOWER_BLUE      , IndexedColors.CORNFLOWER_BLUE      , 153, 153, 255),
    DARKBLUE           (HSSFColor.HSSFColorPredefined.DARK_BLUE            , IndexedColors.DARK_BLUE            , 0, 0, 128),
    DARKGREEN          (HSSFColor.HSSFColorPredefined.DARK_GREEN           , IndexedColors.DARK_GREEN           , 0, 51, 0),
    DARKRED            (HSSFColor.HSSFColorPredefined.DARK_RED             , IndexedColors.DARK_RED             , 128, 0, 0),
    DARKTEAL           (HSSFColor.HSSFColorPredefined.DARK_TEAL            , IndexedColors.DARK_TEAL            , 0, 51, 102),
    DARKYELLOW         (HSSFColor.HSSFColorPredefined.DARK_YELLOW          , IndexedColors.DARK_YELLOW          , 128, 128, 0),
    GOLD               (HSSFColor.HSSFColorPredefined.GOLD                 , IndexedColors.GOLD                 , 255, 204, 0),
    GREEN              (HSSFColor.HSSFColorPredefined.GREEN                , IndexedColors.GREEN                , 0, 128, 0),
    GREY25PERCENT      (HSSFColor.HSSFColorPredefined.GREY_25_PERCENT      , IndexedColors.GREY_25_PERCENT      , 192, 192, 192),
    GREY40PERCENT      (HSSFColor.HSSFColorPredefined.GREY_40_PERCENT      , IndexedColors.GREY_40_PERCENT      , 150, 150, 150),
    GREY50PERCENT      (HSSFColor.HSSFColorPredefined.GREY_50_PERCENT      , IndexedColors.GREY_50_PERCENT      , 128, 128, 128),
    GREY80PERCENT      (HSSFColor.HSSFColorPredefined.GREY_80_PERCENT      , IndexedColors.GREY_80_PERCENT      , 51, 51, 51),
    INDIGO             (HSSFColor.HSSFColorPredefined.INDIGO               , IndexedColors.INDIGO               , 51, 51, 153),
    LAVENDER           (HSSFColor.HSSFColorPredefined.LAVENDER             , IndexedColors.LAVENDER             , 204, 153, 255),
    LEMONCHIFFON       (HSSFColor.HSSFColorPredefined.LEMON_CHIFFON        , IndexedColors.LEMON_CHIFFON        , 255, 255, 204),
    LIGHTBLUE          (HSSFColor.HSSFColorPredefined.LIGHT_BLUE           , IndexedColors.LIGHT_BLUE           , 51, 102, 255),
    LIGHTCORNFLOWERBLUE(HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE, IndexedColors.LIGHT_CORNFLOWER_BLUE, 204, 204, 255),
    LIGHTGREEN         (HSSFColor.HSSFColorPredefined.LIGHT_GREEN          , IndexedColors.LIGHT_GREEN          , 204, 255, 204),
    LIGHTORANGE        (HSSFColor.HSSFColorPredefined.LIGHT_ORANGE         , IndexedColors.LIGHT_ORANGE         , 255, 153, 0),
    LIGHTTURQUOISE     (HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE      , IndexedColors.LIGHT_TURQUOISE      , 204, 255, 255),
    LIGHTYELLOW        (HSSFColor.HSSFColorPredefined.LIGHT_YELLOW         , IndexedColors.LIGHT_YELLOW         , 255, 255, 153),
    LIME               (HSSFColor.HSSFColorPredefined.LIME                 , IndexedColors.LIME                 , 153, 204, 0),
    MAROON             (HSSFColor.HSSFColorPredefined.MAROON               , IndexedColors.MAROON               , 128, 0, 0),
    OLIVEGREEN         (HSSFColor.HSSFColorPredefined.OLIVE_GREEN          , IndexedColors.OLIVE_GREEN          , 51, 51, 0),
    ORANGE             (HSSFColor.HSSFColorPredefined.ORANGE               , IndexedColors.ORANGE               , 255, 102, 0),
    ORCHID             (HSSFColor.HSSFColorPredefined.ORCHID               , IndexedColors.ORCHID               , 102, 0, 102),
    PALEBLUE           (HSSFColor.HSSFColorPredefined.PALE_BLUE            , IndexedColors.PALE_BLUE            , 153, 204, 255),
    PINK               (HSSFColor.HSSFColorPredefined.PINK                 , IndexedColors.PINK                 , 255, 0, 255),
    PLUM               (HSSFColor.HSSFColorPredefined.PLUM                 , IndexedColors.PLUM                 , 153, 51, 102),
    RED                (HSSFColor.HSSFColorPredefined.RED                  , IndexedColors.RED                  , 255, 0, 0),
    ROSE               (HSSFColor.HSSFColorPredefined.ROSE                 , IndexedColors.ROSE                 , 255, 103, 204),
    ROYALBLUE          (HSSFColor.HSSFColorPredefined.ROYAL_BLUE           , IndexedColors.ROYAL_BLUE           , 0, 102, 204),
    SEAGREEN           (HSSFColor.HSSFColorPredefined.SEA_GREEN            , IndexedColors.SEA_GREEN            , 51, 153, 102),
    SKYBLUE            (HSSFColor.HSSFColorPredefined.SKY_BLUE             , IndexedColors.SKY_BLUE             , 0, 204, 255),
    TAN                (HSSFColor.HSSFColorPredefined.TAN                  , IndexedColors.TAN                  , 255, 204, 153),
    TEAL               (HSSFColor.HSSFColorPredefined.TEAL                 , IndexedColors.TEAL                 , 0, 128, 128),
    TURQUOISE          (HSSFColor.HSSFColorPredefined.TURQUOISE            , IndexedColors.TURQUOISE            , 0, 255, 255),
    VIOLET             (HSSFColor.HSSFColorPredefined.VIOLET               , IndexedColors.VIOLET               , 128, 0, 128),
    WHITE              (HSSFColor.HSSFColorPredefined.WHITE                , IndexedColors.WHITE                , 255, 255, 255),
    YELLOW             (HSSFColor.HSSFColorPredefined.YELLOW               , IndexedColors.YELLOW               , 255, 255, 0);

    /**
     * The "automatic" color in HSSF (.xls).
     * @since 0.9.1
     */
    public static final HSSFColor.HSSFColorPredefined HSSF_COLOR_AUTOMATIC = HSSFColor.HSSFColorPredefined.AUTOMATIC;

    /**
     * The color index used by comments in XSSF (.xlsx).
     * @since 0.10.0
     */
    public static final short XSSF_COLOR_COMMENT = 81;

    private HSSFColor.HSSFColorPredefined myHssfColor;
    private XSSFColor myXssfColor;
    private IndexedColors myIndexedColor;
    private int myRed;
    private int myGreen;
    private int myBlue;

    private static HSSFColor[] hssfColors;

    static
    {
        hssfColors = new HSSFColor[65];
        for (ExcelColor excelColor : values())
        {
        	HSSFColor hssfColor = excelColor.getHssfColor();
            hssfColors[hssfColor.getIndex()] = hssfColor;
        }
    }

    /**
     * Creates a <code>ExcelColor</code>.
     * @param hssfColor The <code>HSSFColor</code>.
     * @param indexedColor The <code>IndexedColor</code>.
     * @param red The red value, 0-255.
     * @param green The green value, 0-255.
     * @param blue The blue value, 0-255.
     */
    ExcelColor(HSSFColor.HSSFColorPredefined hssfColor, IndexedColors indexedColor, int red, int green, int blue)
    {
        myHssfColor = hssfColor;
        myXssfColor = new XSSFColor(new byte[] {(byte) red, (byte) green, (byte) blue},new DefaultIndexedColorMap());
        myIndexedColor = indexedColor;
        myRed = red;
        myGreen = green;
        myBlue = blue;
    }

    /**
     * Return the <code>HSSFColor</code>.
     * @return The <code>HSSFColor</code>.
     */
    public HSSFColor getHssfColor()
    {
    	if (myHssfColor == null)
    		return null;
        return myHssfColor.getColor();
    }

    /**
     * Return the <code>XSSFColor</code>.
     * @return The <code>XSSFColor</code>.
     */
    public XSSFColor getXssfColor()
    {
        return myXssfColor;
    }

    /**
     * Returns the index.
     * @return The index.
     */
    public int getIndex()
    {
        return myIndexedColor.getIndex();
    }

    /**
     * Returns the <code>IndexedColors</code>.
     * @return The <code>IndexedColors</code>.
     */
    public IndexedColors getIndexedColor()
    {
        return myIndexedColor;
    }

    /**
     * Returns the red value, 0-255.
     * @return The red value, 0-255.
     */
    public int getRed()
    {
        return myRed;
    }

    /**
     * Returns the green value, 0-255.
     * @return The green value, 0-255.
     */
    public int getGreen()
    {
        return myGreen;
    }

    /**
     * Returns the blue value, 0-255.
     * @return The blue value, 0-255.
     */
    public int getBlue()
    {
        return myBlue;
    }

    /**
     * Returns the hex string, in the format "#RRGGBB".
     * @return The hex string, in the format "#RRGGBB".
     */
    public String getHexString()
    {
        StringBuilder builder = new StringBuilder();
        builder.append("#");

        String value = Integer.toHexString(myRed);
        if (value.length() == 1)
            builder.append("0");
        builder.append(value);

        value = Integer.toHexString(myGreen);
        if (value.length() == 1)
            builder.append("0");
        builder.append(value);

        value = Integer.toHexString(myBlue);
        if (value.length() == 1)
            builder.append("0");
        builder.append(value);

        return builder.toString();
    }

    /**
     * Returns the "distance" of the given RGB triplet from this color, as
     * defined by the sum of each of the differences for the red, green, and
     * blue values.
     * @param red The red value.
     * @param green The green value.
     * @param blue The blue value.
     * @return The sum of each of the differences for the red, green, and blue
     *    values.
     */
    public int distance(int red, int green, int blue)
    {
        return Math.abs(red - myRed) + Math.abs(green - myGreen) + Math.abs(blue - myBlue);
    }

    /**
     * Returns the color name, in all lowercase, no underscores or spaces.
     * @return The color name, in all lowercase, no underscores or spaces.
     */
    @Override
    public String toString()
    {
        return name().trim().toLowerCase().replace("_", "");
    }

    /**
     * Maps a short index color back to an <code>HSSFColor</code>.
     * @param index A short color index.
     * @return An <code>HSSFColor</code>.
     */
    public static HSSFColor getHssfColorByIndex(short index)
    {
        if (index == Font.COLOR_NORMAL || index == XSSF_COLOR_COMMENT)
            return HSSF_COLOR_AUTOMATIC.getColor();
        return hssfColors[index];
    }
}
