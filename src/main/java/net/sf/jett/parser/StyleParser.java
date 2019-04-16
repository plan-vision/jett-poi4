package net.sf.jett.parser;

import java.util.HashMap;
import java.util.Map;
import java.util.logging.LogManager;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.exception.StyleParseException;
import net.sf.jett.model.FontCharset;
import net.sf.jett.model.FontTypeOffset;
import net.sf.jett.model.FontUnderline;
import net.sf.jett.model.Style;

/**
 * <p>A <code>StyleParser</code> parses "CSS" text, from beginning to end, in a
 * CSS-like format:
 * <code>[.styleName { propertyName: value [; propertyName: value]* }]*</code></p>
 * <p>If a property value is an empty string or the property is not present,
 * then it will be ignored.  Unrecognized property names and unrecognized
 * values for a property are ignored.  Property names and values may be
 * specified in a case insensitive-fashion, i.e. "CENTER" = "Center" =
 * "center".</p>
 * <p>Both CSS files and the "style" attribute of the "style" tag recognize the
 * following property names.</p>
 *
 * <p>Properties:  The following properties control alignment, borders, colors,
 * etc., everything but the font characteristics.
 * <ul>
 *    <li><code>alignment</code> - Controls horizontal alignment, with one of
 *    the values taken from <code>Alignment.toString()</code>.</li>
 *    <li><code>border</code> - Controls all 4 borders for the cell, with one
 *    of the values taken from <code>BorderType.toString()</code>.</li>
 *    <li><code>border-bottom</code> - Controls the bottom border for the cell,
 *    with one of the values taken from <code>BorderType.toString()</code>.</li>
 *    <li><code>border-left</code> - Controls the left border for the cell,
 *    with one of the values taken from <code>BorderType.toString()</code>.</li>
 *    <li><code>border-right</code> - Controls the right border for the cell,
 *    with one of the values taken from <code>BorderType.toString()</code>.</li>
 *    <li><code>border-top</code> - Controls the top border for the cell,
 *    with one of the values taken from <code>BorderType.toString()</code>.</li>
 *    <li><code>border-color</code> - Controls the color of all 4 borders for
 *    the cell, with a hex value ("#rrggbb") or one of 48 Excel-based color
 *    names defined by <code>ExcelColor.toString()</code>.  For ".xls" files,
 *    if a hex value is supplied, then the supported color name that is closest
 *    to the given value is used.</li>
 *    <li><code>bottom-border-color</code> - Controls the color of the bottom
 *    border for the cell, with a hex value ("#rrggbb") or one of the above 48
 *    color names mentioned above.</li>
 *    <li><code>left-border-color</code> - Controls the color of the left
 *    border for the cell, with a hex value ("#rrggbb") or one of the above 48
 *    color names mentioned above.</li>
 *    <li><code>right-border-color</code> - Controls the color of the right
 *    border for the cell, with a hex value ("#rrggbb") or one of the above 48
 *    color names mentioned above.</li>
 *    <li><code>top-border-color</code> - Controls the color of the top
 *    border for the cell, with a hex value ("#rrggbb") or one of the above 48
 *    color names mentioned above.</li>
 *    <li><code>column-width-in-chars</code> - Controls the width of the cell's
 *    column, in number of characters.</li>
 *    <li><code>data-format</code> - Controls the Excel numeric or date format
 *    string.</li>
 *    <li><code>fill-background-color</code> - Controls the "background color"
 *    of the fill pattern, with one of the color values mentioned above.</li>
 *    <li><code>fill-foreground-color</code> - Controls the "foreground color"
 *    of the fill pattern, with one of the color values mentioned above.</li>
 *    <li><code>fill-pattern</code> - Controls the "fill pattern", with one of
 *    the values taken from <code>FillPattern.toString()</code>:</li>
 *    <li><code>hidden</code> - Controls the "hidden" property with a
 *    <code>true</code> or <code>false</code> value.</li>
 *    <li><code>indention</code> - Controls the number of characters that the
 *    text is indented.</li>
 *    <li><code>locked</code> - Controls the "locked" property with a
 *    <code>true</code> or <code>false</code> value.</li>
 *    <li><code>rotation</code> - Controls the number of degrees the text is
 *    rotated, from -90 to +90, or <code>ROTATION_STACKED</code> for stacked
 *    text.</li>
 *    <li><code>row-height-in-points</code> - Controls the height of the cell's
 *    row, in points.</li>
 *    <li><code>vertical-alignment</code> - Controls horizontal alignment, with
 *    one of the values taken from <code>VerticalAlignment.toString()</code>:</li>
 *    <li><code>wrap-text</code> - Controls whether long text values are
 *    wrapped onto the next physical line with a cell, with a <code>true</code>
 *    or <code>false</code> value.</li>
 * </ul>
 * <p>Properties:  The following properties control the font characteristics.
 * <ul>
 *    <li><code>font-weight</code> - Controls how bold the text appears, with
 *    the values taken from <code>FontBoldweight.toString()</code>.</li>
 *    <li><code>font-charset</code> - Controls the character set, with the
 *    values taken from <code>Charset.toString()</code>.</li>
 *    <li><code>font-color</code> - Controls the color of the text, with a hex
 *    value ("#rrggbb") or one of the color names mentioned above.</li>
 *    <li><code>font-height-in-points</code> - Controls the font height, in
 *    points.</li>
 *    <li><code>font-name</code> - Controls the font name, e.g. "Arial".</li>
 *    <li><code>font-italic</code> - Controls whether the text is
 *    <em>italic</em>, with a <code>true</code> or <code>false</code> value.</li>
 *    <li><code>font-strikeout</code> - Controls whether the text is
 *    <span style="text-decoration: line-through">strikeout</span>, with a
 *    <code>true</code> or <code>false</code> value.</li>
 *    <li><code>font-type-offset</code> - Controls the text offset, e.g.
 *    <sup>superscript</sup> and <sub>subscript</sub>, with the values taken
 *    from <code>FontTypeOffset.toString()</code>.</li>
 *    <li><code>font-underline</code> - Controls whether and how the text is
 *    underlined, with the values taken from <code>Underline.toString()</code>.</li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class StyleParser
{
    private static final Logger logger = LoggerFactory.getLogger(StyleParser.class);

    /**
     * The property to specify horizontal alignment of the text.
     * @see net.sf.jett.model.Alignment
     */
    public static final String PROPERTY_ALIGNMENT = "alignment";
    /**
     * The property to specify the type of all 4 borders.
     * @see net.sf.jett.model.BorderType
     */
    public static final String PROPERTY_BORDER = "border";
    /**
     * The property to specify the type of the bottom border.
     * @see net.sf.jett.model.BorderType
     */
    public static final String PROPERTY_BORDER_BOTTOM = "border-bottom";
    /**
     * The property to specify the type of the left border.
     * @see net.sf.jett.model.BorderType
     */
    public static final String PROPERTY_BORDER_LEFT = "border-left";
    /**
     * The property to specify the type of the right border.
     * @see net.sf.jett.model.BorderType
     */
    public static final String PROPERTY_BORDER_RIGHT = "border-right";
    /**
     * The property to specify the type of the top border.
     * @see net.sf.jett.model.BorderType
     */
    public static final String PROPERTY_BORDER_TOP = "border-top";
    /**
     * The property to specify the color of all 4 borders.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_BORDER_COLOR = "border-color";
    /**
     * The property to specify the color of the bottom border.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_BOTTOM_BORDER_COLOR = "bottom-border-color";
    /**
     * The property to specify the color of the left border.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_LEFT_BORDER_COLOR = "left-border-color";
    /**
     * The property to specify the color of the right border.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_RIGHT_BORDER_COLOR = "right-border-color";
    /**
     * The property to specify the color of the top border.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_TOP_BORDER_COLOR = "top-border-color";
    /**
     * The property to specify the width of the column in number of characters.
     */
    public static final String PROPERTY_COLUMN_WIDTH_IN_CHARS = "column-width-in-chars";
    /**
     * The property to specify the numeric or date data format string.
     */
    public static final String PROPERTY_DATA_FORMAT = "data-format";
    /**
     * The property to specify the fill background color to be used in a fill
     * pattern.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_FILL_BACKGROUND_COLOR = "fill-background-color";
    /**
     * The property to specify the fill foreground color to be used in a fill
     * pattern.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_FILL_FOREGROUND_COLOR = "fill-foreground-color";
    /**
     * The property to specify the fill pattern to be used with the fill
     * foreground color and the fill background color.
     * @see net.sf.jett.model.FillPattern
     */
    public static final String PROPERTY_FILL_PATTERN = "fill-pattern";
    /**
     * The property to specify the "hidden" property.
     */
    public static final String PROPERTY_HIDDEN = "hidden";
    /**
     * The property to specify the number of characters that the text is
     * indented.
     */
    public static final String PROPERTY_INDENTION = "indention";
    /**
     * The property to specify the "locked" property.
     */
    public static final String PROPERTY_LOCKED = "locked";
    /**
     * The property to specify the number of degrees that the text is rotated,
     * from -90 to +90.
     */
    public static final String PROPERTY_ROTATION = "rotation";
    /**
     * The property to specify the height of the row in points.
     */
    public static final String PROPERTY_ROW_HEIGHT_IN_POINTS = "row-height-in-points";
    /**
     * The property to specify the vertical alignment of the text.
     * @see net.sf.jett.model.VerticalAlignment
     */
    public static final String PROPERTY_VERTICAL_ALIGNMENT = "vertical-alignment";
    /**
     * The property to specify whether long text values are wrapped to the next
     * physical line within the cell.
     */
    public static final String PROPERTY_WRAP_TEXT = "wrap-text";
    /**
     * The property to specify whether the font is bold.
     */
    public static final String PROPERTY_FONT_BOLDWEIGHT = "font-weight";
    /**
     * The property to specify the charset used by the font.
     * @see net.sf.jett.model.FontCharset
     */
    public static final String PROPERTY_FONT_CHARSET = "font-charset";
    /**
     * The property to specify the font color.
     * @see net.sf.jett.model.ExcelColor
     */
    public static final String PROPERTY_FONT_COLOR = "font-color";
    /**
     * The property to specify the font height in points.
     */
    public static final String PROPERTY_FONT_HEIGHT_IN_POINTS = "font-height-in-points";
    /**
     * The property to specify the font name.
     */
    public static final String PROPERTY_FONT_NAME = "font-name";
    /**
     * The property to specify whether the font is italic.
     */
    public static final String PROPERTY_FONT_ITALIC = "font-italic";
    /**
     * The property to specify whether the font is strikeout.
     */
    public static final String PROPERTY_FONT_STRIKEOUT = "font-strikeout";
    /**
     * The property to specify whether the font type is offset, and if it is,
     * whether it's superscript or subscript.
     * @see net.sf.jett.model.FontTypeOffset
     */
    public static final String PROPERTY_FONT_TYPE_OFFSET = "font-type-offset";
    /**
     * The property to specify how the font text is underlined.
     * @see net.sf.jett.model.FontUnderline
     */
    public static final String PROPERTY_FONT_UNDERLINE = "font-underline";

    /**
     * <p>Specify this value of rotation to use to produce vertically</p>
     * <br>s
     * <br>t
     * <br>a
     * <br>c
     * <br>k
     * <br>e
     * <br>d
     * <p>text.</p>
     * @see #PROPERTY_ROTATION
     */
    public static final String ROTATION_STACKED = "STACKED";
    /**
     * <p>POI value of rotation to use to produce vertically</p>
     * <br>s
     * <br>t
     * <br>a
     * <br>c
     * <br>k
     * <br>e
     * <br>d
     * <p>text.</p>
     * @see #PROPERTY_ROTATION
     */
    public static final short POI_ROTATION_STACKED = 0xFF;

    private enum State
    {
        START,
        EXPECT_STYLE_NAME,
        EXPECT_BEGIN_BRACE,
        EXPECT_PROPERTY_NAME,
        EXPECT_COLON,
        EXPECT_VALUE,
        EXPECT_SEMICOLON_OR_END_BRACE
    }

    private String myCssText;
    private State myState;
    private Map<String, Style> myStyleMap;

    /**
     * Create a <code>StyleParser</code>.
     */
    public StyleParser()
    {
        setCssText("");
    }

    /**
     * Create a <code>StyleParser</code> object that will parse the given
     * css text.
     * @param cssText The CSS text.
     */
    public StyleParser(String cssText)
    {
        setCssText(cssText);
    }

    /**
     * Sets the CSS text to the given CSS text and resets the parser.
     * @param cssText The new CSS text.
     */
    public void setCssText(String cssText)
    {
        myCssText = cssText;
        reset();
    }

    /**
     * Resets this <code>StyleParser</code>, usually at creation time and
     * when new input arrives.
     */
    private void reset()
    {
        myState = State.START;
        myStyleMap = new HashMap<>();
    }

    /**
     * Parses the CSS text.
     */
    public void parse()
    {
        StyleScanner scanner = new StyleScanner(myCssText);

        StyleScanner.Token token = scanner.getNextToken();
        if (token == StyleScanner.Token.TOKEN_WHITESPACE)
            token = scanner.getNextToken();

        // Parse any CSS style definitions:
        // [.styleName { propertyName: value [; propertyName: value]* }]*
        String styleName = null;
        String propertyName = null;
        String value = null;
        Style currStyle = null;
        while (token.getCode() >= 0 && token != StyleScanner.Token.TOKEN_EOI)
        {
            logger.debug("Token: {}, lexeme: \"{}\"", token, scanner.getCurrLexeme());
            switch(token)
            {
            case TOKEN_WHITESPACE:
                // Look out for multi-word values.
                if (myState == State.EXPECT_SEMICOLON_OR_END_BRACE)
                {
                    value += scanner.getCurrLexeme();
                }
                break;
            case TOKEN_STRING:
                String lexeme = scanner.getCurrLexeme();
                switch (myState)
                {
                case EXPECT_STYLE_NAME:
                    styleName = lexeme;
                    myState = State.EXPECT_BEGIN_BRACE;
                    break;
                case EXPECT_PROPERTY_NAME:
                    propertyName = lexeme;
                    myState = State.EXPECT_COLON;
                    break;
                case EXPECT_VALUE:
                    value = lexeme;
                    myState = State.EXPECT_SEMICOLON_OR_END_BRACE;
                    break;
                case START:
                    throw new StyleParseException("Expected new style definition, got " + lexeme + ": \"" + myCssText + "\"");
                case EXPECT_BEGIN_BRACE:
                    throw new StyleParseException("Expected '{', got " + lexeme + ": \"" + myCssText + "\"");
                case EXPECT_COLON:
                    throw new StyleParseException("Expected ':', got " + lexeme + ": \"" + myCssText + "\"");
                case EXPECT_SEMICOLON_OR_END_BRACE:
                    // Watch out for multi-word values, e.g. "Times New Roman".
                    value += lexeme;
                    break;
                }
                break;
            case TOKEN_PERIOD:
                if (myState != State.START)
                    throw new StyleParseException("Unexpected '.': \"" + myCssText + "\"");
                myState = State.EXPECT_STYLE_NAME;
                currStyle = new Style();
                break;
            case TOKEN_SEMICOLON:
                if (myState != State.EXPECT_SEMICOLON_OR_END_BRACE)
                    throw new StyleParseException("Unexpected ';': \"" + myCssText + "\"");
                addStyle(currStyle, propertyName, value);
                propertyName = null;
                value = null;
                myState = State.EXPECT_PROPERTY_NAME;
                break;
            case TOKEN_BEGIN_BRACE:
                if (myState != State.EXPECT_BEGIN_BRACE)
                    throw new StyleParseException("Unexpected '{': \"" + myCssText + "\"");
                myState = State.EXPECT_PROPERTY_NAME;
                break;
            case TOKEN_END_BRACE:
                if (myState != State.EXPECT_SEMICOLON_OR_END_BRACE && myState != State.EXPECT_PROPERTY_NAME)
                    throw new StyleParseException("Unexpected '}': \"" + myCssText + "\"");
                if (propertyName != null && value != null)
                {
                    addStyle(currStyle, propertyName, value);
                    propertyName = null;
                    value = null;
                }
                if (styleName != null)
                {
                    myStyleMap.put(styleName, currStyle);
                }
                styleName = null;
                myState = State.START;
                break;
            case TOKEN_COLON:
                if (myState != State.EXPECT_COLON)
                    throw new StyleParseException("Unexpected ':': \"" + myCssText + "\"");
                myState = State.EXPECT_VALUE;
                break;
            case TOKEN_ERROR_EOI_IN_COMMENT:
                throw new StyleParseException("End of input reached while scanning comment: \"" + myCssText + "\"");
            default:
                throw new StyleParseException("Parse error occurred: \"" + myCssText + "\"");
            }
            token = scanner.getNextToken();
        }
        // Found end of input before attribute value found.
        if (myState != State.START)
            throw new StyleParseException("Found end of input before end of style definition at \"" +
                    styleName + "\" (" + myState + "): \"" + myCssText + "\"");
        if (token.getCode() < 0)
            throw new StyleParseException("Found end of input while scanning comment: \"" + myCssText + "\"");
    }

    /**
     * Depending on the given property, parse the given value and set the
     * appropriate attribute in the given <code>Style</code> object.
     * @param style A <code>Style</code>.
     * @param property A property name, which should be one of the property name
     *    constants defined in this class.
     * @param value The property value, the meaning of which is
     *    property-specific.
     */
    public static void addStyle(Style style, String property, String value)
    {
        logger.debug("property: {}, value: {}", property, value);
        // Case insensitive property names and values.
        property = property.toLowerCase();
        value = value.trim().toUpperCase();
        // Try for descending order of popularity.  This order should match
        // the order of properties in examineAndApplyStyle(), but if it
        // doesn't match, then nothing will break.
        if (PROPERTY_FONT_BOLDWEIGHT.equals(property))
        {
            try
            {
                style.setFontBoldweight(value.equals("BOLD") || value.equals("TRUE"));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal font boldweight found: {}.  IllegalArgumentException caught: {}",
                        value, e.getMessage());
            }
        }
        else if (PROPERTY_FONT_ITALIC.equals(property))
        {
            if (value != null)
                style.setFontItalic(Boolean.valueOf(value));
        }
        else if (PROPERTY_FONT_COLOR.equals(property))
        {
            if (value != null)
                style.setFontColor(value);
        }
        else if (PROPERTY_FONT_NAME.equals(property))
        {
            if (value != null)
                style.setFontName(value);
        }
        else if (PROPERTY_FONT_HEIGHT_IN_POINTS.equals(property))
        {
            try
            {
                style.setFontHeightInPoints(Short.valueOf(value));
            }
            catch (NumberFormatException e)
            {
                logger.debug("Illegal font height in points: {}.  NumberFormatException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_ALIGNMENT.equals(property))
        {
        	
        	/* TRANSLATE TODO MOVE */
        	switch (value) 
        	{
        		case "CENTERSELECTION":          value="CENTER_SELECTION";break;
        	}
            try
            {
                style.setAlignment(HorizontalAlignment.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal property alignment: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_BORDER.equals(property))
        {
        	/* TRANSLATE TODO MOVE */
        	switch(value) {
        		case "DASHDOT" : 		  value="DASH_DOT";break;
        		case "MEDIUMDASHDOT" :    value="MEDIUM_DASH_DOT";break;
        		case "DASHDOTDOT" : 	  value="DASH_DOT_DOT";break;
        		case "MEDIUMDASHDOTDOT" : value="MEDIUM_DASH_DOT_DOT";break;
        		case "SLANTEDDASHDOT" :   value="SLANTED_DASH_DOT";break;
        	}
            try
            {
                BorderStyle bt = BorderStyle.valueOf(value);
                style.setBorderBottomType(bt);
                style.setBorderLeftType(bt);
                style.setBorderRightType(bt);
                style.setBorderTopType(bt);
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border type: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_DATA_FORMAT.equals(property))
        {
            if (value != null)
            {
                style.setDataFormat(value);
            }
        }
        else if (PROPERTY_FONT_UNDERLINE.equals(property))
        {
            try
            {
                style.setFontUnderline(FontUnderline.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal font underline type: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_FONT_STRIKEOUT.equals(property))
        {
            if (value != null)
                style.setFontStrikeout(Boolean.valueOf(value));
        }
        else if (PROPERTY_WRAP_TEXT.equals(property))
        {
            if (value != null)
                style.setWrappingText(Boolean.valueOf(value));
        }
        else if (PROPERTY_FILL_BACKGROUND_COLOR.equals(property))
        {
            if (value != null)
                style.setFillBackgroundColor(value);
        }
        else if (PROPERTY_FILL_FOREGROUND_COLOR.equals(property))
        {
            if (value != null)
                style.setFillForegroundColor(value);
        }
        else if (PROPERTY_FILL_PATTERN.equals(property))
        {
        	/* TRANSLATE TODO MOVE */
            try
            {
            	switch (value) 
            	{
            		case "NOFILL":                   value="NO_FILL";break;
            		case "SOLID":                    value="SOLID_FOREGROUND";break;
            		case "GRAY50PERCENT":            value="FINE_DOTS";break;
            		case "GRAY75PERCENT":            value="ALT_BARS";break;
            		case "GRAY25PERCENT":            value="SPARSE_DOTS";break;
            		case "HORIZONTALSTRIPE":         value="THICK_HORZ_BANDS";break;
            		case "VERTICALSTRIPE":           value="THICK_VERT_BANDS";break;
            		case "REVERSEDIAGONALSTRIPE":    value="THICK_BACKWARD_DIAG";break;
            		case "DIAGONALSTRIPE":           value="THICK_FORWARD_DIAG";break;
            		case "DIAGONALCROSSHATCH":       value="BIG_SPOTS";break;
            		case "THICKDIAGONALCROSSHATCH":  value="BRICKS";break;
            		case "THINHORIZONTALSTRIPE":     value="THIN_HORZ_BANDS";break;
            		case "THINVERTICALSTRIPE":       value="THIN_VERT_BANDS";break;
            		case "THINREVERSEDIAGONALSTRIPE":value="THIN_BACKWARD_DIAG";break;
            		case "THINDIAGONALSTRIPE":       value="THIN_FORWARD_DIAG";break;
            		case "THINHORIZONTALCROSSHATCH": value="SQUARES";break;
            		case "THINDIAGONALCROSSHATCH":   value="DIAMONDS";break;
            		case "GRAY12PERCENT":            value="LESS_DOTS";break;
            		case "GRAY6PERCENT":             value="LEAST_DOTS";break;
            	}
                style.setFillPatternType(FillPatternType.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal fill pattern: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_VERTICAL_ALIGNMENT.equals(property))
        {
            try
            {
                style.setVerticalAlignment(VerticalAlignment.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal vertical alignment: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_INDENTION.equals(property))
        {
            try
            {
                style.setIndention(Short.valueOf(value));
            }
            catch (NumberFormatException e)
            {
                logger.debug("Illegal property indention: {}.  NumberFormatException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_ROTATION.equals(property))
        {
            if (ROTATION_STACKED.equals(value))
            {
                style.setRotationDegrees(POI_ROTATION_STACKED);
            }
            else
            {
                try
                {
                    style.setRotationDegrees(Short.valueOf(value));
                }
                catch (NumberFormatException e)
                {
                    logger.debug("Illegal property rotation: {}.  NumberFormatException caught: {}", value,  e.getMessage());
                }
            }
        }
        else if (PROPERTY_COLUMN_WIDTH_IN_CHARS.equals(property))
        {
            try
            {
                double width = Double.parseDouble(value);
                style.setColumnWidth((int) Math.round(256 * width));
            }
            catch (NumberFormatException e)
            {
                logger.debug("Illegal column width in chars: {}.  NumberFormatException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_ROW_HEIGHT_IN_POINTS.equals(property))
        {
            try
            {
                double height = Double.parseDouble(value);
                style.setRowHeight((short) Math.round(20 * height));
            }
            catch (NumberFormatException e)
            {
                logger.debug("Illegal row height in points: {}.  NumberFormatException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_BORDER_COLOR.equals(property))
        {
            try
            {
                style.setBorderBottomColor(value);
                style.setBorderLeftColor(value);
                style.setBorderRightColor(value);
                style.setBorderTopColor(value);
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border color: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_FONT_CHARSET.equals(property))
        {
            try
            {
                style.setFontCharset(FontCharset.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal font charset: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_FONT_TYPE_OFFSET.equals(property))
        {
            try
            {
                style.setFontTypeOffset(FontTypeOffset.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal font type offset: {}.  IllegalArgumentException caught: {}", value, e.getMessage());
            }
        }
        else if (PROPERTY_LOCKED.equals(property))
        {
            if (value != null)
                style.setLocked(Boolean.valueOf(value));
        }
        else if (PROPERTY_HIDDEN.equals(property))
        {
            if (value != null)
                style.setHidden(Boolean.valueOf(value));
        }
        else if (PROPERTY_BORDER_BOTTOM.equals(property))
        {
            try
            {
                style.setBorderBottomType(BorderStyle.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border bottom: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_BORDER_LEFT.equals(property))
        {
            try
            {
                style.setBorderLeftType(BorderStyle.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border left: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_BORDER_RIGHT.equals(property))
        {
            try
            {
                style.setBorderRightType(BorderStyle.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border right: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_BORDER_TOP.equals(property))
        {
            try
            {
                style.setBorderTopType(BorderStyle.valueOf(value));
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border top: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_BOTTOM_BORDER_COLOR.equals(property))
        {
            try
            {
                if (value != null)
                    style.setBorderBottomColor(value);
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal border color: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_LEFT_BORDER_COLOR.equals(property))
        {
            try
            {
                if (value != null)
                    style.setBorderLeftColor(value);
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal left border color: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_RIGHT_BORDER_COLOR.equals(property))
        {
            try
            {
                if (value != null)
                    style.setBorderRightColor(value);
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal right border color: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
        else if (PROPERTY_TOP_BORDER_COLOR.equals(property))
        {
            try
            {
                if (value != null)
                    style.setBorderTopColor(value);
            }
            catch (IllegalArgumentException e)
            {
                logger.debug("Illegal top border color: {}.  IllegalArgumentException caught: ", value, e.getMessage());
            }
        }
    }

    /**
     * Returns the style map of style names to <code>Styles</code>.
     * @return A <code>Map</code> of style names to <code>Styles</code>.
     */
    public Map<String, Style> getStyleMap()
    {
        return myStyleMap;
    }
}
