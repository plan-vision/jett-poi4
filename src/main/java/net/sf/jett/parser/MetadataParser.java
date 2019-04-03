package net.sf.jett.parser;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;

import net.sf.jett.exception.MetadataParseException;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>MetadataParser</code> parses metadata at the end of cell text.
 *
 * @author Randy Gettman
 */
public class MetadataParser
{
    /**
     * Metadata variable name for extra rows in the Block.
     */
    public static final String VAR_NAME_EXTRA_ROWS = "extraRows";
    /**
     * Metadata variable name for additional columns to the left in the Block.
     */
    public static final String VAR_NAME_LEFT = "left";
    /**
     * Metadata variable name for additional columns to the right in the Block.
     */
    public static final String VAR_NAME_RIGHT = "right";
    /**
     * Metadata variable name for copying right instead of down.  This is
     * ignored if neither "left" nor "right" is specified.
     */
    public static final String VAR_NAME_COPY_RIGHT = "copyRight";
    /**
     * Metadata variable name for not shifting other content out of the way of a
     * looping block.  This turns on the implicit version of the "fixed"
     * attribute of looping tags.
     */
    public static final String VAR_NAME_FIXED = "fixed";
    /**
     * Metadata variable name specifying the "past end action" to take whenever
     * a <code>Collection</code> is exhausted before the end of iteration.
     * @since 0.2.0
     */
    public static final String VAR_NAME_PAST_END_ACTION = "pastEndAction";
    /**
     * Metadata variable name specifying the value with which to replace any
     * expression that refers to a <code>Collection</code> that is exhausted
     * before the end of iteration.  This defaults to an empty string
     * <code>""</code>.  This is only relevant if the past end action is to
     * replace expressions.
     * @since 0.7.0
     */
    public static final String VAR_NAME_REPLACE_VALUE = "replaceValue";
    /**
     * Metadata variable name specifying to create an Excel grouping for rows,
     * columns, or no grouping.
     * @since 0.2.0
     */
    public static final String VAR_NAME_GROUP_DIR = "groupDir";
    /**
     * Metadata variable name specifying whether any Excel grouping created
     * should be collapsed, defaulting to <code>false</code>.
     * @since 0.2.0
     */
    public static final String VAR_NAME_COLLAPSE = "collapse";
    /**
     * Metadata variable name specifying a <code>TagLoopListener</code> to
     * listen for <code>TagLoopEvents</code>.
     * @since 0.3.0
     */
    public static final String VAR_NAME_ON_LOOP_PROCESSED = "onLoopProcessed";
    /**
     * Metadata variable name specifying a <code>TagListener</code> to listen
     * for <code>TagEvents</code>.
     * @since 0.3.0
     */
    public static final String VAR_NAME_ON_PROCESSED = "onProcessed";
    /**
     * Metadata variable name specifying the name of the zero-based "looping"
     * variable.
     * @since 0.3.0
     */
    public static final String VAR_NAME_INDEXVAR = "indexVar";
    /**
     * Metadata variable name specifying a limit to the number of iterations
     * processed.
     * @since 0.3.0
     */
    public static final String VAR_NAME_LIMIT = "limit";
    /**
     * Metadata variable name specifying the name of the
     * {@link net.sf.jett.model.LoopTagStatus} variable.
     * @since 0.9.1
     */
    public static final String VAR_NAME_VAR_STATUS = "varStatus";

    private Cell myCell;
    private boolean amIExpectingAValue;
    private String myMetadataText;
    private String myExtraRows;
    private String myColsLeft;
    private String myColsRight;
    private boolean amIDefiningCols;
    private String myCopyingRight;
    private String myFixed;
    private String myPastEndActionValue;
    private String myReplacementValue;
    private String myGroupDir;
    private String myCollapsingGroup;
    private String myTagLoopListener;
    private String myTagListener;
    private String myIndexVarName;
    private String myLimit;
    private String myVarStatusName;

    /**
     * Create a <code>MetadataParser</code>.
     */
    public MetadataParser()
    {
        setMetadataText("");
    }

    /**
     * Create a <code>MetadataParser</code> object that will parse the given
     * metadata text.
     * @param metadataText The text of the metadata.
     */
    public MetadataParser(String metadataText)
    {
        setMetadataText(metadataText);
    }

    /**
     * Sets the <code>Cell</code> that contains the formula to be parsed.
     * @param cell The <code>Cell</code>.
     * @since 0.7.0
     */
    public void setCell(Cell cell)
    {
        myCell = cell;
    }

    /**
     * Sets the metadata text to the given metadata text and resets the parser.
     * @param metadataText The new metadata text.
     */
    public void setMetadataText(String metadataText)
    {
        myMetadataText = metadataText;
        reset();
    }

    /**
     * Resets this <code>MetadataParser</code>, usually at creation time and
     * when new input arrives.
     */
    private void reset()
    {
        amIExpectingAValue = false;
        myExtraRows = null;
        myColsLeft = null;
        myColsRight = null;
        amIDefiningCols = false;
        myCopyingRight = null;
        myFixed = null;
        myPastEndActionValue = null;
        myReplacementValue = "";  // default to empty string
        myGroupDir = null;
        myCollapsingGroup = null;
        myTagLoopListener = null;
        myTagListener = null;
        myIndexVarName = null;
        myLimit = null;
    }

    /**
     * Parses the metadata text.
     */
    public void parse()
    {
        MetadataScanner scanner = new MetadataScanner(myMetadataText);
        Map<String, String> abbr = getAbbreviations();

        MetadataScanner.Token token = scanner.getNextToken();
        if (token == MetadataScanner.Token.TOKEN_WHITESPACE)
            token = scanner.getNextToken();

        // Parse any metadata variable name/value pairs: varName=value.
        String varName = null;
        while (token.getCode() >= 0 && token != MetadataScanner.Token.TOKEN_EOI)
        {
            switch(token)
            {
            case TOKEN_WHITESPACE:
                // Ignore.
                break;
            case TOKEN_STRING:
                if (amIExpectingAValue)
                {
                    String lexeme = scanner.getCurrLexeme();
                    if (abbr != null && abbr.containsKey(varName))
                    {
                        varName = abbr.get(varName);
                    }

                    // Add newly complete variable name/value pair.
                    if (VAR_NAME_EXTRA_ROWS.equals(varName))
                    {
                        myExtraRows = lexeme;
                    }
                    else if (VAR_NAME_LEFT.equals(varName))
                    {
                        myColsLeft = lexeme;
                        amIDefiningCols = true;
                    }
                    else if (VAR_NAME_RIGHT.equals(varName))
                    {
                        myColsRight = lexeme;
                        amIDefiningCols = true;
                    }
                    else if (VAR_NAME_COPY_RIGHT.equals(varName))
                    {
                        myCopyingRight = lexeme;
                    }
                    else if (VAR_NAME_FIXED.equals(varName))
                    {
                        myFixed = lexeme;
                    }
                    else if (VAR_NAME_PAST_END_ACTION.equals(varName))
                    {
                        myPastEndActionValue = lexeme;
                    }
                    else if (VAR_NAME_REPLACE_VALUE.equals(varName))
                    {
                        myReplacementValue = lexeme;
                    }
                    else if (VAR_NAME_GROUP_DIR.equals(varName))
                    {
                        myGroupDir = lexeme;
                    }
                    else if (VAR_NAME_COLLAPSE.equals(varName))
                    {
                        myCollapsingGroup = lexeme;
                    }
                    else if (VAR_NAME_ON_LOOP_PROCESSED.equals(varName))
                    {
                        myTagLoopListener = lexeme;
                    }
                    else if (VAR_NAME_ON_PROCESSED.equals(varName))
                    {
                        myTagListener = lexeme;
                    }
                    else if (VAR_NAME_INDEXVAR.equals(varName))
                    {
                        myIndexVarName = lexeme;
                    }
                    else if (VAR_NAME_LIMIT.equals(varName))
                    {
                        myLimit = lexeme;
                    }
                    else if (VAR_NAME_VAR_STATUS.equals(varName))
                    {
                        myVarStatusName = lexeme;
                    }
                    else
                    {
                        throw new MetadataParseException("Unrecognized variable name: \"" +
                                varName + "\"." + SheetUtil.getCellLocation(myCell));
                    }
                    if (isRestricted(varName))
                    {
                        throw new MetadataParseException("Variable name \"" + varName +
                                "\" is restricted in this context.");
                    }
                    varName = null;
                    amIExpectingAValue = false;
                }
                else
                    varName = scanner.getCurrLexeme();
                break;
            case TOKEN_EQUALS:
                if (varName == null)
                    throw new MetadataParseException("Variable name missing before \"=\": " + myMetadataText +
                            SheetUtil.getCellLocation(myCell));
                amIExpectingAValue = true;
                break;
            case TOKEN_SEMICOLON:
                // Just a delimiter between var/value pairs.
                break;
            case TOKEN_DOUBLE_QUOTE:
            case TOKEN_SINGLE_QUOTE:
                // These just delimit Strings.
                break;
            default:
                throw new MetadataParseException("Parse error occurred: " + myMetadataText + SheetUtil.getCellLocation(myCell));
            }
            token = scanner.getNextToken();
        }
        // Found end of input before attribute value found.
        if (varName != null)
            throw new MetadataParseException("Found end of metadata before equals sign at \"" +
                    varName + "\": " + myMetadataText);
        if (amIExpectingAValue)
            throw new MetadataParseException("Found end of metadata before variable value: " + myMetadataText +
                    SheetUtil.getCellLocation(myCell));
        if (token.getCode() < 0)
            throw new MetadataParseException("Found end of input while scanning metadata value: " + myMetadataText +
                    SheetUtil.getCellLocation(myCell));
    }

    /**
     * Determines whether the given metadata key is <em>restricted</em> -- not
     * allowed in this context.  Subclasses can override this method to
     * implement their own restricted lists.  This implementation doesn't
     * restrict anything - it always returns <code>false</code>.
     * @param metadataKey The metadata key.
     * @return <code>true</code> if it's restricted, else <code>false</code>.
     * @since 0.9.1
     */
    protected boolean isRestricted(String metadataKey)
    {
        return false;
    }

    /**
     * Returns a <code>Map</code> of abbreviations.  The map's keys are the
     * abbreviations, and the values are the metadata keys.
     * @return A <code>Map</code> of abbreviations, <code>null</code> means no
     *    abbreviations are permitted.
     * @since 0.9.1
     */
    protected Map<String, String> getAbbreviations()
    {
        return null;
    }

    /**
     * Returns the "extra rows" lexeme.
     * @return The "extra rows" lexeme.
     */
    public String getExtraRows()
    {
        return myExtraRows;
    }

    /**
     * Returns the "columns left" lexeme.
     * @return The "columns left" lexeme.
     */
    public String getColsLeft()
    {
        return myColsLeft;
    }

    /**
     * Returns the "columns right" lexeme.
     * @return The "columns right" lexeme.
     */
    public String getColsRight()
    {
        return myColsRight;
    }

    /**
     * Returns whether column definitions are present.
     * @return Whether column definitions are present.
     */
    public boolean isDefiningCols()
    {
        return amIDefiningCols;
    }

    /**
     * Returns the "copy right" lexeme.
     * @return The "copy right" lexeme.
     */
    public String getCopyingRight()
    {
        return myCopyingRight;
    }

    /**
     * Returns the "fixed" lexeme.
     * @return The "fixed" lexeme.
     */
    public String getFixed()
    {
        return myFixed;
    }

    /**
     * Returns the "past end action" lexeme.
     * @return The "past end action" lexeme.
     * @since 0.2.0
     */
    public String getPastEndAction()
    {
        return myPastEndActionValue;
    }

    /**
     * Returns the replacement value, which is used to replace expressions that
     * have been exhausted prior to the end of iteration, if the past end action
     * is to replace expressions.
     * @return The replacement value, which defaults to an empty string
     *    <code>""</code>.
     * @since 0.7.0
     */
    public String getReplacementValue()
    {
        return myReplacementValue;
    }

    /**
     * Returns the "group dir" lexeme.
     * @return The "group dir" lexeme.
     * @since 0.2.0
     */
    public String getGroupDir()
    {
        return myGroupDir;
    }

    /**
     * Returns the "collapse" lexeme.
     * @return The "collapse" lexeme.
     * @since 0.2.0
     */
    public String getCollapsingGroup()
    {
        return myCollapsingGroup;
    }

    /**
     * Returns the "tag loop listener" lexeme.
     * @return The "tag loop listener" lexeme.
     * @since 0.3.0
     */
    public String getTagLoopListener()
    {
        return myTagLoopListener;
    }

    /**
     * Returns the "tag listener" lexeme.
     * @return The "tag listener" lexeme.
     * @since 0.3.0
     */
    public String getTagListener()
    {
        return myTagListener;
    }

    /**
     * Returns the "looping" variable name.
     * @return The "looping" variable name.
     * @since 0.3.0
     */
    public String getIndexVarName()
    {
        return myIndexVarName;
    }

    /**
     * Returns the "limit" lexeme.
     * @return The "limit" lexeme.
     * @since 0.3.0
     */
    public String getLimit()
    {
        return myLimit;
    }

    /**
     * Returns the "varStatus" lexeme.
     * @return The "varStatus" lexeme.
     * @since 0.9.1
     */
    public String getVarStatusName()
    {
        return myVarStatusName;
    }
}