package net.sf.jett.parser;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.formula.Formula;
import net.sf.jett.util.RichTextStringUtil;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>TagParser</code> parses one JETT XML tag, either a begin tag or an
 * end tag, including the tag namespace (if any), the tag name, and any
 * attributes.
 *
 * @author Randy Gettman
 */
public class TagParser
{
    /**
     * Determines the beginning of an XML start tag.
     */
    public static final String BEGIN_START_TAG = "<";
    /**
     * Determines the beginning of an XML end tag.
     */
    public static final String BEGIN_END_TAG = "</";
    /**
     * Determines the ending of an XML start tag with a body.
     */
    public static final String END_TAG = ">";
    /**
     * Determines the ending of an XML start tag that is bodiless.
     */
    public static final String END_BODILESS_TAG = "/>";

    private Cell myCell;
    private String myCellText;
    private RichTextString myCellRichTextString;
    private int myStartIdx;
    private String myNamespace;
    private String myTagName;
    private boolean amIATag;
    private boolean amIEndTag;
    private boolean amIBodiless;
    private Map<String, RichTextString> myAttributes = new HashMap<String, RichTextString>();
    private int myTagStartIdx;
    private int myTagEndIdx;

    /**
     * Create a <code>TagParser</code> object that will parse the given tag text.
     * @param cell The <code>Cell</code> that contains text of the tag.
     */
    public TagParser(Cell cell)
    {
        this(cell, 0);
    }

    /**
     * Create a <code>TagParser</code> object that will parse the given tag
     * text, starting at the given position in the string..
     * @param cell The <code>Cell</code> that contains text of the tag.
     * @param startIdx The 0-based index into the string.
     */
    public TagParser(Cell cell, int startIdx)
    {
        myCell = cell;
        setCellText(cell.getStringCellValue().substring(startIdx));
        myStartIdx = startIdx;
        myCellRichTextString = cell.getRichStringCellValue();
    }

    /**
     * Sets the tag text to the given tag text and resets the parser.
     * @param tagText The new tag text.
     */
    public void setCellText(String tagText)
    {
        myCellText = tagText;
        reset();
    }

    /**
     * Resets this <code>TagParser</code>, usually at creation time and when new
     * input arrives.
     */
    private void reset()
    {
        myNamespace = null;
        myTagName = null;
        amIATag = false;
        amIEndTag = false;
        amIBodiless = false;
        myAttributes.clear();
        myTagStartIdx = -1;
        myTagEndIdx = -1;
    }

    /**
     * Parses the tag text.
     */
    public void parse()
    {
        TagScanner scanner = new TagScanner(myCellText);
        boolean insideJettFormula = false;

        // Tags must begin with "<" or "</", else it's not a tag (and it's not an error).
        // Text may occur before an ending tag or after a starting tag.
        TagScanner.Token token = scanner.getNextToken();
        // If we found a <, but it wasn't a tag, keep looking!
        while (!amIATag)
        {
            while ((token != TagScanner.Token.TOKEN_BEGIN_ANGLE_BRACKET &&
                    token != TagScanner.Token.TOKEN_BEGIN_ANGLE_BRACKET_SLASH &&
                    token != TagScanner.Token.TOKEN_EOI &&
                    token != TagScanner.Token.TOKEN_ERROR_EOI_IN_DQUOTES) ||
                    insideJettFormula)
            {
                String lexeme = scanner.getCurrLexeme();
                //System.err.println(lexeme);
                // Bypass any tokens normally indicating beginning of a tag if found
                // inside a JETT Formula.
                if (token == TagScanner.Token.TOKEN_STRING) {
                    if (lexeme.contains(Formula.BEGIN_FORMULA)) {
                        insideJettFormula = true;
                    }
                    if (lexeme.contains(Formula.END_FORMULA)) {
                        insideJettFormula = false;
                    }
                }

                // Prepare for next loop.
                token = scanner.getNextToken();
            }
            int begPos = scanner.getNextPosition();
            switch (token)
            {
            case TOKEN_BEGIN_ANGLE_BRACKET:
                // Start at the "<" position.
                myTagStartIdx = begPos - 1;
                amIEndTag = false;
                amIATag = true;
                break;
            case TOKEN_BEGIN_ANGLE_BRACKET_SLASH:
                // Start at the "</" position.
                myTagStartIdx = begPos - 2;
                amIEndTag = true;
                amIATag = true;
                break;
            default:
                myTagStartIdx = -1;
                myTagEndIdx = -1;
                amIATag = false;
                return;
            }

            // Extract possible namespace and tag name.
            token = scanner.getNextToken();
            // Not a tag: "<whitespace", "<=", "<<", "<>", "<\""
            // But "<:" is a bad tag, with no namespace.
            if (token != TagScanner.Token.TOKEN_STRING && token != TagScanner.Token.TOKEN_COLON)
            {
                //System.err.println("  \"<\" found but not a tag.  Continuing scan.");
                myTagStartIdx = -1;
                amIATag = false;
            }
        }

        // At this point, we know we have a tag, good or not.

        if (token == TagScanner.Token.TOKEN_STRING)
        {
            String lexeme = scanner.getCurrLexeme();
            token = scanner.getNextToken();
            if (token == TagScanner.Token.TOKEN_COLON)
            {
                token = scanner.getNextToken();
                if (token == TagScanner.Token.TOKEN_STRING)
                {
                    // namespace:tagName
                    myNamespace = lexeme;
                    myTagName = scanner.getCurrLexeme();
                    token = scanner.getNextToken();
                }
                else
                {
                    throw new TagParseException("Cannot find tag name in tag text: " + myCellText + SheetUtil.getCellLocation(myCell));
                }
            }
            else
            {
                // tagName
                myNamespace = "";
                myTagName = lexeme;
            }
        }
        else if (token == TagScanner.Token.TOKEN_COLON)
        {
            throw new TagParseException("Cannot find namespace in tag text: " + myCellText + SheetUtil.getCellLocation(myCell));
        }

        // Parse any attribute name/value pairs: attrName="value".
        String attrName = null;
        boolean insideDoubleQuotes = false;
        while (token.getCode() >= 0 && token != TagScanner.Token.TOKEN_END_ANGLE_BRACKET &&
                token != TagScanner.Token.TOKEN_SLASH_END_ANGLE_BRACKET)
        {
            switch(token)
            {
            case TOKEN_WHITESPACE:
                // Ignore.
                break;
            case TOKEN_STRING:
                if (insideDoubleQuotes)
                {
                    // Add newly complete attribute name/value pair.
                    if (attrName == null)
                        throw new TagParseException("Value found without attribute name: " + myCellText + SheetUtil.getCellLocation(myCell));
                    // Store the RichTextString attribute value.
                    int pos = myStartIdx + scanner.getNextPosition();
                    CreationHelper helper = myCell.getSheet().getWorkbook().getCreationHelper();
                    RichTextString attrValue = RichTextStringUtil.substring(myCellRichTextString,
                            helper, pos - scanner.getCurrLexeme().length(), pos);
                    // Replace _all_ tabs, carriage returns, linefeeds with spaces.
                    attrValue = RichTextStringUtil.replaceValues(attrValue, helper,
                            Arrays.asList("\n", "\r", "\t"),
                            Arrays.asList(" " , " " , " " ),
                            true);
                    // Perform escape-sequence replacement.
                    attrValue = RichTextStringUtil.performEscaping(attrValue, helper);
                    myAttributes.put(attrName, attrValue);
                    attrName = null;
                }
                else
                    attrName = scanner.getCurrLexeme();
                break;
            case TOKEN_EQUALS:
                if (attrName == null)
                    throw new TagParseException("Attribute name missing before \"=\": " + myCellText + SheetUtil.getCellLocation(myCell));
                break;
            case TOKEN_COLON:
                throw new TagParseException("Colon not allowed in attribute name: " + myCellText + SheetUtil.getCellLocation(myCell));
            case TOKEN_DOUBLE_QUOTE:
                insideDoubleQuotes = !insideDoubleQuotes;
                break;
            case TOKEN_BEGIN_ANGLE_BRACKET:
            case TOKEN_BEGIN_ANGLE_BRACKET_SLASH:
                throw new TagParseException("Cannot start a tag within another tag: " + myCellText + SheetUtil.getCellLocation(myCell));
            case TOKEN_EOI:
                throw new TagParseException("Tags must start with \"" + BEGIN_START_TAG + "\" or \"" +
                        BEGIN_END_TAG + "\" and end with \"" + END_TAG + "\" or \"" + END_BODILESS_TAG +
                        "\": " + myCellText + " at " + SheetUtil.getCellLocation(myCell));
            default:
                throw new TagParseException("Parse error occurred: " + myCellText + SheetUtil.getCellLocation(myCell));
            }
            token = scanner.getNextToken();
        }
        // Found end angle bracket before attribute value found.
        if (attrName != null)
            throw new TagParseException("Found end of tag before attribute value: " + myCellText + SheetUtil.getCellLocation(myCell));
        if (token.getCode() < 0)
            throw new TagParseException("Found end of input while scanning attribute value: " + myCellText + SheetUtil.getCellLocation(myCell));

        // If "/>", then the tag is bodiless, else (">") there is a body.
        amIBodiless = (token == TagScanner.Token.TOKEN_SLASH_END_ANGLE_BRACKET);

        // We have reached the end angle bracket.  Bodiless tags cannot have tag
        // text before or after the tag.
        myTagEndIdx = scanner.getNextPosition();
    }

    /**
     * Returns whether the given tag text is in fact a tag.  That is, if the tag
     * text starts with <code>BEGIN_START_TAG</code> or
     * <code>BEGIN_END_TAG</code> and ends with <code>END_TAG</code>.
     * @return <code>true</code> if the tag text represents a tag,
     *    <code>false</code> otherwise.
     * @see #BEGIN_START_TAG
     * @see #BEGIN_END_TAG
     * @see #END_TAG
     */
    public boolean isTag()
    {
        return amIATag;
    }

    /**
     * Returns whether this tag is the end of the tag or not.  That is, if the
     * tag text starts with <code>BEGIN_END_TAG</code>.
     * @return <code>true</code> if the tag text represents an end tag,
     *    <code>false</code> if the tag text represents a start tag.
     * @see #BEGIN_START_TAG
     * @see #BEGIN_END_TAG
     */
    public boolean isEndTag()
    {
        return amIEndTag;
    }

    /**
     * Returns whether this tag is bodiless.  That is, if the
     * tag text ends with <code>END_BODILESS_TAG</code>.
     * @return <code>true</code> if the tag text represents an end tag,
     *    <code>false</code> if the tag text represents a start tag.
     * @see #END_TAG
     * @see #END_BODILESS_TAG
     */
    public boolean isBodiless()
    {
        return amIBodiless;
    }

    /**
     * Returns the namespace found, if any.  That is, the text before the colon
     * in the tag name.  E.g. <code>&lt;namespace:tagname ...&gt;</code>
     * @return The namespace, or <code>null</code> if missing.
     */
    public String getNamespace()
    {
        return myNamespace;
    }

    /**
     * Returns the tag name found, if any.  That is, the text after the colon in
     * the tag name, or the whole tag name if no colon is found.  E.g.
     * <code>&lt;namespace:tagname ...&gt;</code> or <code>&lt;tagname ...&gt;</code>.
     * @return The tag name.
     */
    public String getTagName()
    {
        return myTagName;
    }

    /**
     * Returns a formatted string containing the namespace, followed by a colon
     * (if the  namespace exists), followed by the tag name, e.g.
     * <code>getNamespace() + ":" + getTagName()</code>.
     * @return A formatted string.
     */
    public String getNamespaceAndTagName()
    {
        if (myNamespace != null && myNamespace.length() > 0)
            return myNamespace + ":" + myTagName;
        else
            return myTagName;
    }

    /**
     * Returns a <code>Map</code> of attribute names mapped to attribute values,
     * possibly empty.
     * E.g.<code>&lt;namespace:tagname attr1="value1" attr2="value2"&gt;</code>
     * is returned as <code>["attr1"=&gt;"value1", "attr2"=&gt;"value2"]</code>.
     * @return A <code>Map</code> of attribute names and attribute values.
     */
    public Map<String, RichTextString> getAttributes()
    {
        return myAttributes;
    }

    /**
     * Returns the <code>Cell</code> whose tag text is being parsed.
     * @return The <code>Cell</code>.
     */
    public Cell getCell()
    {
        return myCell;
    }

    /**
     * Returns the portion of the cell text that is the tag text.
     * @return The portion of the cell text that is the tag text.
     */
    public String getTagText()
    {
        if (myTagStartIdx != -1 && myTagEndIdx != -1)
            return myCellText.substring(myTagStartIdx, myTagEndIdx);
        return null;
    }

    /**
     * Returns the 0-based index into the cell text that is after the tag.
     * @return The 0-based index into the cell text that is after the tag.
     */
    public int getAfterTagIdx()
    {
        return myTagEndIdx;
    }
}

