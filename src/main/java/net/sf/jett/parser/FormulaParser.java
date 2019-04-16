package net.sf.jett.parser;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.ss.formula.SheetNameFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.exception.FormulaParseException;
import net.sf.jett.formula.CellRef;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>FormulaParser</code> parses formulas in formula text, extracting out
 * all cell references.  A cell reference consists of an optional sheet name
 * (optionally enclosed in single quotes) followed by an exclamation ("!"),
 * followed by a legal cell reference (alpha-number format), optionally
 * followed by a default value clause, which is two pipes followed by the
 * default value: "||value".
 *
 * @author Randy Gettman
 */
public class FormulaParser
{
    private static final Logger logger = LoggerFactory.getLogger(FormulaParser.class);

    private static final Pattern CELL_REF_PATTERN = Pattern.compile("\\$?[A-Za-z]+\\$?[1-9][0-9]*");

    private String myFormulaText;
    private List<CellRef> myCellReferences;
    private String mySheetName;
    private Cell myCell;
    private String myCellReference;
    private String myDefaultValue;
    private boolean amIInsideSingleQuotes;
    private boolean amIExpectingADefaultValue;

    /**
     * Create a <code>FormulaParser</code>.
     */
    public FormulaParser()
    {
        setFormulaText("");
    }

    /**
     * Create a <code>FormulaParser</code> object that will parse the given
     * formula text.
     * @param formulaText The text of the formula.
     */
    public FormulaParser(String formulaText)
    {
        setFormulaText(formulaText);
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
     * Sets the formula text to the given formula text and resets the parser.
     * @param formulaText The new formula text.
     */
    public void setFormulaText(String formulaText)
    {
        myFormulaText = formulaText;
        reset();
    }

    /**
     * Resets this <code>FormulaParser</code>, usually at creation time and
     * when new input arrives.
     */
    private void reset()
    {
        myCellReferences = new ArrayList<>();
        mySheetName = null;
        myCellReference = null;
        myDefaultValue = null;
        amIInsideSingleQuotes = false;
        amIExpectingADefaultValue = false;
    }

    /**
     * Parses the formula text.
     */
    public void parse()
    {
        FormulaScanner scanner = new FormulaScanner(myFormulaText);

        FormulaScanner.Token token = scanner.getNextToken();
        if (token == FormulaScanner.Token.TOKEN_WHITESPACE)
            token = scanner.getNextToken();

        // Parse any cell references found.
        while (token.getCode() >= 0 && token != FormulaScanner.Token.TOKEN_EOI)
        {
            switch(token)
            {
            case TOKEN_WHITESPACE:
                addCellReferenceIfFound();
                mySheetName = null;
                myCellReference = null;
                break;
            case TOKEN_STRING:
                if (amIExpectingADefaultValue)
                {
                    // Default value.
                    if (myDefaultValue == null)
                        myDefaultValue = scanner.getCurrLexeme();
                    else
                        myDefaultValue += scanner.getCurrLexeme();
                    amIExpectingADefaultValue = false;
                }
                else
                {
                    // For now, store it in the cell reference field.  Upon finding an
                    // exclamation, the value will be stored in the sheet name field.
                    myCellReference = myCellReference == null ? scanner.getCurrLexeme() : myCellReference + scanner.getCurrLexeme();
                }
                logger.debug("  parse: Token String: \"{}\".", scanner.getCurrLexeme());
                break;
            case TOKEN_EXCLAMATION:
                // If we had text from before the "!", then the text that's
                // currently in "myCellReference" is really the sheet reference.
                // Move it to the sheet name field.
                if (myCellReference == null)
                    throw new FormulaParseException("Sheet name delimiter (\"!\") found with no sheet name: " + myFormulaText
                            + SheetUtil.getCellLocation(myCell));
                if (amIExpectingADefaultValue)
                    throw new FormulaParseException("Sheet name delimiter (\"!\") found while expecting a default value: "
                            + myFormulaText + SheetUtil.getCellLocation(myCell));
                mySheetName = myCellReference;
                myCellReference = null;
                break;
            case TOKEN_LEFT_PAREN:
                // This can turn a potential cell reference into a function call!
                mySheetName = null;
                myCellReference = null;
                break;
            case TOKEN_OPERATOR:
                if (amIExpectingADefaultValue && scanner.getCurrLexeme().charAt(0) == '-')
                {
                    // Allow a "-" to indicate a negative default value.
                    myDefaultValue = "-";
                    break;
                }
                // Allow operators to "continue" a sheet name (currently stored in myCellReference).
                if (addCellReferenceIfFound())
                {
                    myCellReference = myCellReference + scanner.getCurrLexeme();
                }
                break;
            case TOKEN_RIGHT_PAREN:
            case TOKEN_COMMA:
            case TOKEN_DOUBLE_QUOTE:
                // Just delimiters between strings.  Validate the cell reference.
                addCellReferenceIfFound();
                mySheetName = null;
                myCellReference = null;
                break;
            case TOKEN_SINGLE_QUOTE:
                // Must keep track of whether a sheet reference occurs within single quotes.
                amIInsideSingleQuotes = !amIInsideSingleQuotes;
                break;
            case TOKEN_DOUBLE_PIPE:
                if (amIExpectingADefaultValue)
                    throw new FormulaParseException("Cannot have two default values for a cell reference: " + myFormulaText
                            + SheetUtil.getCellLocation(myCell));
                if (myCellReference == null)
                    throw new FormulaParseException("Default value indicator (\"||\") found without a cell reference: "
                            + myFormulaText + SheetUtil.getCellLocation(myCell));
                amIExpectingADefaultValue = true;
                break;
            default:
                throw new FormulaParseException("Parse error occurred: " + myFormulaText + SheetUtil.getCellLocation(myCell));
            }
            token = scanner.getNextToken();

            if (token == FormulaScanner.Token.TOKEN_EOI)
                break;
        }
        // Found end of input but something else was expected.
        if (token.getCode() < 0)
            throw new FormulaParseException("Found end of input while scanning formula text: " + myFormulaText + SheetUtil.getCellLocation(myCell));
        // Don't forget any last cell reference!
        addCellReferenceIfFound();
    }

    /**
     * If there is a valid cell reference, and it's not a duplicate, then add it
     * to the list.  Always null-out the cell reference, sheet name, and default
     * value.
     * @return Returns <code>true</code> if the sheet name is currently
     *    <code>null</code> and no cell reference is found.  This means that the
     *    string representing the current reference should be continued instead
     *    of discarded.  Else, <code>false</code>.
     * @since 0.2.0
     */
    private boolean addCellReferenceIfFound()
    {
        logger.trace("  aCRIF: Trying to match \"{}\".", myCellReference);
        if (myCellReference != null)
        {
            if (CELL_REF_PATTERN.matcher(myCellReference).matches())
            {
                CellRef ref;

                logger.trace("    aCRIF: Cell Reference is \"{}\".", myCellReference);
                if (mySheetName != null)
                    ref = new CellRef(SheetNameFormatter.format(mySheetName) + "!" + myCellReference);
                else
                    ref = new CellRef(myCellReference);
                if (myDefaultValue != null)
                {
                    logger.trace("    aCRIF: Default value found is \"{}\".", myDefaultValue);
                    ref.setDefaultValue(myDefaultValue);
                }

                logger.trace("    aCRIF: Cell Reference detected: {}", ref.formatAsString());
                // Don't add duplicates.
                if (!myCellReferences.contains(ref))
                {
                    logger.trace("      aCRIF: Not in list, adding ref: row={}, col={}, rowAbs={}, colAbs={}.",
                            ref.getRow(), ref.getCol(), ref.isRowAbsolute(), ref.isColAbsolute());
                    myCellReferences.add(ref);
                }
            }
            else if (mySheetName == null)
            {
                // Allow non-String tokens to be a part of the sheet name.
                // This allows sheet names constructed for implicit cloning
                // purposes to be recognized in JETT formulas, e.g.
                // $[SUM(${dvs.name}$@i=n;l=10;v=s;r=DNE!B3)]
                // Else the Excel Operator "=" will make this reference the
                // shortened "DNE!B3" erroneously.
                return true;
            }
        }
        mySheetName = null;
        myCellReference = null;
        myDefaultValue = null;
        return false;
    }

    /**
     * Returns a <code>List</code> of <code>CellRefs</code> that this parser
     * found in the formula text.
     * @return A <code>List</code> of <code>CellRefs</code>, possibly empty.
     */
    public List<CellRef> getCellReferences()
    {
        return myCellReferences;
    }
}