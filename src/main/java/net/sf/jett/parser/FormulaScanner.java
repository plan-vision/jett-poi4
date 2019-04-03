package net.sf.jett.parser;

/**
 * A <code>FormulaScanner</code> scans formula text and returns tokens.
 *
 * @author Randy Gettman
 */
public class FormulaScanner
{
    /**
     * Enumeration for the different types of Tokens in Formulas.
     */
    public enum Token
    {
        TOKEN_ERROR_EOI_IN_SQUOTES(-4),
        TOKEN_ERROR_EOI_IN_DQUOTES(-3),
        TOKEN_ERROR_BUF_NULL(-2),
        TOKEN_UNKNOWN(-1),
        TOKEN_WHITESPACE(0),
        TOKEN_STRING(1),
        TOKEN_SINGLE_QUOTE(11),
        TOKEN_DOUBLE_QUOTE(12),
        TOKEN_EXCLAMATION(13),
        TOKEN_LEFT_PAREN(14),
        TOKEN_RIGHT_PAREN(15),
        TOKEN_COMMA(16),
        TOKEN_DOUBLE_PIPE(17),
        TOKEN_OPERATOR(18),
        TOKEN_EOI(99);

        private int myCode;

        // Create a token with a code.
        private Token(int code)
        {
            myCode = code;
        }

        /**
         * Returns the unique code associated with this <code>Token</code>.
         * @return The unique code.
         */
        public int getCode()
        {
            return myCode;
        }
    }

    /**
     * <p>! for sheet!cellRef</p>
     * <p>() for function calls</p>
     * <p>, for parameter separation</p>
     * <p>=<>&+*-/^%: for Excel operators</p>
     */
    private static final String PUNCT_CHARS_NOT_AS_STRING = "'!(),=<>&+*-/^%:|";

    private String myFormulaText;
    private int myOffset;
    private boolean amIInsideSingleQuotes;
    private boolean amIInsideDoubleQuotes;
    private String myCurrLexeme;

    /**
     * Construct a <code>FormulaScanner</code> object, with empty input.
     */
    public FormulaScanner()
    {
        this("");
    }

    /**
     * Construct a <code>FormulaScanner</code> object, with the given input.
     * @param formulaText The formula text to scan.
     */
    public FormulaScanner(String formulaText)
    {
        setFormulaText(formulaText);
    }

    /**
     * Returns the <code>Token</code>.  After this call completes, the current
     * lexeme is available via a call to <code>getCurrLexeme</code>.
     * Starts looking at the current offset, and once the token is found, then
     * the offset is advanced to the start of the next token.
     * @return A <code>Token</code>.
     * @see #getCurrLexeme
     */
    public Token getNextToken()
    {
        int iStartOfToken = myOffset;
        int iTokenLength = 0;
        Token tokenType = Token.TOKEN_UNKNOWN;

        // Inside single-quotes, the whole thing until EOI or another single-quote
        // is one string!
        if (amIInsideSingleQuotes)
        {
            if (iStartOfToken >= myFormulaText.length())
            {
                // EOI while in single quotes -- error!
                iTokenLength = 0;
                tokenType = Token.TOKEN_ERROR_EOI_IN_SQUOTES;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '\'')
            {
                iTokenLength = 1;
                tokenType = Token.TOKEN_SINGLE_QUOTE;
                amIInsideSingleQuotes = false;
            }
            else
            {
                while ((iStartOfToken + iTokenLength) < myFormulaText.length() &&
                        myFormulaText.charAt(iStartOfToken + iTokenLength) != '\'')
                    iTokenLength++;
                tokenType = Token.TOKEN_STRING;
            }
        }
        else if (amIInsideDoubleQuotes)
        {
            if (iStartOfToken >= myFormulaText.length())
            {
                // EOI while in single quotes -- error!
                iTokenLength = 0;
                tokenType = Token.TOKEN_ERROR_EOI_IN_DQUOTES;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '"')
            {
                iTokenLength = 1;
                tokenType = Token.TOKEN_DOUBLE_QUOTE;
                amIInsideDoubleQuotes = false;
            }
            else
            {
                while ((iStartOfToken + iTokenLength) < myFormulaText.length() &&
                        myFormulaText.charAt(iStartOfToken + iTokenLength) != '"')
                    iTokenLength++;
                tokenType = Token.TOKEN_STRING;
            }
        }
        else
        {
            // EOI test.
            if (iStartOfToken >= myFormulaText.length())
            {
                // End of input string.
                iTokenLength = 0;
                tokenType = Token.TOKEN_EOI;
            }
            // First char starts a string consisting of letters, numbers, and
            // all but a few punctuation characters.
            else if ((iStartOfToken + iTokenLength) < myFormulaText.length() &&
                    !Character.isWhitespace(myFormulaText.charAt(iStartOfToken + iTokenLength)) &&
                    PUNCT_CHARS_NOT_AS_STRING.indexOf(myFormulaText.charAt(iStartOfToken + iTokenLength)) == -1)
            {
                // String mode.
                while ((iStartOfToken + iTokenLength) < myFormulaText.length() &&
                        !Character.isWhitespace(myFormulaText.charAt(iStartOfToken + iTokenLength)) &&
                        PUNCT_CHARS_NOT_AS_STRING.indexOf(myFormulaText.charAt(iStartOfToken + iTokenLength)) == -1)
                {
                    iTokenLength++;
                }
                tokenType = Token.TOKEN_STRING;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '!')
            {
                // Exclamation.
                iTokenLength = 1;
                tokenType = Token.TOKEN_EXCLAMATION;
            }
            else if (myFormulaText.charAt(iStartOfToken) == ',')
            {
                // Comma.
                iTokenLength = 1;
                tokenType = Token.TOKEN_COMMA;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '\'')
            {
                // Single Quote.
                iTokenLength = 1;
                tokenType = Token.TOKEN_SINGLE_QUOTE;
                amIInsideSingleQuotes = true;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '"')
            {
                // Double Quote.
                iTokenLength = 1;
                tokenType = Token.TOKEN_DOUBLE_QUOTE;
                amIInsideDoubleQuotes = true;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '(')
            {
                // Left Paren.
                iTokenLength = 1;
                tokenType = Token.TOKEN_LEFT_PAREN;
            }
            else if (myFormulaText.charAt(iStartOfToken) == ')')
            {
                // Right Paren.
                iTokenLength = 1;
                tokenType = Token.TOKEN_RIGHT_PAREN;
            }
            else if (myFormulaText.charAt(iStartOfToken) == '|')
            {
                // Pipe.
                if (iStartOfToken + 1 < myFormulaText.length() &&
                        myFormulaText.charAt(iStartOfToken + 1) == '|')
                {
                    // Double pipe for default value.
                    iTokenLength = 2;
                    tokenType = Token.TOKEN_DOUBLE_PIPE;
                }
                else
                {
                    // Just treat a single pipe as if it were an operator.
                    iTokenLength = 1;
                    tokenType = Token.TOKEN_OPERATOR;
                }
            }
            else if (myFormulaText.charAt(iStartOfToken) == '=' ||
                    myFormulaText.charAt(iStartOfToken) == '<' ||
                    myFormulaText.charAt(iStartOfToken) == '>' ||
                    myFormulaText.charAt(iStartOfToken) == '&' ||
                    myFormulaText.charAt(iStartOfToken) == '+' ||
                    myFormulaText.charAt(iStartOfToken) == '*' ||
                    myFormulaText.charAt(iStartOfToken) == '-' ||
                    myFormulaText.charAt(iStartOfToken) == '/' ||
                    myFormulaText.charAt(iStartOfToken) == '^' ||
                    myFormulaText.charAt(iStartOfToken) == '%' ||
                    myFormulaText.charAt(iStartOfToken) == ':'
                    )
            {
                // Excel Operators
                iTokenLength = 1;
                tokenType = Token.TOKEN_OPERATOR;
            }
            else if (Character.isWhitespace(myFormulaText.charAt(iStartOfToken)))
            {
                // Whitespace.
                while ((iStartOfToken + iTokenLength) < myFormulaText.length() &&
                        Character.isWhitespace(myFormulaText.charAt(iStartOfToken + iTokenLength)))
                    iTokenLength++;
                tokenType = Token.TOKEN_WHITESPACE;
            }
        }  // End else from if (amIInsideDoubleQuotes)

        // Note down lexeme for access later.
        myCurrLexeme = myFormulaText.substring(iStartOfToken, iStartOfToken + iTokenLength);

        // Update the offset.
        myOffset += iTokenLength;

        return tokenType;
    }

    /**
     * Returns the current lexeme after a call to <code>getNextToken</code>.
     * @return The current lexeme, or <code>null</code> if
     *    <code>getNextToken</code> hasn't been called yet after a reset.
     * @see #getNextToken
     * @see #reset
     */
    public String getCurrLexeme()
    {
        return myCurrLexeme;
    }

    /**
     * Resets the scanner to the beginning of the formula text.
     */
    public void reset()
    {
        myOffset = 0;
        amIInsideSingleQuotes = false;
        amIInsideDoubleQuotes = false;
        myCurrLexeme = null;
    }

    /**
     * Give the <code>FormulaScanner</code> another formula text to scan.
     * Resets to the beginning of the string.
     * @param formulaText The formula text to scan.
     */
    public void setFormulaText(String formulaText)
    {
        myFormulaText = formulaText;
        reset();
    }
}