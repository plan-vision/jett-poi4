package net.sf.jett.parser;

/**
 * A <code>TagScanner</code> scans tag XML text and returns tokens.
 *
 * @author Randy Gettman
 */
public class TagScanner
{
    /**
     * Enumeration for the different types of Tokens in Tags.
     */
    public enum Token
    {
        TOKEN_ERROR_EOI_IN_DQUOTES(-3),
        TOKEN_ERROR_BUF_NULL(-2),
        TOKEN_UNKNOWN(-1),
        TOKEN_WHITESPACE(0),
        TOKEN_STRING(1),
        TOKEN_COLON(11),
        TOKEN_DOUBLE_QUOTE(12),
        TOKEN_BEGIN_ANGLE_BRACKET(13),
        TOKEN_BEGIN_ANGLE_BRACKET_SLASH(14),
        TOKEN_END_ANGLE_BRACKET(15),
        TOKEN_SLASH_END_ANGLE_BRACKET(16),
        TOKEN_EQUALS(17),
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
    private static final String PUNCT_CHARS_NOT_AS_STRING = "\":=<>/";

    private String myTagText;
    private int myOffset;
    private boolean amIInsideDoubleQuotes;
    private String myCurrLexeme;

    /**
     * Construct a <code>TagScanner</code> object, with empty input.
     */
    public TagScanner()
    {
        this("");
    }

    /**
     * Construct a <code>TagScanner</code> object, with the given input.
     * @param tagText The tag text to scan.
     */
    public TagScanner(String tagText)
    {
        setTagText(tagText);
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
        if (amIInsideDoubleQuotes)
        {
            if (iStartOfToken >= myTagText.length())
            {
                // EOI while in double quotes -- error!
                iTokenLength = 0;
                tokenType = Token.TOKEN_ERROR_EOI_IN_DQUOTES;
            }
            else if (myTagText.charAt(iStartOfToken) == '"')
            {
                iTokenLength = 1;
                tokenType = Token.TOKEN_DOUBLE_QUOTE;
                amIInsideDoubleQuotes = false;
            }
            else
            {
                while ((iStartOfToken + iTokenLength) < myTagText.length() &&
                        myTagText.charAt(iStartOfToken + iTokenLength) != '"')
                {
                    // Include escaped characters.

                    // \' -> '
                    // \n -> (new line)
                    // \t => (tab)
                    // \r -> (carriage return)
                    // \b -> (backspace)
                    // \f -> (form feed)
                    if (myTagText.charAt(iStartOfToken + iTokenLength) == '\\' &&
                            (iStartOfToken + iTokenLength + 1) < myTagText.length() &&
                            ("\"\\'ntrbf".indexOf(myTagText.charAt(iStartOfToken + iTokenLength + 1)) >= 0))
                    {
                        iTokenLength += 2;
                    }
                    else
                    {
                        iTokenLength++;
                    }
                }
                tokenType = Token.TOKEN_STRING;
            }
        }
        else
        {
            // EOI test.
            if (iStartOfToken >= myTagText.length())
            {
                // End of input string.
                iTokenLength = 0;
                tokenType = Token.TOKEN_EOI;
            }
            // First char starts a string consisting of letters, numbers, and
            // all but a few punctuation characters.
            else if ((iStartOfToken + iTokenLength) < myTagText.length() &&
                    !Character.isWhitespace(myTagText.charAt(iStartOfToken + iTokenLength)) &&
                    PUNCT_CHARS_NOT_AS_STRING.indexOf(myTagText.charAt(iStartOfToken + iTokenLength)) == -1)
            {
                // String mode.
                while ((iStartOfToken + iTokenLength) < myTagText.length() &&
                        !Character.isWhitespace(myTagText.charAt(iStartOfToken + iTokenLength)) &&
                        PUNCT_CHARS_NOT_AS_STRING.indexOf(myTagText.charAt(iStartOfToken + iTokenLength)) == -1)
                {
                    iTokenLength++;
                }
                tokenType = Token.TOKEN_STRING;
            }
            else if (myTagText.charAt(iStartOfToken) == ':')
            {
                // Colon.
                iTokenLength = 1;
                tokenType = Token.TOKEN_COLON;
            }
            else if (myTagText.charAt(iStartOfToken) == '=')
            {
                // Equals.
                iTokenLength = 1;
                tokenType = Token.TOKEN_EQUALS;
            }
            else if (myTagText.charAt(iStartOfToken) == '>')
            {
                // End angle bracket.
                iTokenLength = 1;
                tokenType = Token.TOKEN_END_ANGLE_BRACKET;
            }
            else if (myTagText.charAt(iStartOfToken) == '<')
            {
                // Begin angle bracket.
                if (iStartOfToken + 1 < myTagText.length() && myTagText.charAt(iStartOfToken + 1) == '/')
                {
                    // Begin angle bracket and slash.
                    tokenType = Token.TOKEN_BEGIN_ANGLE_BRACKET_SLASH;
                    iTokenLength = 2;
                }
                else
                {
                    // Just begin angle bracket.
                    iTokenLength = 1;
                    tokenType = Token.TOKEN_BEGIN_ANGLE_BRACKET;
                }
            }
            else if (myTagText.charAt(iStartOfToken) == '/')
            {
                // Slash
                if (iStartOfToken + 1 < myTagText.length() && myTagText.charAt(iStartOfToken + 1) == '>')
                {
                    // Slash and end angle bracket.
                    tokenType = Token.TOKEN_SLASH_END_ANGLE_BRACKET;
                    iTokenLength = 2;
                }
                else
                {
                    // Can't have slash by itself.  This will cause a Parser error.
                    tokenType = Token.TOKEN_UNKNOWN;
                    iTokenLength = 1;
                }
            }
            else if (myTagText.charAt(iStartOfToken) == '"')
            {
                // Double Quote.
                iTokenLength = 1;
                tokenType = Token.TOKEN_DOUBLE_QUOTE;
                amIInsideDoubleQuotes = true;
            }
            else if (Character.isWhitespace(myTagText.charAt(iStartOfToken)))
            {
                // Whitespace.
                while ((iStartOfToken + iTokenLength) < myTagText.length() &&
                        Character.isWhitespace(myTagText.charAt(iStartOfToken + iTokenLength)))
                    iTokenLength++;
                tokenType = Token.TOKEN_WHITESPACE;
            }
        }  // End else from if (amIInsideDoubleQuotes)

        // Note down lexeme for access later.
        myCurrLexeme = myTagText.substring(iStartOfToken, iStartOfToken + iTokenLength);

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
     * Returns the current position of the next token.
     * @return The current position of the next token.
     */
    public int getNextPosition()
    {
        return myOffset;
    }

    /**
     * Resets the scanner to the beginning of the tag text.
     */
    public void reset()
    {
        myOffset = 0;
        amIInsideDoubleQuotes = false;
        myCurrLexeme = null;
    }

    /**
     * Give the <code>TagScanner</code> another tag text to scan.
     * Resets to the beginning of the string.
     * @param tagText The tagText to scan.
     */
    public void setTagText(String tagText)
    {
        myTagText = tagText;
        reset();
    }
}
