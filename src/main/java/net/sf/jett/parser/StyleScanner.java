package net.sf.jett.parser;

/**
 * A <code>StyleScanner</code> scans CSS text and returns tokens.
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class StyleScanner
{
    /**
     * Enumeration for the different types of Tokens in "CSS".
     */
    public enum Token
    {
        TOKEN_ERROR_EOI_IN_COMMENT(-3),
        TOKEN_ERROR_BUF_NULL(-2),
        TOKEN_UNKNOWN(-1),
        TOKEN_WHITESPACE(0),
        TOKEN_STRING(1),
        TOKEN_COLON(11),
        TOKEN_PERIOD(12),
        TOKEN_BEGIN_BRACE(13),
        TOKEN_END_BRACE(14),
        TOKEN_SEMICOLON(15),
        TOKEN_COMMENT(98),
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
    private static final String PUNCT_CHARS_NOT_AS_STRING = ":.{};/*";

    private String myCssText;
    private int myOffset;
    private String myCurrLexeme;

    /**
     * Construct a <code>StyleScanner</code> object, with empty input.
     */
    public StyleScanner()
    {
        this("");
    }

    /**
     * Construct a <code>StyleScanner</code> object, with the given input.
     * @param cssText The CSS text to scan.
     */
    public StyleScanner(String cssText)
    {
        setCssText(cssText);
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

        // EOI test.
        if (iStartOfToken >= myCssText.length())
        {
            // End of input string.
            return Token.TOKEN_EOI;
        }
        if (myCssText.charAt(iStartOfToken) == '/' &&
                (iStartOfToken + 1) < myCssText.length() && myCssText.charAt(iStartOfToken + 1) == '*')
        {
            // Comment.
            // Skip everything until "*/" found, or error if not found.
            iTokenLength += 2;
            boolean endOfCommentFound = false;
            while ((iStartOfToken + iTokenLength) < myCssText.length())
            {
                if (myCssText.charAt(iStartOfToken + iTokenLength) == '*' &&
                        (iStartOfToken + iTokenLength + 1 < myCssText.length()) && myCssText.charAt(iStartOfToken + iTokenLength + 1) == '/')
                {
                    iTokenLength += 2;
                    endOfCommentFound = true;
                    break;
                }
                iTokenLength++;
            }
            if (!endOfCommentFound)
            {
                myCurrLexeme = null;
                return Token.TOKEN_ERROR_EOI_IN_COMMENT;
            }
            myOffset += iTokenLength;
            iStartOfToken = myOffset;
            iTokenLength = 0;
        }
        // First char starts a string consisting of letters, numbers, and
        // all but a few punctuation characters.
        if ((iStartOfToken + iTokenLength) < myCssText.length() &&
                !Character.isWhitespace(myCssText.charAt(iStartOfToken + iTokenLength)) &&
                PUNCT_CHARS_NOT_AS_STRING.indexOf(myCssText.charAt(iStartOfToken + iTokenLength)) == -1)
        {
            // String mode.
            while ((iStartOfToken + iTokenLength) < myCssText.length() &&
                    !Character.isWhitespace(myCssText.charAt(iStartOfToken + iTokenLength)) &&
                    PUNCT_CHARS_NOT_AS_STRING.indexOf(myCssText.charAt(iStartOfToken + iTokenLength)) == -1)
            {
                iTokenLength++;
            }
            tokenType = Token.TOKEN_STRING;
        }
        else if (myCssText.charAt(iStartOfToken) == ':')
        {
            // Colon.
            iTokenLength = 1;
            tokenType = Token.TOKEN_COLON;
        }
        else if (myCssText.charAt(iStartOfToken) == '.')
        {
            // Period.
            iTokenLength = 1;
            tokenType = Token.TOKEN_PERIOD;
        }
        else if (myCssText.charAt(iStartOfToken) == '}')
        {
            // End brace.
            iTokenLength = 1;
            tokenType = Token.TOKEN_END_BRACE;
        }
        else if (myCssText.charAt(iStartOfToken) == '{')
        {
            // Begin brace.
            iTokenLength = 1;
            tokenType = Token.TOKEN_BEGIN_BRACE;
        }
        else if (myCssText.charAt(iStartOfToken) == ';')
        {
            // Semicolon.
            iTokenLength = 1;
            tokenType = Token.TOKEN_SEMICOLON;
        }
        else if (Character.isWhitespace(myCssText.charAt(iStartOfToken)))
        {
            // Whitespace.
            while ((iStartOfToken + iTokenLength) < myCssText.length() &&
                    Character.isWhitespace(myCssText.charAt(iStartOfToken + iTokenLength)))
                iTokenLength++;
            tokenType = Token.TOKEN_WHITESPACE;
        }

        // Note down lexeme for access later.
        myCurrLexeme = myCssText.substring(iStartOfToken, iStartOfToken + iTokenLength);

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
     * Resets the scanner to the beginning of the CSS text.
     */
    public void reset()
    {
        myOffset = 0;
        myCurrLexeme = null;
    }

    /**
     * Give the <code>StyleScanner</code> another CSS text to scan.
     * Resets to the beginning of the string.
     * @param cssText The css Text to scan.
     */
    public void setCssText(String cssText)
    {
        myCssText = cssText;
        reset();
    }
}
