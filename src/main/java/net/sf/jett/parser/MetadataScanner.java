package net.sf.jett.parser;

/**
 * A <code>MetadataScanner</code> object scans metadata text and returns tokens.
 *
 * @author Randy Gettman
 */
public class MetadataScanner
{
    /**
     * Enumeration for the different types of Tokens in Metadata.
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
        TOKEN_SEMICOLON(13),
        TOKEN_EQUALS(14),
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
    //private static final String PUNCT_CHARS_NOT_AS_STRING = "\"';=";
    private static final String PUNCT_CHARS_NOT_AS_STRING = "\";=";
    //private static final String PUNCT_CHARS_NOT_AS_STRING = ";=";

    private String myMetadataText;
    private int myOffset;
    private boolean amIInsideSingleQuotes;
    private boolean amIInsideDoubleQuotes;
    private String myCurrLexeme;

    /**
     * Construct a <code>MetadataScanner</code> object, with empty input.
     */
    public MetadataScanner()
    {
        this("");
    }

    /**
     * Construct a <code>MetadataScanner</code> object, with the given input.
     * @param metadataText The metadata text to scan.
     */
    public MetadataScanner(String metadataText)
    {
        setMetadataText(metadataText);
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
            if (iStartOfToken >= myMetadataText.length())
            {
                // EOI while in single quotes -- error!
                iTokenLength = 0;
                tokenType = Token.TOKEN_ERROR_EOI_IN_SQUOTES;
            }
            else if (myMetadataText.charAt(iStartOfToken) == '\'')
            {
                iTokenLength = 1;
                tokenType = Token.TOKEN_SINGLE_QUOTE;
                amIInsideSingleQuotes = false;
            }
            else
            {
                while ((iStartOfToken + iTokenLength) < myMetadataText.length() &&
                        myMetadataText.charAt(iStartOfToken + iTokenLength) != '\'')
                    iTokenLength++;
                tokenType = Token.TOKEN_STRING;
            }
        }
        else if (amIInsideDoubleQuotes)
        {
            if (iStartOfToken >= myMetadataText.length())
            {
                // EOI while in double quotes -- error!
                iTokenLength = 0;
                tokenType = Token.TOKEN_ERROR_EOI_IN_DQUOTES;
            }
            else if (myMetadataText.charAt(iStartOfToken) == '"')
            {
                iTokenLength = 1;
                tokenType = Token.TOKEN_DOUBLE_QUOTE;
                amIInsideDoubleQuotes = false;
            }
            else
            {
                while ((iStartOfToken + iTokenLength) < myMetadataText.length() &&
                        myMetadataText.charAt(iStartOfToken + iTokenLength) != '"')
                    iTokenLength++;
                tokenType = Token.TOKEN_STRING;
            }
        }
        else
        {
            // EOI test.
            if (iStartOfToken >= myMetadataText.length())
            {
                // End of input string.
                iTokenLength = 0;
                tokenType = Token.TOKEN_EOI;
            }
            // First char starts a string consisting of letters, numbers, and
            // all but a few punctuation characters.
            else if ((iStartOfToken + iTokenLength) < myMetadataText.length() &&
                    //!Character.isWhitespace(myMetadataText.charAt(iStartOfToken + iTokenLength)) &&
                    PUNCT_CHARS_NOT_AS_STRING.indexOf(myMetadataText.charAt(iStartOfToken + iTokenLength)) == -1)
            {
                // String mode.
                while ((iStartOfToken + iTokenLength) < myMetadataText.length() &&
                        //!Character.isWhitespace(myMetadataText.charAt(iStartOfToken + iTokenLength)) &&
                        PUNCT_CHARS_NOT_AS_STRING.indexOf(myMetadataText.charAt(iStartOfToken + iTokenLength)) == -1)
                {
                    iTokenLength++;
                }
                tokenType = Token.TOKEN_STRING;
            }
            else if (myMetadataText.charAt(iStartOfToken) == ';')
            {
                // Semicolon.
                iTokenLength = 1;
                tokenType = Token.TOKEN_SEMICOLON;
            }
            else if (myMetadataText.charAt(iStartOfToken) == '=')
            {
                // Equals.
                iTokenLength = 1;
                tokenType = Token.TOKEN_EQUALS;
            }
//         else if (myMetadataText.charAt(iStartOfToken) == '\'')
//         {
//            // Single Quote.
//            iTokenLength = 1;
//            tokenType = Token.TOKEN_SINGLE_QUOTE;
//            amIInsideSingleQuotes = true;
//         }
            else if (myMetadataText.charAt(iStartOfToken) == '"')
            {
                // Double Quote.
                iTokenLength = 1;
                tokenType = Token.TOKEN_DOUBLE_QUOTE;
                amIInsideDoubleQuotes = true;
            }
            else if (Character.isWhitespace(myMetadataText.charAt(iStartOfToken)))
            {
                // Whitespace.
                while ((iStartOfToken + iTokenLength) < myMetadataText.length() &&
                        Character.isWhitespace(myMetadataText.charAt(iStartOfToken + iTokenLength)))
                    iTokenLength++;
                tokenType = Token.TOKEN_WHITESPACE;
            }
        }  // End else from if (amIInsideDoubleQuotes)

        // Note down lexeme for access later.
        myCurrLexeme = myMetadataText.substring(iStartOfToken, iStartOfToken + iTokenLength);

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
     * Resets the scanner to the beginning of the metadata text.
     */
    public void reset()
    {
        myOffset = 0;
        amIInsideDoubleQuotes = false;
        amIInsideSingleQuotes = false;
        myCurrLexeme = null;
    }

    /**
     * Give the <code>MetadataScanner</code> another metadata text to scan.
     * Resets to the beginning of the string.
     * @param metadataText The metadata text to scan.
     */
    public void setMetadataText(String metadataText)
    {
        myMetadataText = metadataText;
        reset();
    }
}