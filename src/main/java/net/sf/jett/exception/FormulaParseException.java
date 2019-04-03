package net.sf.jett.exception;

/**
 * A <code>FormulaParseException</code> occurs when there is an error parsing
 * formula text.
 *
 * @author Randy Gettman
 */
public class FormulaParseException extends ParseException
{

    /**
     * Create a <code>FormulaParseException</code>.
     */
    public FormulaParseException()
    {
        super();
    }

    /**
     * Create a <code>FormulaParseException</code> with the given message.
     * @param message The message.
     */
    public FormulaParseException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>FormulaParseException</code>.
     * @param cause The cause.
     */
    public FormulaParseException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>FormulaParseException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public FormulaParseException(String message, Throwable cause)
    {
        super(message, cause);
    }
}
