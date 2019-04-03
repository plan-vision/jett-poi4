package net.sf.jett.exception;

/**
 * A <code>ParseException</code> occurs when there is an error parsing anything
 * related to JETT processing.
 *
 * @author Randy Gettman
 */
public class ParseException extends RuntimeException
{
    /**
     * Create a <code>ParseException</code>.
     */
    public ParseException()
    {
        super();
    }

    /**
     * Create a <code>ParseException</code> with the given message.
     * @param message The message.
     */
    public ParseException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>ParseException</code>.
     * @param cause The cause.
     */
    public ParseException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>ParseException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public ParseException(String message, Throwable cause)
    {
        super(message, cause);
    }
}