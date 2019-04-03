package net.sf.jett.exception;

/**
 * A <code>StyleParseException</code> occurs when there is an error parsing
 * CSS text.
 *
 * @author Randy Gettman
 * @since 0.5.0
 */
public class StyleParseException extends ParseException
{
    /**
     * Create a <code>StyleParseException</code>.
     */
    public StyleParseException()
    {
        super();
    }

    /**
     * Create a <code>StyleParseException</code> with the given message.
     * @param message The message.
     */
    public StyleParseException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>StyleParseException</code>.
     * @param cause The cause.
     */
    public StyleParseException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>StyleParseException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public StyleParseException(String message, Throwable cause)
    {
        super(message, cause);
    }
}