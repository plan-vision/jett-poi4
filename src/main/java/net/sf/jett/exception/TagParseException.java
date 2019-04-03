package net.sf.jett.exception;

/**
 * A <code>TagParseException</code> indicates when there is an error parsing
 * the tag text.
 *
 * @author Randy Gettman
 */
public class TagParseException extends ParseException
{
    /**
     * Create a <code>TagParseException</code>.
     */
    public TagParseException()
    {
        super();
    }

    /**
     * Create a <code>TagParseException</code> with the given message.
     * @param message The message.
     */
    public TagParseException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>TagParseException</code>.
     * @param cause The cause.
     */
    public TagParseException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>TagParseException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public TagParseException(String message, Throwable cause)
    {
        super(message, cause);
    }
}

