package net.sf.jett.exception;

/**
 * A <code>MetadataParseException</code> occurs when there is an error parsing
 * the metadata text.
 *
 * @author Randy Gettman
 */
public class MetadataParseException extends ParseException
{

    /**
     * Create a <code>MetadataParseException</code>.
     */
    public MetadataParseException()
    {
        super();
    }

    /**
     * Create a <code>MetadataParseException</code> with the given message.
     * @param message The message.
     */
    public MetadataParseException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>MetadataParseException</code>.
     * @param cause The cause.
     */
    public MetadataParseException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>MetadataParseException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public MetadataParseException(String message, Throwable cause)
    {
        super(message, cause);
    }
}
