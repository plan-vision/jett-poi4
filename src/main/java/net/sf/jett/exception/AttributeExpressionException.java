package net.sf.jett.exception;

/**
 * <p>An <code>AttributeExpressionException</code> is a particular kind of
 * <code>TagParseException</code> that indicates that there was a problem
 * evaluating a JEXL Expression that was encountered as part of an attribute
 * value on a JETT tag.</p>
 *
 * <p>Usually JEXL has its own mechanism of either creating a warning log
 * message or returning <code>null</code>, but <code>nulls</code> are
 * problematic for <code>AttributeUtil</code>, which usually expects some kind
 * of value to result for an attribute of a JETT tag.</p>
 *
 * @author Randy Gettman
 * @since 0.7.0
 */
public class AttributeExpressionException extends TagParseException
{

    /**
     * Create a <code>AttributeExpressionException</code>.
     */
    public AttributeExpressionException()
    {
        super();
    }

    /**
     * Create a <code>AttributeExpressionException</code> with the given message.
     * @param message The message.
     */
    public AttributeExpressionException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>AttributeExpressionException</code>.
     * @param cause The cause.
     */
    public AttributeExpressionException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>AttributeExpressionException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public AttributeExpressionException(String message, Throwable cause)
    {
        super(message, cause);
    }
}
