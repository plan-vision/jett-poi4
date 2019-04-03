package net.sf.jett.exception;
/**
 * A <code>TransformException</code> is a general-purpose exception for when
 * <code>RuntimeExceptions</code> occur during transformation.
 *
 * @author Randy Gettman
 * @since 0.11.0
 */
public class TransformException extends RuntimeException
{
    /**
     * Create a <code>TransformException</code>.
     */
    public TransformException()
    {
        super();
    }

    /**
     * Create a <code>TransformException</code> with the given message.
     * @param message The message.
     */
    public TransformException(String message)
    {
        super(message);
    }

    /**
     * Create a <code>TransformException</code>.
     * @param cause The cause.
     */
    public TransformException(Throwable cause)
    {
        super(cause);
    }

    /**
     * Create a <code>TransformException</code> with the given message.
     * @param message The message.
     * @param cause The cause.
     */
    public TransformException(String message, Throwable cause)
    {
        super(message, cause);
    }
}
