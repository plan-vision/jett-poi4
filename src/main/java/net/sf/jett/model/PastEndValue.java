package net.sf.jett.model;

/**
 * A <code>PastEndValue</code> represents the fact that an expression
 * references a collection item beyond the end of the iteration.  This is
 * distinct from <code>null</code>, which may be a legitimate value.  This is
 * closer to <code>Void</code>, but that can't be instantiated.
 *
 * @author Randy Gettman
 */
public class PastEndValue
{
    /**
     * The singleton <code>PastEndValue</code> value.
     */
    public static final PastEndValue PAST_END_VALUE = new PastEndValue();

    /**
     * Don't instantiate directly.
     */
    private PastEndValue() {}
}
