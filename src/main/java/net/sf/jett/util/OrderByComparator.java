package net.sf.jett.util;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import net.sf.jagg.exception.JaggException;
import net.sf.jagg.util.MethodCache;
import net.sf.jett.exception.ParseException;

/**
 * <p>An <code>OrderByComparator</code> is a <code>Comparator</code> that is
 * capable of comparing two objects based on a dynamic list of properties of
 * the objects of type <code>T</code>.  It can sort any of its properties
 * ascending or descending, and for any of its properties, it can place nulls
 * first or last.  Like SQL, this will default to ascending.  Nulls default to
 * last if ascending, and first if descending.</p>
 * <p>This is based on jAgg's <code>PropertiesComparator</code>, which as of
 * the time of creation of this class always does ascending, nulls last.</p>
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class OrderByComparator<T> implements Comparator<T>
{
    /**
     * Sort ascending (default).
     */
    public static final String ASC = "ASC";
    /**
     * Sort descending.
     */
    public static final String DESC = "DESC";
    /**
     * Use this to indicate which sequence nulls should be ordered.
     */
    public static final String NULLS = "NULLS";
    /**
     * Sort nulls first (default if descending order).
     */
    public static final String FIRST = "FIRST";
    /**
     * Sort nulls last (default if ascending order).
     */
    public static final String LAST = "LAST";

    /**
     * Constant to order ascending.
     */
    public static final int ORDER_ASC = 1;
    /**
     * Constant to order descending.
     */
    public static final int ORDER_DESC = -1;
    /**
     * Constant to order nulls last.
     */
    public static final int NULLS_LAST = 1;
    /**
     * Constant to order nulls first.
     */
    public static final int NULLS_FIRST = -1;

    private List<String> myProperties;
    private List<Integer> myOrderings;
    private List<Integer> myNullOrderings;
    private int mySize;

    /**
     * Constructs an <code>OrderByComparator</code> based on a <code>List</code>
     * of expressions, of the format "property [ASC|DESC] [NULLS FIRST|LAST]".
     * @param expressions A <code>List</code> of expressions.
     * @throws ParseException If there is a problem parsing the expressions.
     */
    public OrderByComparator(List<String> expressions)
    {
        setExpressions(expressions);
    }

    /**
     * Sets the internal lists for all properties, order sequences, and null
     * order sequences.
     * @param expressions A <code>List</code> of expressions.
     * @throws ParseException If there is a problem parsing the expressions.
     */
    private void setExpressions(List<String> expressions)
    {
        if (expressions == null || expressions.size() <= 0)
            throw new ParseException("No order by expressions found.");
        mySize = expressions.size();
        myProperties = new ArrayList<>(mySize);
        myOrderings = new ArrayList<>(mySize);
        myNullOrderings = new ArrayList<>(mySize);
        for (String expr : expressions)
        {
            String[] parts = expr.split("\\s+");
            String property;
            int ordering;
            int nullOrdering;
            if (parts.length > 0 && parts.length < 5)
            {
                property = parts[0];
                ordering = ORDER_ASC;
                nullOrdering = NULLS_LAST;

                if (parts.length == 2 || parts.length == 4)
                {
                    // ordering is next.
                    if (ASC.equalsIgnoreCase(parts[1]))
                    {
                        ordering = ORDER_ASC;
                        nullOrdering = NULLS_LAST;
                    }
                    else if (DESC.equalsIgnoreCase(parts[1]))
                    {
                        ordering = ORDER_DESC;
                        nullOrdering = NULLS_FIRST;
                    }
                    else
                        throw new ParseException("Expected \"" + ASC + "\" or \"" + DESC + ": " + expr);
                }
                if (parts.length == 3 || parts.length == 4)
                {
                    if (!NULLS.equalsIgnoreCase(parts[parts.length - 2]))
                        throw new ParseException("Expected \"" + NULLS + " " + FIRST + "|" + LAST + ": " + expr);
                    if (LAST.equalsIgnoreCase(parts[parts.length - 1]))
                        nullOrdering = NULLS_LAST;
                    else if (FIRST.equalsIgnoreCase(parts[parts.length - 1]))
                        nullOrdering = NULLS_FIRST;
                    else
                        throw new ParseException("Expected \"" + FIRST + "\" or \"" + LAST + ": " + expr);
                }
            }
            else
            {
                throw new ParseException("Expected \"property\" [" + ASC + "|" + DESC + "] [" + NULLS + " ]" +
                        FIRST + "|" + LAST + ": " + expr);
            }

            myProperties.add(property);
            myOrderings.add(ordering);
            myNullOrderings.add(nullOrdering);
        }
    }

    /**
     * <p>Compares the given objects to determine order.  Fulfills the
     * <code>Comparator</code> contract by returning a negative integer, 0, or a
     * positive integer if <code>o1</code> is less than, equal to, or greater
     * than <code>o2</code>.</p>
     * <p>This compare method respects all properties, their order sequences,
     * and their null order sequences.</p>
     *
     * @param o1 The left-hand-side object to compare.
     * @param o2 The right-hand-side object to compare.
     * @return A negative integer, 0, or a positive integer if <code>o1</code>
     *    is less than, equal to, or greater than <code>o2</code>.
     * @throws UnsupportedOperationException If any property specified in the
     *    constructor doesn't correspond to a no-argument "get&lt;Property&gt;"
     *    getter method in <code>T</code>, or if the property's type is not
     *    <code>Comparable</code>.
     */
    @Override
    @SuppressWarnings("unchecked")
    public int compare(T o1, T o2) throws UnsupportedOperationException
    {
        int comp;
        for (int i = 0; i < mySize; i++)
        {
            String property = myProperties.get(i);
            int ordering = myOrderings.get(i);
            int nullOrdering = myNullOrderings.get(i);

            Comparable value1, value2;
            // This had to be copied from Aggregator.java, because Aggregator's
            // static method "getValueFromProperty" is protected.
            // Otherwise, we could call "Aggregator.getValueFromProperty", which
            // wraps all of the checked exceptions in an
            // UnsupportedOperationException.
            MethodCache cache = MethodCache.getMethodCache();
            try
            {
                value1 = (Comparable)
                        cache.getValueFromProperty(o1, property);
                value2 = (Comparable)
                        cache.getValueFromProperty(o2, property);
            }
            catch (JaggException e)
            {
                throw new UnsupportedOperationException("No matching method found for \"" +
                        property + "\".", e);
            }
            try
            {
                if (value1 == null)
                {
                    if (value2 == null)
                        comp = 0;
                    else
                        comp = nullOrdering;
                }
                else
                {
                    if (value2 == null)
                        comp = -nullOrdering;
                    else
                        comp = ordering * value1.compareTo(value2);
                }
                if (comp != 0) return comp;
            }
            catch (ClassCastException e)
            {
                throw new UnsupportedOperationException("Property \"" + property + "\" needs to be Comparable.");
            }
        }
        return 0;
    }

    /**
     * Indicates whether the given <code>OrderByComparator</code> is equal to
     * this <code>OrderByComparator</code>.  All property names must match in
     * order, and all of the order sequences and null order sequences must
     * match.
     *
     * @param obj The other <code>OrderByComparator</code>.
     */
    @Override
    public boolean equals(Object obj)
    {
        if (obj instanceof OrderByComparator)
        {
            OrderByComparator otherComp = (OrderByComparator) obj;
            if (mySize != otherComp.mySize)
                return false;
            for (int i = 0; i < mySize; i++)
            {
                if (!myProperties.get(i).equals(otherComp.myProperties.get(i)))
                    return false;
                if (myOrderings.get(i) != otherComp.myOrderings.get(i))
                    return false;
                if (myNullOrderings.get(i) != otherComp.myNullOrderings.get(i))
                    return false;
            }
            return true;
        }
        return false;
    }

    /**
     * Returns a <code>List</code> of all properties.
     * @return A <code>List</code> of all properties.
     */
    public List<String> getProperties()
    {
        return myProperties;
    }

    /**
     * Returns a <code>List</code> of orderings.
     * @return A <code>List</code> of orderings.
     * @see #ORDER_ASC
     * @see #ORDER_DESC
     */
    public List<Integer> getOrderings()
    {
        return myOrderings;
    }

    /**
     * Returns a <code>List</code> of null orderings.
     * @return A <code>List</code> of null orderings.
     * @see #NULLS_FIRST
     * @see #NULLS_LAST
     */
    public List<Integer> getNullOrderings()
    {
        return myNullOrderings;
    }
}
