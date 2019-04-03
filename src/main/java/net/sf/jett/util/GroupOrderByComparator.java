package net.sf.jett.util;

import java.util.Comparator;
import java.util.List;

import net.sf.jett.exception.ParseException;
import net.sf.jett.model.Group;

/**
 * A <code>GroupOrderByComparator</code> is an <code>OrderByComparator</code>
 * that operates on groups and understands that some of the "order by"
 * properties are group properties (these must all be before the other "order
 * by" properties that aren't group properties).
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class GroupOrderByComparator<T extends Group> implements Comparator<T>
{
    private OrderByComparator<Object> myOrderByComparator;
    private List<String> myGroupByProperties;

    /**
     * Constructs a <code>GroupOrderByComparator</code> that operates on the
     * given <code>List</code> of "order by" expressions, and assumes that the
     * <code>Groups</code> are grouped by the given <code>List</code> of "group
     * by" expressions.
     * @param orderByComparator An <code>OrderByComparator</code>.
     * @param groupByProperties A <code>List</code> of "group by" properties.
     * @throws ParseException If there is a problem parsing the expressions, or
     *    if any "group by" properties are preceded in the "order by" properties
     *    by other properties that are not "group by" properties.
     */
    public GroupOrderByComparator(OrderByComparator<Object> orderByComparator, List<String> groupByProperties)
    {
        myOrderByComparator = orderByComparator;
        myGroupByProperties = groupByProperties;
        ensureOrderByGroupByLegal();
    }

    /**
     * Ensures that no "order by" properties are before any "group by"
     * properties in the expressions list.
     * @throws ParseException If any "group by" properties are preceded in the
     *    "order by" properties by other properties that are not "group by"
     *    properties.
     */
    private void ensureOrderByGroupByLegal()
    {
        boolean foundOrderByNotInGroupBy = false;
        boolean orderByInGroupBy;
        List<String> orderByProperties = myOrderByComparator.getProperties();
        for (String orderBy : orderByProperties)
        {
            orderByInGroupBy = myGroupByProperties.contains(orderBy);
            if (foundOrderByNotInGroupBy)
            {
                if (!orderByInGroupBy)
                    throw new ParseException("The \"order by\" property \"" + orderBy +
                            "\" is in the \"group by\" properties (" + myGroupByProperties.toString() +
                            "), but it must not follow other \"order by\" properties (" +
                            orderByProperties.toString() + ") that aren't \"group by\" properties.");
            }
            else
            {
                // Set flag indicating that all further "order by" properties must
                // NOT be a "group by" property.
                if (!orderByInGroupBy)
                    foundOrderByNotInGroupBy = true;
            }
        }
    }

    /**
     * <p>Compares the given <code>Groups</code> to determine order.  Fulfills
     * the <code>Comparator</code> contract by returning a negative integer, 0,
     * or a positive integer if <code>o1</code> is less than, equal to, or
     * greater than <code>o2</code>.</p>
     * <p>This compare method compares the group's representative objects.</p>
     *
     * @param g1 The left-hand-side <code>Group</code> to compare.
     * @param g2 The right-hand-side <code>Group</code> to compare.
     * @return A negative integer, 0, or a positive integer if <code>g1</code>
     *    is less than, equal to, or greater than <code>g2</code>.
     * @throws UnsupportedOperationException If any property specified in the
     *    constructor doesn't correspond to a no-argument "get&lt;Property&gt;"
     *    getter method in <code>T</code>, or if the property's type is not
     *    <code>Comparable</code>.
     */
    @Override
    public int compare(T g1, T g2)
    {
        return myOrderByComparator.compare(g1.getObj(), g2.getObj());
    }

    /**
     * Returns the <code>List</code> of "group by" properties.
     * @return The <code>List</code> of "group by" properties.
     */
    public List<String> getGroupByProperties()
    {
        return myGroupByProperties;
    }
}
