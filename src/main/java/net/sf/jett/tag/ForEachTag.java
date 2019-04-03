package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jagg.AggregateFunction;
import net.sf.jagg.Aggregations;
import net.sf.jagg.Aggregator;
import net.sf.jagg.CollectAggregator;
import net.sf.jagg.model.AggregateValue;
import net.sf.jett.exception.TagParseException;
import net.sf.jett.expression.Expression;
import net.sf.jett.model.Block;
import net.sf.jett.model.Group;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.GroupOrderByComparator;
import net.sf.jett.util.OrderByComparator;

/**
 * <p>A <code>ForEachTag</code> represents a repetitively placed
 * <code>Block</code> of <code>Cells</code>, with each repetition corresponding
 * to an element of a <code>Collection</code>.  The <code>var</code> attribute
 * represents the variable name by which a collection item is known. The
 * optional <code>indexVar</code> attribute is the name of the variable that
 * holds the iterator index.  The optional <code>where</code> attribute filters
 * the collection by the given condition.  The optional <code>limit</code>
 * attribute specifies a limit to the number of iterations to be run from the
 * collection.  If the limit is greater than the number of items in the
 * collection, then blank blocks will result, with the exact result dependent
 * on "past end action" rules.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li><em>Inherits all attributes from {@link BaseLoopTag}.</em></li>
 * <li>items (required): <code>Collection</code></li>
 * <li>var (required): <code>String</code></li>
 * <li>indexVar (optional): <code>String</code></li>
 * <li>where (optional): <code>boolean</code></li>
 * <li>limit (optional): <code>int</code></li>
 * <li>groupBy (optional): <code>List&lt;String&gt;</code></li>
 * <li>orderBy (optional): <code>List&lt;String&gt;</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class ForEachTag extends BaseLoopTag
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Attribute for specifying the <code>Collection</code> over which to
     * iterate.
     */
    public static final String ATTR_ITEMS = "items";
    /**
     * Attribute for specifying the "looping variable" name.
     */
    public static final String ATTR_VAR = "var";
    /**
     * Attribute for specifying the name of the variable to be exposed that
     * indicates the 0-based index position into the <code>Collection</code>.
     */
    public static final String ATTR_INDEXVAR = "indexVar";
    /**
     * Attribute for specifying the condition that filters the
     * <code>Collection</code> contents before display.
     */
    public static final String ATTR_WHERE = "where";
    /**
     * Attribute for specifying the number of iterations to be displayed.
     */
    public static final String ATTR_LIMIT = "limit";
    /**
     * Attribute for specifying the property or properties by which to group the
     * <code>Collection</code> items, if any.
     * @since 0.3.0
     */
    public static final String ATTR_GROUP_BY = "groupBy";
    /**
     * Attribute for specifying the property or properties by which to order the
     * <code>Collection</code> items, if any.
     * @since 0.3.0
     */
    public static final String ATTR_ORDER_BY = "orderBy";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ITEMS, ATTR_VAR));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(
                    ATTR_INDEXVAR, ATTR_WHERE, ATTR_LIMIT, ATTR_GROUP_BY, ATTR_ORDER_BY));

    private Collection<Object> myCollection = null;
    private String myCollectionName = null;
    private String myVarName = null;
    private String myIndexVarName = null;
    private int myLimit = 0;
    private List<String> myGroupByProperties;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "forEach";
    }

    /**
     * Returns the <code>List</code> of required attribute names.
     * @return The <code>List</code> of required attribute names.
     */
    @Override
    public List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = new ArrayList<>(super.getRequiredAttributes());
        reqAttrs.addAll(REQ_ATTRS);
        return reqAttrs;
    }

    /**
     * Returns the <code>List</code> of optional attribute names.
     * @return The <code>List</code> of optional attribute names.
     */
    @Override
    public List<String> getOptionalAttributes()
    {
        List<String> optAttrs = new ArrayList<>(super.getOptionalAttributes());
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Validates the attributes for this <code>Tag</code>.  The "items"
     * attribute must be a <code>Collection</code>.  The "limit", if present,
     * must be a non-negative integer.
     */
    @Override
    @SuppressWarnings("unchecked")
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (isBodiless())
            throw new TagParseException("ForEach tags must have a body.  Bodiless ForEach tag found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();
        myCollection = AttributeUtil.evaluateObject(this, attributes.get(ATTR_ITEMS), beans, ATTR_ITEMS, Collection.class,
                new ArrayList<>(0));

        // Collection name.
        String attrItems = attributes.get(ATTR_ITEMS).getString();
        int beginExprIdx = attrItems.indexOf(Expression.BEGIN_EXPR);
        int endExprIdx = attrItems.indexOf(Expression.END_EXPR);
        if (beginExprIdx != -1 && endExprIdx != -1 && endExprIdx > beginExprIdx)
        {
            myCollectionName = attrItems.substring(beginExprIdx +
                    Expression.BEGIN_EXPR.length(), endExprIdx);
        }

        logger.debug("Collection \"{}\" has size {}", attrItems, myCollection.size());

        myVarName = AttributeUtil.evaluateString(this, attributes.get(ATTR_VAR), beans, null);

        myIndexVarName = AttributeUtil.evaluateString(this, attributes.get(ATTR_INDEXVAR), beans, null);

        RichTextString rtsCondition = attributes.get(ATTR_WHERE);
        if (rtsCondition != null)
        {
            // Create a new Collection containing only those items where the given
            // condition is true.
            ArrayList<Object> newCollection = new ArrayList<>();
            for (Object item : myCollection)
            {
                beans.put(myVarName, item);
                boolean condition = AttributeUtil.evaluateBoolean(this, rtsCondition, beans, true);
                if (condition)
                {
                    newCollection.add(item);
                }
            }
            beans.remove(myVarName);
            myCollection = newCollection;
        }

        List<String> orderByProperties = AttributeUtil.evaluateList(this, attributes.get(ATTR_ORDER_BY), beans, new ArrayList<String>(0));
        OrderByComparator<Object> comp = null;
        if (!orderByProperties.isEmpty())
        {
            comp = new OrderByComparator<>(orderByProperties);
            sortTheCollection(comp);
        }

        myGroupByProperties = AttributeUtil.evaluateList(this, attributes.get(ATTR_GROUP_BY), beans, new ArrayList<String>(0));
        if (!myGroupByProperties.isEmpty())
        {
            List<Group> groups = groupTheCollection();
            if (!orderByProperties.isEmpty())
            {
                sortTheGroups(groups, comp);
            }
            myCollection = new ArrayList<Object>(groups);
        }

        myLimit = AttributeUtil.evaluateNonNegativeInt(this, attributes.get(ATTR_LIMIT), beans, ATTR_LIMIT, myCollection.size());

        logger.debug("vA: myLimit={}", myLimit);
    }

    /**
     * Returns the names of the <code>Collections</code> that are being used in
     * this <code>ForEachTag</code>.
     * @return A <code>List</code> of one collection name.
     */
    @Override
    protected List<String> getCollectionNames()
    {
        return Arrays.asList(myCollectionName);
    }

    /**
     * Returns the name of the variable that is being used in this
     * <code>ForEachTag</code>.
     * @return A <code>List</code> of one variable name.
     * @since 0.7.0
     */
    @Override
    protected List<String> getVarNames()
    {
        return Arrays.asList(myVarName);
    }

    /**
     * Returns the number of iterations.
     * @return The number of iterations.
     */
    @Override
    protected int getNumIterations()
    {
        return myLimit;
    }

    /**
     * Returns the size of the collection being iterated.
     * @return The size of the collection being iterated.
     */
    @Override
    protected int getCollectionSize()
    {
        return myCollection.size();
    }

    /**
     * Returns an <code>Iterator</code> that iterates over some
     * <code>Collection</code> of objects.
     * @return An <code>Iterator</code>.
     */
    @Override
    protected Iterator<Object> getLoopIterator()
    {
        return new ForEachTagIterator();
    }

    /**
     * Place the <code>Iterator</code> item into the <code>Map</code> of beans.
     *
     * @param context The <code>TagContext</code>.
     * @param currBlock The <code>Block</code> that is about to processed.
     * @param item The <code>Object</code> that resulted from the iterator.
     * @param index The iteration index (0-based).
     */
    @Override
    protected void beforeBlockProcessed(TagContext context, Block currBlock, Object item, int index)
    {
        Map<String, Object> beans = context.getBeans();
        beans.put(myVarName, item);

        logger.debug("beforeBP: index={}", index);

        // Optional index counter variable.
        if (myIndexVarName != null && myIndexVarName.length() > 0)
            beans.put(myIndexVarName, index);
    }

    /**
     * Remove the <code>Iterator</code> item from the <code>Map</code> of beans.
     *
     * @param context The <code>TagContext</code>.
     * @param index The iteration index (0-based).
     * @param item The <code>Object</code> that resulted from the iterator.
     * @param currBlock The <code>Block</code> that was just processed.
     */
    @Override
    protected void afterBlockProcessed(TagContext context, Block currBlock, Object item, int index)
    {
        Map<String, Object> beans = context.getBeans();
        beans.remove(myVarName);

        // Optional index counter variable.
        if (myIndexVarName != null && myIndexVarName.length() > 0)
            beans.remove(myIndexVarName);
    }

    /**
     * Use an <code>OrderByComparator</code> to sort the collection of objects
     * by the "order by" properties.  It will sort it in place if it's a
     * <code>List</code>, otherwise it will make a copy of the list, sort it,
     * and assign it to the collection.
     * @param comp An <code>OrderByComparator</code>.
     */
    private void sortTheCollection(OrderByComparator<Object> comp)
    {
        if (myCollection instanceof List)
        {
            Collections.sort((List<Object>) myCollection, comp);
        }
        else
        {
            List<Object> temp = new ArrayList<>(myCollection);
            Collections.sort(temp, comp);
            myCollection = temp;
        }
    }

    /**
     * Create and use a <code>GroupOrderByComparator</code> to sort the groups.
     * @param groups A <code>List</code> of <code>Groups</code>.
     * @param comp An <code>OrderByComparator</code>.
     */
    private void sortTheGroups(List<Group> groups, OrderByComparator<Object> comp)
    {
        GroupOrderByComparator<Group> gComp = new GroupOrderByComparator<>(comp, myGroupByProperties);
        Collections.sort(groups, gComp);
    }

    /**
     * Use a <code>CollectAggregator</code> to partition the collection of
     * objects by the "group by" properties into <code>Groups</code>.  When
     * complete, this method will have replaced all items in the collection with
     * <code>Groups</code> of items.
     * @return A <code>List</code> of <code>Groups</code>.
     */
    private List<Group> groupTheCollection()
    {
        List<Object> items = new ArrayList<>(myCollection);
        List<AggregateFunction> aggregators = Arrays.<AggregateFunction>asList(new CollectAggregator(Aggregator.PROP_SELF));
        List<AggregateValue<Object>> aggValues = Aggregations.groupBy(items, myGroupByProperties, aggregators);
        List<Group> groups = new ArrayList<>(aggValues.size());
        for (AggregateValue aggValue : aggValues)
        {
            Group g = new Group();
            g.setItems((List<?>) aggValue.getAggregateValue(0));
            g.setObj(aggValue.getObject());
            groups.add(g);
        }
        return groups;
    }

    /**
     * The <code>Iterator</code> over the collection items, which can be
     * extended by a large limit to return <code>nulls</code> beyond the limit
     * of the <code>Collection</code>.
     */
    private class ForEachTagIterator implements Iterator<Object>
    {
        private int myIndex;
        private Iterator<Object> myInternalIterator;

        /**
         * Construct a <code>ForEachTagIterator</code> whose index is initialized
         * to zero.
         */
        private ForEachTagIterator()
        {
            myIndex = 0;
            myInternalIterator = myCollection.iterator();
        }

        /**
         * It doesn't make sense to remove values.
         */
        @Override
        public void remove()
        {
            throw new UnsupportedOperationException("ForEachTagIterator: Remove not supported!");
        }

        /**
         * Returns the next value.
         * @return The next value.
         */
        @Override
        public Object next()
        {
            Object next = null;
            myIndex++;
            if (myIndex <= myCollection.size())
                next = myInternalIterator.next();
            if (logger.isDebugEnabled())
            {
                try
                {
                    logger.debug("ForEachTagIterator: next: \"{}\".", ((next == null) ? "(null)" : next.toString()));
                }
                catch (RuntimeException e)
                {
                    logger.warn("ForEachTagIterator: next: \"{}\".", e.getMessage());
                }
            }
            return next;
        }

        /**
         * Determines if there are any items left, possibly <code>null</code>
         * items if the limit is larger than the collection size.
         * @return <code>true</code> if there are more items to process;
         *    <code>false</code> otherwise.
         */
        @Override
        public boolean hasNext()
        {
            return myIndex < myLimit;
        }
    }
}

