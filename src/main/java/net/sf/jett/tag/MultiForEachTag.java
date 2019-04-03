package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.expression.Expression;
import net.sf.jett.model.Block;
import net.sf.jett.model.PastEndValue;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>MultiForEachTag</code> represents a repetitively placed
 * <code>Block</code> of <code>Cells</code>, with each repetition corresponding
 * to the same index into multiple <code>Collections</code>.
 * The <code>vars</code> attribute represents the variable names corresponding
 * to what each <code>Collection</code>'s item is known by.  The optional
 * <code>indexVar</code> attribute is the name of the variable that holds the
 * iterator index.  The optional <code>limit</code> attribute specifies a limit
 * to the number of iterations to be run for the <code>Collections</code>.  If
 * the limit is greater than the number of items in any of the collections,
 * then blank blocks will result, with the exact result dependent on "past end
 * action" rules.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li><em>Inherits all attributes from {@link BaseLoopTag}.</em></li>
 * <li>collections (required): <code>Collection</code></li>
 * <li>var (required): <code>String</code></li>
 * <li>indexVar (optional): <code>String</code></li>
 * <li>limit (optional): <code>int</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class MultiForEachTag extends BaseLoopTag
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Attribute for specifying the <code>Collections</code> over which to
     * iterate.
     */
    public static final String ATTR_COLLECTIONS = "collections";
    /**
     * Attribute for specifying the "looping variable" names.
     */
    public static final String ATTR_VARS = "vars";
    /**
     * Attribute for specifying the name of the variable to be exposed that
     * indicates the 0-based index position into the <code>Collection</code>.
     */
    public static final String ATTR_INDEXVAR = "indexVar";
    /**
     * Attribute for specifying the number of iterations to be displayed.
     */
    public static final String ATTR_LIMIT = "limit";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_COLLECTIONS, ATTR_VARS));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_INDEXVAR, ATTR_LIMIT));

    private List<Collection<Object>> myCollections = null;
    private List<String> myCollectionNames = null;
    private List<String> myVarNames = null;
    private String myIndexVarName = null;
    private int myLimit = 0;
    private int myMaxSize = 0;

    /**
     * Sets the <code>List</code> of <code>Collections</code> to be processed.
     * @param collections A <code>List</code> of <code>Collections</code>.
     * @since 0.3.0
     */
    public void setCollections(List<Collection<Object>> collections)
    {
        myCollections = collections;
    }

    /**
     * Sets the <code>List</code> of collection expressions.
     * @param collExpressions A <code>List</code> of collection expressions.
     * @since 0.3.0
     */
    public void setCollectionNames(List<String> collExpressions)
    {
        for (String collExpression : collExpressions)
        {
            addCollectionName(collExpression);
        }
    }

    /**
     * Extracts the collection expression from the delimiters and adds it to the
     * collection names.
     * @param collExpression A collection expression, with "${" and "}".
     * @since 0.3.0
     */
    private void addCollectionName(String collExpression)
    {
        int beginExprIdx = collExpression.indexOf(Expression.BEGIN_EXPR);
        int endExprIdx = collExpression.indexOf(Expression.END_EXPR);
        if (beginExprIdx != -1 && endExprIdx != -1 && endExprIdx > beginExprIdx)
        {
            myCollectionNames.add(collExpression.substring(beginExprIdx +
                    Expression.BEGIN_EXPR.length(), endExprIdx));
        }
    }

    /**
     * Sets the <code>List</code> of variable names.
     * @param varNames The <code>List</code> of variable names.
     * @since 0.3.0
     */
    public void setVarNames(List<String> varNames)
    {
        myVarNames = varNames;
    }

    /**
     * Sets the "looping" variable name.
     * @param indexVarName The "looping" variable name.
     * @since 0.3.0
     */
    public void setIndexVarName(String indexVarName)
    {
        myIndexVarName = indexVarName;
    }

    /**
     * Sets the limit on the number of iterations.
     * @param limit The limit on the number of iterations.
     */
    public void setLimit(int limit)
    {
        myLimit = limit;
    }

    /**
     * Sets the maximum size of all collections.
     * @since 0.3.0
     */
    private void setMaxSize()
    {
        myMaxSize = 0;
        for (Collection<Object> collection : myCollections)
        {
            int size = collection.size();
            if (size > myMaxSize)
                myMaxSize = size;
        }
    }

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "multiForEach";
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
            throw new TagParseException("MultiForEach tags must have a body.  Bodiless MultiForEach tag found " + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        String attrCollExpressions = attributes.get(ATTR_COLLECTIONS).getString();
        String[] collExpressions = attrCollExpressions.split(SPEC_SEP);
        myCollections = new ArrayList<>();
        myCollectionNames = new ArrayList<>();
        for (String collExpression : collExpressions)
        {
            Object items = Expression.evaluateString(collExpression.trim(), getWorkbookContext().getExpressionFactory(), beans);
            if (items == null)
            {
                // Allow null to be interpreted as an empty collection.
                items = new ArrayList<>(0);
            }
            if (!(items instanceof Collection))
                throw new TagParseException("One of the items in the \"collections\" attribute is not a Collection in MultiForEach tag found"
                        + getLocation() + ": " + collExpression);
            Collection<Object> collection = AttributeUtil.evaluateObject(this, collExpression.trim(), beans, ATTR_COLLECTIONS,
                    Collection.class, null);
            myCollections.add(collection);
            // Collection names.
            addCollectionName(collExpression);
            logger.debug("Collection \"{}\" has size {}", collExpression, collection.size());
        }

        myVarNames = AttributeUtil.evaluateList(this, attributes.get(ATTR_VARS), beans, new ArrayList<String>(0));

        if (myCollections.size() < 1)
            throw new TagParseException("Must specify at least one Collection in a MultiForEachTag.  None found" + getLocation());
        if (myCollections.size() != myVarNames.size())
            throw new TagParseException("The number of collections and the number of variable names must be the same.  Mismatch found" +
                    getLocation());

        myIndexVarName = AttributeUtil.evaluateString(this, attributes.get(ATTR_INDEXVAR), beans, null);

        // Determine the maximum size of all collections.
        setMaxSize();

        myLimit = AttributeUtil.evaluateNonNegativeInt(this, attributes.get(ATTR_LIMIT), beans, ATTR_LIMIT, myMaxSize);

        logger.debug("vA: myLimit={}", myLimit);
    }

    /**
     * Returns the names of the <code>Collections</code> that are being used in
     * this <code>MultiForEachTag</code>.
     * @return A <code>List</code> of multiple collection names.
     */
    @Override
    protected List<String> getCollectionNames()
    {
        return myCollectionNames;
    }

    /**
     * Returns the names of the variables that are being used in this
     * <code>MultiForEachTag</code>.
     * @return A <code>List</code> of variable names.
     * @since 0.7.0
     */
    @Override
    protected List<String> getVarNames()
    {
        return myVarNames;
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
     * Returns the maximum size of the collections being iterated.
     * @return The maximum size of the collections being iterated.
     */
    @Override
    protected int getCollectionSize()
    {
        return myMaxSize;
    }

    /**
     * Returns an <code>Iterator</code> that iterates over all the items of all
     * specified <code>Collections</code> of values.  Its item is a
     * <code>List</code> of items created by pulling values from all
     * <code>Collections</code> using the same index for each
     * <code>Collection</code>.
     * @return An <code>Iterator</code>.
     */
    @Override
    protected Iterator<List<Object>> getLoopIterator()
    {
        return new MultiForEachTagIterator();
    }

    /**
     * Place the values from the <code>List</code> of collection item values
     * into the <code>Map</code> of beans.
     *
     * @param context The <code>TagContext</code>.
     * @param currBlock The <code>Block</code> that is about to processed.
     * @param item The <code>Object</code> that resulted from the iterator.
     * @param index The iteration index (0-based).
     */
    @Override
    @SuppressWarnings("unchecked")
    protected void beforeBlockProcessed(TagContext context, Block currBlock, Object item, int index)
    {
        Map<String, Object> beans = context.getBeans();
        List<Object> listOfValues = (List<Object>) item;
        List<String> pastEndRefs = new ArrayList<>();
        for (int i = 0; i < myCollections.size(); i++)
        {
            String varName = myVarNames.get(i);
            Object value = listOfValues.get(i);
            if (value != null && value instanceof PastEndValue)
                pastEndRefs.add(varName);
            else
                beans.put(varName, value);
        }

        logger.debug("beforeBP: index={}", index);
        // If not past the "collection" size, but a Collection is exhausted, then
        // take "past end actions" on individual Cells in tbe Block.
        if (index < getCollectionSize())
            SheetUtil.takePastEndAction(context.getSheet(), currBlock, pastEndRefs, getPastEndAction(), getReplacementExprValue());

        // Optional index counter variable.
        if (myIndexVarName != null && myIndexVarName.length() > 0)
            beans.put(myIndexVarName, index);
    }

    /**
     * Remove the values from the <code>List</code> of collection item values
     * from the <code>Map</code> of beans.
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
        for (int i = 0; i < myCollections.size(); i++)
            beans.remove(myVarNames.get(i));

        // Optional index counter variable.
        if (myIndexVarName != null && myIndexVarName.length() > 0)
            beans.remove(myIndexVarName);
    }

    /**
     * The <code>Iterator</code> over the items in all collections, which can be
     * extended by a large limit to return <code>nulls</code> beyond the size
     * of each <code>Collection</code>.
     */
    private class MultiForEachTagIterator implements Iterator<List<Object>>
    {
        private int myIndex;
        private List<Iterator<Object>> myIterators;

        /**
         * Construct a <code>MultiForEachTagIterator</code> that is initialized to
         * zero.
         */
        private MultiForEachTagIterator()
        {
            myIndex = 0;
            myIterators = new ArrayList<>();
            for (Collection<Object> collection : myCollections)
                myIterators.add(collection.iterator());
        }

        /**
         * It doesn't make sense to remove values.
         */
        @Override
        public void remove()
        {
            throw new UnsupportedOperationException("MultiForEachTagIterator: Remove not supported!");
        }

        /**
         * Returns the next value.  Each iteration produces a <code>List</code>
         * of variable values.  The values are the <code>Collection</code> values
         * from all specified collections, using the same index into all
         * <code>Collections</code>.
         * @return A <code>List</code> of variable values.
         */
        @Override
        public List<Object> next()
        {
            List<Object> next = new ArrayList<>();
            for (int i = 0; i < myCollections.size(); i++)
            {
                Object value = PastEndValue.PAST_END_VALUE;
                Iterator<Object> iterator = myIterators.get(i);
                if (iterator.hasNext())
                    value = iterator.next();
                next.add(value);
            }
            myIndex++;
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

