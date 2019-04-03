package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.ForLoopTagStatus;
import net.sf.jett.util.AttributeUtil;

/**
 * <p>A <code>ForTag</code> represents a repetitively placed <code>Block</code>
 * of <code>Cells</code>, with each repetition corresponding to an increment of
 * an index.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li><em>Inherits all attributes from {@link BaseLoopTag}.</em></li>
 * <li>var (required): <code>String</code></li>
 * <li>start (required): <code>int</code></li>
 * <li>end (required): <code>int</code></li>
 * <li>step (optional): <code>int</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class ForTag extends BaseLoopTag
{
    /**
     * Attribute for specifying the name of the looping variable.
     */
    public static final String ATTR_VAR = "var";
    /**
     * Attribute for specifying the starting value.
     */
    public static final String ATTR_START = "start";
    /**
     * Attribute for specifying the ending value (included in the range).
     */
    public static final String ATTR_END = "end";
    /**
     * Attribute for specifying how much the value increments per iteration.
     */
    public static final String ATTR_STEP = "step";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_VAR, ATTR_START, ATTR_END));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_STEP));

    private String myVarName;
    private int myStart;
    private int myEnd;
    private int myStep;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "for";
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
     * Validates the attributes for this <code>Tag</code>.  The "start", "end",
     * and "step" attributes must evaluate to <code>int</code>s.  If "step" is
     * not present, then it defaults to <code>1</code>.  The "step" must not be
     * zero.  It is possible for no loops to be processed if "step" is positive
     * and "start" is greater than "end", or if "step" is negative and "start"
     * is less than "end".
     */
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (isBodiless())
            throw new TagParseException("For tags must have a body.  Bodiless For tag found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        Map<String, RichTextString> attributes = getAttributes();

        myVarName = AttributeUtil.evaluateString(this, attributes.get(ATTR_VAR), beans, null);

        myStart = AttributeUtil.evaluateInt(this, attributes.get(ATTR_START), beans, ATTR_START, 0);

        myEnd = AttributeUtil.evaluateInt(this, attributes.get(ATTR_END), beans, ATTR_END, 0);

        myStep = AttributeUtil.evaluateNonZeroInt(this, attributes.get(ATTR_STEP), beans, ATTR_STEP, 1);
    }

    /**
     * Returns the names of the <code>Collections</code> that are being used in
     * this <code>ForTag</code>.
     * @return <code>null</code>, no collections are being used.
     */
    @Override
    protected List<String> getCollectionNames()
    {
        return null;
    }

    /**
     * Returns the variable names of the <code>Collections</code> that are being used in
     * this <code>ForTag</code>.
     * @return <code>null</code>, no variables are being used for any collections.
     * @since 0.7.0
     */
    @Override
    protected List<String> getVarNames()
    {
        return null;
    }

    /**
     * Returns the number of iterations.  Note that this effectively disables
     * the "limit" attribute for <code>ForTags</code>.
     * @return The number of iterations.
     */
    @Override
    protected int getNumIterations()
    {
        if ((myStep > 0 && myStart <= myEnd) || (myStep < 0 && myStart >= myEnd))
            return (myEnd - myStart) / myStep + 1;
        return 0;
    }

    /**
     * Returns the number of iterations.
     * @return The number of iterations.
     */
    @Override
    protected int getCollectionSize()
    {
        return getNumIterations();
    }

    /**
     * Returns a <code>ForLoopTagStatus</code> that will be exposed in the
     * beans map if the appropriate attribute is given.
     * @return A <code>ForLoopTagStatus</code>.
     * @since 0.9.1
     */
    @Override
    protected ForLoopTagStatus getLoopTagStatus()
    {
        return new ForLoopTagStatus(this, getNumIterations(), myStart, myEnd, myStep);
    }

    /**
     * Returns an <code>Iterator</code> that iterates over the desired values.
     * @return An <code>Iterator</code>.
     */
    @Override
    protected Iterator<Integer> getLoopIterator()
    {
        return new ForTagIterator();
    }

    /**
     * Place the index "item" into the <code>Map</code> of beans.
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
    }

    /**
     * Remove the index "item" from the <code>Map</code> of beans.
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
    }

    /**
     * The <code>Iterator</code> over the index values.
     */
    private class ForTagIterator implements Iterator<Integer>
    {
        private int myValue;

        /**
         * Construct a <code>ForTagIterator</code> that is initialized to the
         * start value.
         */
        private ForTagIterator()
        {
            myValue = myStart;
        }

        /**
         * It doesn't make sense to remove values.
         */
        @Override
        public void remove()
        {
            throw new UnsupportedOperationException("ForTagIterator: Remove not supported!");
        }

        /**
         * Returns the next value.
         * @return The next value.
         */
        @Override
        public Integer next()
        {
            int value = myValue;
            // Prepare the next value.
            myValue += myStep;
            return value;
        }

        /**
         * Returns <code>true</code> if there are more items to process;
         * <code>false</code> otherwise.
         * @return <code>true</code> if there are more items to process;
         *    <code>false</code> otherwise.
         */
        @Override
        public boolean hasNext()
        {
            return ((myStep > 0 && myValue <= myEnd) || (myStep < 0 && myValue >= myEnd));
        }
    }
}
