package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.event.TagLoopListener;
import net.sf.jett.event.TagLoopEvent;
import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.BaseLoopTagStatus;
import net.sf.jett.model.Block;
import net.sf.jett.model.PastEndAction;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>The abstract class <code>BaseLoopTag</code> is the base class for all tags
 * that represent loops.
 * </p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>copyRight (optional): <code>boolean</code></li>
 * <li>fixed (optional): <code>boolean</code></li>
 * <li>pastEndAction (optional): <code>String</code></li>
 * <li>replaceValue (optional): <code>String</code></li>
 * <li>groupDir (optional): <code>String</code></li>
 * <li>collapse (optional): <code>boolean</code></li>
 * <li>onLoopProcessed (optional): <code>TagLoopListener</code></li>
 * <li>varStatus (optional): <code>String</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public abstract class BaseLoopTag extends BaseTag
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Attribute for forcing "copy right" behavior.  (Default is copy down.)
     */
    public static final String ATTR_COPY_RIGHT = "copyRight";
    /**
     * Attribute for not shifting other content out of the way; works the same
     * as "fixed size collections".
     */
    public static final String ATTR_FIXED = "fixed";
    /**
     * Attribute for specifying the "past end action", an action for dealing
     * with content beyond the range of looping content.
     * @see #PAST_END_ACTION_CLEAR
     * @see #PAST_END_ACTION_REMOVE
     * @see #PAST_END_ACTION_REPLACE_EXPR
     */
    public static final String ATTR_PAST_END_ACTION = "pastEndAction";
    /**
     * Attribute for specifying the direction of the grouping.  This defaults to
     * no grouping.
     * @since 0.2.0
     * @see #GROUP_DIR_ROWS
     * @see #GROUP_DIR_COLS
     * @see #GROUP_DIR_NONE
     */
    public static final String ATTR_GROUP_DIR = "groupDir";
    /**
     * Attribute for specifying whether the group should be displayed collapsed.
     * The default is <code>false</code>, for not collapsed.  It is ignored if
     * neither rows nor columns are being grouped.
     * @since 0.2.0
     */
    public static final String ATTR_COLLAPSE = "collapse";
    /**
     * Attribute for specifying a <code>TagLoopListener</code> to listen for
     * <code>TagLoopEvents</code>.
     * @since 0.3.0
     */
    public static final String ATTR_ON_LOOP_PROCESSED = "onLoopProcessed";
    /**
     * Attribute for specifying a replacement value for expressions that
     * reference a collection that is past the end of iteration.  This defaults
     * to an empty string <code>""</code>, and is only relevant when the "past
     * end action" is <code>replaceExpr</code>.
     * @see #ATTR_PAST_END_ACTION
     * @see #PAST_END_ACTION_REPLACE_EXPR
     * @since 0.7.0
     */
    public static final String ATTR_REPLACE_VALUE = "replaceValue";
    /**
     * Attribute for specifying the name of the {@link net.sf.jett.model.LoopTagStatus}
     * object that will be exposed in the beans map.  If this attribute is not
     * present, or the value is <code>null</code>, then no such object will be
     * exposed.
     * @since 0.9.1
     */
    public static final String ATTR_VAR_STATUS = "varStatus";

    /**
     * The "past end action" value to clear the content of cells.
     */
    public static final String PAST_END_ACTION_CLEAR = "clear";
    /**
     * The "past end action" value to remove the cells, including things like
     * borders and formatting.
     */
    public static final String PAST_END_ACTION_REMOVE = "remove";
    /**
     * The "past end action" value to replace only the expressions that contain
     * past-end references.
     * @since 0.7.0
     */
    public static final String PAST_END_ACTION_REPLACE_EXPR = "replaceExpr";

    /**
     * The "group dir" value to specify that columns should be grouped.
     * @since 0.2.0
     */
    public static final String GROUP_DIR_COLS = "cols";
    /**
     * The "group dir" value to specify that rows should be grouped.
     * @since 0.2.0
     */
    public static final String GROUP_DIR_ROWS = "rows";
    /**
     * The "group dir" value to specify that neither rows nor columns should be
     * grouped.
     * @since 0.2.0
     */
    public static final String GROUP_DIR_NONE = "none";

    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_COPY_RIGHT, ATTR_FIXED, ATTR_PAST_END_ACTION,
                    ATTR_REPLACE_VALUE, ATTR_GROUP_DIR, ATTR_COLLAPSE, ATTR_ON_LOOP_PROCESSED, ATTR_VAR_STATUS));

    private boolean amIExplicitlyCopyingRight = false;
    private boolean amIFixed = false;
    private PastEndAction myPastEndAction = PastEndAction.CLEAR_CELL;
    private String myReplaceExprValue = "";
    private Block.Direction myGroupDir;
    private boolean amICollapsed;
    private TagLoopListener myTagLoopListener;
    private String myVarStatusName;

    /**
     * Sets whether the repeated blocks are to be copied to the right (true) or
     * downward (default, false).
     * @param copyRight Whether the repeated blocks are to be copied to the right (true) or
     *    downward (default, false).
     * @since 0.3.0
     */
    public void setCopyRight(boolean copyRight)
    {
        amIExplicitlyCopyingRight = copyRight;
    }

    /**
     * Sets "fixed" mode, which doesn't shift other content out of the way when
     * copying repeated blocks of cells.
     * @param fixed Whether to execute in "fixed" mode.
     * @since 0.3.0
     */
    public void setFixed(boolean fixed)
    {
        amIFixed = fixed;
    }

    /**
     * Sets the <code>PastEndAction</code>.
     * @param pae The <code>PastEndAction</code>.
     * @since 0.3.0
     */
    public void setPastEndAction(PastEndAction pae)
    {
        myPastEndAction = pae;
    }

    /**
     * Sets the replacement expression value.  This defaults to an empty string
     * <code>""</code>.  This is only relevant if a past end action of
     * "replaceExpr" is used.
     * @param value The replacement expression value.
     * @see #setPastEndAction
     * @see #PAST_END_ACTION_REPLACE_EXPR
     * @since 0.7.0
     */
    public void setReplaceExprValue(String value)
    {
        myReplaceExprValue = value;
    }

    /**
     * Sets the directionality of the Excel Group to be created, if any.
     * @param direction The directionality.
     * @since 0.3.0
     */
    public void setGroupDirection(Block.Direction direction)
    {
        myGroupDir = direction;
    }

    /**
     * Sets whether any Excel Group created is collapsed.
     * @param collapsed Whether any Excel group created is collapsed.
     * @since 0.3.0
     */
    public void setCollapsed(boolean collapsed)
    {
        amICollapsed = collapsed;
    }

    /**
     * Sets the <code>TagLoopListener</code>.
     * @param listener The <code>TagLoopListener</code>.
     * @since 0.3.0
     */
    public void setOnLoopProcessed(TagLoopListener listener)
    {
        myTagLoopListener = listener;
    }

    /**
     * There are no required attributes that all <code>BaseLoopTags</code>
     * support.
     * @return An empty <code>List</code>.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        return super.getRequiredAttributes();
    }

    /**
     * All <code>BaseLoopTags</code> support the optional copy down tag.
     * @return A <code>List</code> of optional attribute names.
     */
    @Override
    protected List<String> getOptionalAttributes()
    {
        List<String> optAttrs = new ArrayList<>(super.getOptionalAttributes());
        optAttrs.addAll(OPT_ATTRS);
        return optAttrs;
    }

    /**
     * Ensure that the past end action (if specified) is a valid value.  Ensure
     * that the group direction (if specified) is a valid value.
     * @throws TagParseException If the attribute values are illegal or
     *    unacceptable.
     */
    @Override
    protected void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();
        Block block = context.getBlock();

        amIExplicitlyCopyingRight = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_COPY_RIGHT), beans, false);
        if (amIExplicitlyCopyingRight)
            block.setDirection(Block.Direction.HORIZONTAL);

        amIFixed = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_FIXED), beans, false);

        String strPastEndAction = AttributeUtil.evaluateStringSpecificValues(this, attributes.get(ATTR_PAST_END_ACTION), beans,
                ATTR_PAST_END_ACTION,
                Arrays.asList(PAST_END_ACTION_CLEAR, PAST_END_ACTION_REMOVE, PAST_END_ACTION_REPLACE_EXPR),
                PAST_END_ACTION_CLEAR);
        if (PAST_END_ACTION_CLEAR.equalsIgnoreCase(strPastEndAction))
            myPastEndAction = PastEndAction.CLEAR_CELL;
        else if (PAST_END_ACTION_REMOVE.equalsIgnoreCase(strPastEndAction))
            myPastEndAction = PastEndAction.REMOVE_CELL;
        else if (PAST_END_ACTION_REPLACE_EXPR.equalsIgnoreCase(strPastEndAction))
            myPastEndAction = PastEndAction.REPLACE_EXPR;

        myReplaceExprValue = AttributeUtil.evaluateString(this, attributes.get(ATTR_REPLACE_VALUE), beans, "");

        String strGroupDir = AttributeUtil.evaluateStringSpecificValues(this, attributes.get(ATTR_GROUP_DIR), beans,
                ATTR_GROUP_DIR, Arrays.asList(GROUP_DIR_ROWS, GROUP_DIR_COLS, GROUP_DIR_NONE), GROUP_DIR_NONE);
        if (GROUP_DIR_ROWS.equals(strGroupDir))
            myGroupDir = Block.Direction.VERTICAL;
        else if (GROUP_DIR_COLS.equals(strGroupDir))
            myGroupDir = Block.Direction.HORIZONTAL;
        else if (GROUP_DIR_NONE.equals(strGroupDir))
            myGroupDir = Block.Direction.NONE;

        amICollapsed = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_COLLAPSE), beans, false);

        myTagLoopListener = AttributeUtil.evaluateObject(this, attributes.get(ATTR_ON_LOOP_PROCESSED), beans,
                ATTR_ON_LOOP_PROCESSED, TagLoopListener.class, null);

        myVarStatusName = AttributeUtil.evaluateString(this, attributes.get(ATTR_VAR_STATUS), beans, null);
    }

    /**
     * Returns the <code>PastEndAction</code>, which is controlled by the
     * attribute specified by <code>ATTR_PAST_END_ACTION</code>.  It defaults to
     * <code>CLEAR_CELL</code>.
     * @return A <code>PastEndAction</code>.
     * @see PastEndAction
     */
    protected PastEndAction getPastEndAction()
    {
        return myPastEndAction;
    }

    /**
     * Returns the replacement expression value, which defaults to am empty
     * string <code>""</code>.  This is only relevant if the past end action is
     * "replaceExpr".
     * @return The replacement expression value.
     * @see #getPastEndAction
     * @see #PAST_END_ACTION_REPLACE_EXPR
     * @see #ATTR_REPLACE_VALUE
     * @since 0.7.0
     */
    protected String getReplacementExprValue()
    {
        return myReplaceExprValue;
    }

    /**
     * <p>Provide a generic way to process a tag that loops, with the Template
     * Method pattern.</p>
     * <ol>
     * <li>Decide whether content needs to be shifted out of the way, and shift
     * the content out of the way if necessary.  This involves calling
     * <code>getCollectionNames()</code> to determine if any of the collection
     * names are "fixed".  This also involves calling <code>getVarNames()</code>
     * to determine which, if any, past end actions, need to be taken.</li>
     * <li>Call <code>getNumIterations</code> to determine the number of Blocks
     * needed.</li>
     * <li>Copy the Block the needed number of times.</li>
     * <li>Get the loop iterator by calling <code>getLoopIterator()</code>.</li>
     * <li>Over each loop of the iterator...
     * <ol>
     * <li>Create a <code>Block</code> for the iteration.</li>
     * <li>If the collection values are exhausted, apply any "past end actions".
     * </li>
     * <li>Call <code>beforeBlockProcessed()</code>.</li>
     * <li>Process the current <code>Block</code> with a
     * <code>BlockTransformer</code>.</li>
     * <li>Call <code>afterBlockProcessed()</code>.</li>
     * </ol>
     * </li>
     * </ol>
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     * @see #getCollectionNames
     * @see #getVarNames
     * @see #getNumIterations
     * @see #getLoopIterator
     * @see #beforeBlockProcessed
     * @see #afterBlockProcessed
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Block block = context.getBlock();
        WorkbookContext workbookContext = getWorkbookContext();
        // Important for formulas, so different cell reference map entries from
        // different loops can be distinguished.
        workbookContext.incrSequenceNbr();
        int seqNbr = workbookContext.getSequenceNbr();

        Sheet sheet = context.getSheet();
        Map<String, Object> beans = context.getBeans();

        // Decide whether this is "fixed" in 2 ways:
        // 1. A fixed size collection name was specified and is present.
        // 2. The "fixed" attribute is true.
        boolean fixed = amIFixed;
        if (!fixed)
        {
            // Shallow copy.
            List<String> fixedSizeCollNames = new ArrayList<>(
                    workbookContext.getFixedSizedCollectionNames());
            List<String> collNames = getCollectionNames();
            if (collNames != null)
            {
                logger.debug("collNames found: {}", collNames);
                // Remove all collection names not found.
                for (Iterator<String> itr = fixedSizeCollNames.iterator(); itr.hasNext(); )
                {
                    String fixedSizeCollName = itr.next();
                    if (!collNames.contains(fixedSizeCollName))
                        itr.remove();
                }
            }
            else
            {
                fixedSizeCollNames.clear();
            }
            fixed = !fixedSizeCollNames.isEmpty();
        }

        int numIterations = getNumIterations();
        List<String> varNames = getVarNames();
        logger.debug("BaseLoopTag: numIterations={}", numIterations);
        if (numIterations == 0)
        {
            // If fixed, no shifting is to occur for the removed block.
            if (fixed)
            {
                switch(myPastEndAction)
                {
                case CLEAR_CELL:
                    clearBlock();
                    break;
                case REMOVE_CELL:
                    deleteBlock();
                    break;
                case REPLACE_EXPR:
                    SheetUtil.takePastEndAction(sheet, block, varNames, myPastEndAction, myReplaceExprValue);
                    block.collapse();
                    break;
                default:
                    throw new IllegalStateException("BaseLoopTag: Unknown PastEndAction: " + myPastEndAction);
                }
            }
            else
                removeBlock();
            return false;
        }
        else
        {
            BlockTransformer transformer = new BlockTransformer();
            List<Block> blocksToProcess = new ArrayList<>(numIterations);
            // Create room for the additional Blocks; the Block knows the proper
            // direction (right or down).
            // Don't create room if the collection is "fixed size", i.e. we can
            // assume that room exists already.
            if (!fixed)
                shiftForBlock();

            // Copy the Block.
            for (int i = 0; i < numIterations; i++)
            {
                Block copy = copyBlock(i);
                logger.debug("  Adding copied block: {}", copy);
                blocksToProcess.add(copy);
            }

            int index = 0;
            Iterator<?> iterator = getLoopIterator();
            BaseLoopTagStatus status = null;
            if (myVarStatusName != null && !myVarStatusName.isEmpty())
            {
                status = getLoopTagStatus();
                beans.put(myVarStatusName, status);
            }
            int right, bottom, colGrowth, rowGrowth;
            int maxRight = 0;
            int maxBottom = 0;
            while(iterator.hasNext())
            {
                Object item = iterator.next();
                Block currBlock = blocksToProcess.get(index);

                // Off the end of the collection!
                if (index >= getCollectionSize())
                {
                    switch(myPastEndAction)
                    {
                    case CLEAR_CELL:
                        SheetUtil.clearBlock(sheet, currBlock, getWorkbookContext());
                        break;
                    case REMOVE_CELL:
                        SheetUtil.deleteBlock(sheet, context, currBlock, getWorkbookContext());
                        break;
                    case REPLACE_EXPR:
                        SheetUtil.takePastEndAction(sheet, currBlock, varNames, myPastEndAction, myReplaceExprValue);
                        break;
                    default:
                        throw new IllegalStateException("BaseLoopTag: Unknown PastEndAction: " + myPastEndAction);
                    }
                }

                // Before Block Processing.
                beforeBlockProcessed(context, currBlock, item, index);

                // Fire a before tag loop processed event here, after the Before
                // Block Processing occurs.
                if (fireBeforeTagLoopProcessedEvent(currBlock, index))
                {
                    // Process the block.
                    TagContext blockContext = new TagContext();
                    blockContext.setSheet(sheet);
                    blockContext.setBeans(beans);
                    blockContext.setBlock(currBlock);
                    blockContext.setProcessedCellsMap(context.getProcessedCellsMap());
                    blockContext.setDrawing(context.getDrawing());
                    blockContext.setMergedRegions(context.getMergedRegions());
                    blockContext.setCurrentTag(this);
                    String suffix = context.getFormulaSuffix() + "[" + seqNbr + "," + index + "]";
                    blockContext.setFormulaSuffix(suffix);

                    logger.debug("  Block Before: {}", currBlock);
                    right = currBlock.getRightColNum();
                    bottom = currBlock.getBottomRowNum();

                    transformer.transform(blockContext, workbookContext);
                    // See if the block transformation grew or shrunk the current block.
                    logger.debug("  Block After: {}", currBlock);
                    colGrowth = currBlock.getRightColNum() - right;
                    rowGrowth = currBlock.getBottomRowNum() - bottom;
                    // If it did, then all pending blocks must react!
                    if (colGrowth != 0 || rowGrowth != 0)
                    {
                        logger.trace("  colGrowth is {}, rowGrowth is {}", colGrowth, rowGrowth);
                        for (int j = index + 1; j < numIterations; j++)
                        {
                            Block pendingBlock = blocksToProcess.get(j);
                            logger.trace("    Reacting Block: {}", pendingBlock);
                            pendingBlock.reactToGrowth(currBlock, colGrowth, rowGrowth);
                        }
                    }
                    // Get max right/bottom to expand the tag's block later.
                    if (currBlock.getRightColNum() > maxRight)
                        maxRight = currBlock.getRightColNum();
                    if (currBlock.getBottomRowNum() > maxBottom)
                        maxBottom = currBlock.getBottomRowNum();

                    // Fire a tag loop processed event here, before the After Block Processing
                    // occurs.
                    fireTagLoopProcessedEvent(currBlock, index);
                }

                // After Block Processing.
                afterBlockProcessed(context, currBlock, item, index);

                // End of loop processing.
                if (status != null)
                {
                    status.incrementIndex(this);
                }
                index++;
            }  // End while loop over collection items

            if (status != null)
            {
                beans.remove(myVarStatusName);
            }

            // Expand the tag block.
            block.expand(maxRight - block.getRightColNum(), maxBottom - block.getBottomRowNum());

            // Grouping - only if there was at least one item to process.
            groupRowsOrCols(sheet, context.getBlock(), blocksToProcess.get(blocksToProcess.size() - 1));
        }
        return true;
    }

    /**
     * If there is a <code>TagLoopListener</code>, then create and fire a
     * <code>TagLoopEvent</code>, with beans and sheet taken from this
     * <code>BaseLoopTag</code>, and with the given loop index and given
     * <code>Block</code>.
     * @param block The current <code>Block</code>.
     * @param index The zero-based loop index.
     * @return Whether processing of the <code>Tag</code> loop iteration should
     *    occur.  If the <code>TagLoopListener's</code>
     *    <code>beforeTagLoopProcessed</code> method returns <code>false</code>,
     *    then this method returns <code>false</code>.
     * @since 0.8.0
     */
    private boolean fireBeforeTagLoopProcessedEvent(Block block, int index)
    {
        if (myTagLoopListener != null)
        {
            TagContext context = getContext();
            TagLoopEvent tagLoopEvent = new TagLoopEvent(context.getSheet(), block, context.getBeans(), index);
            return myTagLoopListener.beforeTagLoopProcessed(tagLoopEvent);
        }
        return true;
    }

    /**
     * If there is a <code>TagLoopListener</code>, then create and fire a
     * <code>TagLoopEvent</code>, with beans and sheet taken from this
     * <code>BaseLoopTag</code>, and with the given loop index and given
     * <code>Block</code>.
     * @param block The current <code>Block</code>.
     * @param index The zero-based loop index.
     */
    private void fireTagLoopProcessedEvent(Block block, int index)
    {
        if (myTagLoopListener != null)
        {
            TagContext context = getContext();
            TagLoopEvent tagLoopEvent = new TagLoopEvent(context.getSheet(), block, context.getBeans(), index);
            myTagLoopListener.onTagLoopProcessed(tagLoopEvent);
        }
    }

    /**
     * Decide to and place an Excel Group for rows, columns, or nothing,
     * depending on attribute settings and the first and last
     * <code>Blocks</code>.
     * @param sheet The <code>Sheet</code> on which to group rows or columns.
     * @param first The first <code>Block</code>.
     * @param last The last <code>Block</code>.
     */
    private void groupRowsOrCols(Sheet sheet, Block first, Block last)
    {
        int begin, end;
        logger.debug("gROC: {}, {}", myGroupDir, amICollapsed);
        switch(myGroupDir)
        {
        case VERTICAL:
            begin = first.getTopRowNum();
            end = last.getBottomRowNum();
            SheetUtil.groupRows(sheet, begin, end, amICollapsed);
            break;
        case HORIZONTAL:
            begin = first.getLeftColNum();
            end = last.getRightColNum();
            SheetUtil.groupColumns(sheet, begin, end, amICollapsed);
            break;
        // Do nothing on NONE.
        }
    }

    /**
     * Shifts cells out of the way of where copied blocks will go.
     */
    private void shiftForBlock()
    {
        TagContext context = getContext();
        Block block = context.getBlock();
        Sheet sheet = context.getSheet();
        int numIterations = getNumIterations();
        SheetUtil.shiftForBlock(sheet, context, block, getWorkbookContext(), numIterations);
    }

    /**
     * Copies the <code>Block</code> in a particular direction.
     * @param numBlocksAway How many blocks away the <code>Block</code> will be
     *    copied.
     * @return The newly copied <code>Block</code>.
     */
    private Block copyBlock(int numBlocksAway)
    {
        TagContext context = getContext();
        Block block = context.getBlock();
        Sheet sheet = context.getSheet();
        return SheetUtil.copyBlock(sheet, context, block, getWorkbookContext(), numBlocksAway);
    }

    /**
     * Returns the names of the <code>Collections</code> that are being used in
     * this <code>BaseLoopTag</code>.
     * @return A <code>List</code> collection names, or <code>null</code> if
     *    not operating on any <code>Collections</code>.
     */
    protected abstract List<String> getCollectionNames();

    /**
     * Returns the names of the variables that are being used in this
     * <code>BaseLoopTag</code>.
     * @return A <code>List</code> of variable names, or <code>null</code> if
     *    there are not any variable names into any <code>Collections</code>.
     * @since 0.7.0
     */
    protected abstract List<String> getVarNames();

    /**
     * Returns the number of iterations.
     * @return The number of iterations.
     */
    protected abstract int getNumIterations();

    /**
     * Returns the size of the collection being iterated.  This may be different
     * than the number of iterations because of the "limit" attribute.
     * @return The size of the collection being iterated.
     */
    protected abstract int getCollectionSize();

    /**
     * Returns a <code>BaseLoopTagStatus</code> that will be exposed in the
     * beans map if the appropriate attribute is given.  Subclasses may want to
     * override this method to return an object that provides more information.
     * @return A <code>BaseLoopTagStatus</code>.
     * @since 0.9.1
     */
    protected BaseLoopTagStatus getLoopTagStatus()
    {
        return new BaseLoopTagStatus(this, getNumIterations());
    }

    /**
     * Returns an <code>Iterator</code> that iterates over some
     * <code>Collection</code> of objects.  The <code>Iterator</code> doesn't
     * need to support the <code>remove</code> operation.
     * @return An <code>Iterator</code>.
     */
    protected abstract Iterator<?> getLoopIterator();

    /**
     * This method is called once per iteration loop, immediately before the
     * given <code>Block</code> is processed.  An iteration index is supplied as
     * well.
     * @param context The <code>TagContext</code>.
     * @param currBlock The <code>Block</code> that is about to processed.
     * @param item The <code>Object</code> that resulted from the iterator.
     * @param index The iteration index (0-based).
     */
    protected abstract void beforeBlockProcessed(TagContext context, Block currBlock, Object item, int index);

    /**
     * This method is called once per iteration loop, immediately after the
     * given <code>Block</code> is processed.  An iteration index is supplied as
     * well.
     * @param context The <code>TagContext</code>.
     * @param currBlock The <code>Block</code> that was just processed.
     * @param item The <code>Object</code> that resulted from the iterator.
     * @param index The iteration index (0-based).
     */
    protected abstract void afterBlockProcessed(TagContext context, Block currBlock, Object item, int index);
}
