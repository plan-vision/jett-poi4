package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import net.sf.jagg.AggregateFunction;
import net.sf.jagg.Aggregations;
import net.sf.jagg.Aggregator;
import net.sf.jagg.model.AggregateValue;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;
import net.sf.jett.util.SheetUtil;

/**
 * <p>A <code>TotalTag</code> represents an aggregate value calculated from a
 * <code>List</code> of values already exposed to the context.  This uses
 * <code>jAgg</code> functionality.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>items (required): <code>List</code></li>
 * <li>value (required): <code>String</code></li>
 * <li>parallel (optional): <code>int</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class TotalTag extends BaseTag
{
    /**
     * Attribute that specifies the <code>List</code> of items to aggregate.
     */
    public static final String ATTR_ITEMS = "items";
    /**
     * Attribute that specifies the aggregator to use.
     */
    public static final String ATTR_VALUE = "value";
    /**
     * Attribute that specifies the degree of parallelism to use.
     */
    public static final String ATTR_PARALLEL = "parallel";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ITEMS, ATTR_VALUE));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_PARALLEL));

    private List<Object> myList = null;
    private AggregateFunction myAggregateFunction = null;
    private int myParallelism = 1;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "total";
    }

    /**
     * Returns a <code>List</code> of required attribute names.
     * @return A <code>List</code> of required attribute names.
     */
    @Override
    protected List<String> getRequiredAttributes()
    {
        List<String> reqAttrs = new ArrayList<>(super.getRequiredAttributes());
        reqAttrs.addAll(REQ_ATTRS);
        return reqAttrs;
    }

    /**
     * Returns a <code>List</code> of optional attribute names.
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
     * Validates the attributes for this <code>Tag</code>.  The "items"
     * attribute must be a <code>List</code>.  The "parallel" attribute must be
     * a positive integer (defaults to 1).  The "value" attribute must be a
     * valid <code>Aggregator</code> specification string.  The "total" tag must
     * not have a body.
     */
    @Override
    @SuppressWarnings("unchecked")
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (!isBodiless())
            throw new TagParseException("Total tags must not have a body.  Total tag with body found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        myList = AttributeUtil.evaluateObject(this, attributes.get(ATTR_ITEMS), beans, ATTR_ITEMS, List.class,
                new ArrayList<>(0));

        myParallelism = AttributeUtil.evaluatePositiveInt(this, attributes.get(ATTR_PARALLEL), beans, ATTR_PARALLEL, 1);

        String aggSpec = AttributeUtil.evaluateString(this, attributes.get(ATTR_VALUE), beans, null);
        myAggregateFunction = Aggregator.getAggregator(aggSpec);
    }

    /**
     * Run a "group by" operation on the specified <code>Aggregator</code>, get
     * the result, and set the cell value appropriately.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();

        List<String> propsList = new ArrayList<>(0);
        List<AggregateFunction> aggList = new ArrayList<>(1);
        aggList.add(myAggregateFunction);
        List<AggregateValue<Object>> aggValues =
                Aggregations.groupBy(myList, propsList, aggList, myParallelism);
        // There should be only one AggregateValue with no properties to group by.
        AggregateValue aggValue = aggValues.get(0);
        Object value = aggValue.getAggregateValue(myAggregateFunction);
        // Replace the bodiless tag text with the proper result.
        Row row = sheet.getRow(block.getTopRowNum());
        Cell cell = row.getCell(block.getLeftColNum());
        WorkbookContext workbookContext = getWorkbookContext();
        SheetUtil.setCellValue(workbookContext, cell, value, getAttributes().get(ATTR_VALUE));

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, workbookContext);
        return true;
    }
}