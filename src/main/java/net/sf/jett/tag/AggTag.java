package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jagg.AggregateFunction;
import net.sf.jagg.Aggregation;
import net.sf.jagg.Aggregator;
import net.sf.jagg.exception.JaggException;
import net.sf.jagg.model.AggregateValue;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;

/**
 * <p>An <code>AggTag</code> represents possibly many aggregate values
 * calculated from a <code>List</code> of values already exposed to the
 * context.  It uses <code>jAgg</code> functionality and exposes the results
 * and <code>AggregateFunctions</code> used for display later.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>items (required): <code>List</code></li>
 * <li>aggs (required): <code>String</code></li>
 * <li>aggsVar (optional): <code>String</code></li>
 * <li>valuesVar (required): <code>String</code></li>
 * <li>groupBy (optional): <code>String</code></li>
 * <li>parallel (optional): <code>int</code></li>
 * <li>useMsd (optional): <code>boolean</code></li>
 * <li>rollup (optional): <code>int[]</code></li>
 * <li>rollups (optional): <code>int[][]</code></li>
 * <li>cube (optional): <code>int[]</code></li>
 * <li>groupingSets (optional): <code>int[][]</code></li>
 * </ul>
 *
 * @author Randy Gettman
 */
public class AggTag extends BaseTag
{
    /**
     * Attribute that specifies the <code>List</code> of items to aggregate.
     */
    public static final String ATTR_ITEMS = "items";
    /**
     * Attribute that specifies the <code>List</code> of AggregateFunctions to use.
     */
    public static final String ATTR_AGGS = "aggs";
    /**
     * Attribute that specifies the name of the <code>List</code> of exposed
     * aggregate functions.
     */
    public static final String ATTR_AGGS_VAR = "aggsVar";
    /**
     * Attribute that specifies name of the <code>List</code> of exposed
     * aggregation values.
     */
    public static final String ATTR_VALUES_VAR = "valuesVar";
    /**
     * Attribute that specifies the <code>List</code> of group-by properties.
     */
    public static final String ATTR_GROUP_BY = "groupBy";
    /**
     * Attribute that specifies the degree of parallelism to use.
     */
    public static final String ATTR_PARALLEL = "parallel";
    /**
     * Attribute that specifies whether to use Multiset Discrimination instead
     * of sorting during aggregation processing.  This means that the results
     * will not be sorted by the "group by" properties.  It defaults to
     * <code>false</code> -- don't use Multiset Discrimination, use sorting.
     * @since 0.4.0
     */
    public static final String ATTR_USE_MSD = "useMsd";
    /**
     * Attribute that specifies a rollup that should occur on the results.
     * Specify a <code>List</code> of 0-based integer indexes that reference the
     * original list of properties.  E.g. when grouping by two properties,
     * <code>prop1</code> and <code>prop2</code>, specifying a <code>List</code>
     * of <code>{0}</code> specifies a rollup on <code>prop1</code>.
     * @since 0.4.0
     * @see #ATTR_GROUP_BY
     */
    public static final String ATTR_ROLLUP = "rollup";
    /**
     * Attribute that specifies multiple rollups that should occur on the
     * results.  Specify a <code>List</code> of <code>Lists</code> of 0-based
     * integer indexes that reference the original list of properties.  E.g.
     * when grouping by three properties,
     * <code>prop1</code>, <code>prop2</code>, and <code>prop3</code>,
     * specifying a <code>List</code> of <code>{{0}, {1, 2}}</code> specifies a
     * rollup on <code>prop1</code>, and a separate rollup on <code>prop2</code>
     * and <code>prop3</code>.
     * @since 0.4.0
     * @see #ATTR_GROUP_BY
     */
    public static final String ATTR_ROLLUPS = "rollups";
    /**
     * Attribute that specifies a data cube that should occur on the results.
     * Specify a <code>List</code> of 0-based integer indexes that reference the
     * original list of properties.  E.g. when grouping by three properties,
     * <code>prop1</code>, <code>prop2</code>, and <code>prop3</code>,
     * specifying a <code>List</code> of <code>{0, 1}</code> specifies a cube on
     * <code>prop1</code> and <code>prop2</code>.
     * @since 0.4.0
     * @see #ATTR_GROUP_BY
     */
    public static final String ATTR_CUBE = "cube";
    /**
     * Attribute that specifies the exact grouping sets that should occur on the
     * results.  Specify a <code>List</code> of <code>Lists</code> of 0-based
     * integer indexes that reference the original list of properties.  E.g.
     * when grouping by three properties, <code>prop1</code>,
     * <code>prop2</code>, and <code>prop3</code>, specifying a
     * <code>List</code> of <code>{{0}, {1, 2}, {}}</code> specifies a
     * grouping set on <code>prop1</code>, a separate grouping set on
     * <code>prop2</code> and <code>prop3</code>, and a third grouping set on no
     * properties (a grand total).
     * @since 0.4.0
     * @see #ATTR_GROUP_BY
     */
    public static final String ATTR_GROUPING_SETS = "groupingSets";
    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ITEMS, ATTR_AGGS, ATTR_VALUES_VAR));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_AGGS_VAR, ATTR_GROUP_BY, ATTR_PARALLEL, ATTR_USE_MSD, ATTR_ROLLUP,
                    ATTR_ROLLUPS, ATTR_CUBE, ATTR_GROUPING_SETS));

    private List<Object> myList = null;
    private List<AggregateFunction> myAggs = null;
    private String myAggsVar = null;
    private String myValuesVar = null;
    private Aggregation myAggregation;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    @Override
    public String getName()
    {
        return "agg";
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
     * a positive integer (defaults to 1).  The "aggs" attribute must be a
     * semicolon-separated list of valid <code>Aggregator</code> specification
     * strings.  The "valuesVar" attribute must be a string that indicates the
     * name to which the aggregate values will be exposed in the
     * <code>Map</code> of beans.  The "aggsVar" attribute must be a string that
     * indicates the name of the <code>List</code> that contains all created
     * <code>AggregateFunctions</code> and to which that will be exposed in the
     * <code>Map</code> of beans.  The "groupBy" attribute must be a semicolon-
     * separated list of properties with which to "group" aggregated
     * calculations (defaults to no "group by" properties).  The "agg" tag must
     * have a body.
     */
    @SuppressWarnings("unchecked")
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (isBodiless())
            throw new TagParseException("Agg tags must have a body.  Bodiless agg tag found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        myList = AttributeUtil.evaluateObject(this, attributes.get(ATTR_ITEMS), beans, ATTR_ITEMS, List.class, null);

        List<String> aggsList = AttributeUtil.evaluateList(this, attributes.get(ATTR_AGGS), beans, null);
        myAggs = new ArrayList<>(aggsList.size());
        for (String aggSpec : aggsList)
            myAggs.add(Aggregator.getAggregator(aggSpec));

        myAggsVar = AttributeUtil.evaluateString(this, attributes.get(ATTR_AGGS_VAR), beans, null);

        myValuesVar = AttributeUtil.evaluateString(this, attributes.get(ATTR_VALUES_VAR), beans, null);

        List<String> groupByProps = AttributeUtil.evaluateList(this, attributes.get(ATTR_GROUP_BY), beans, new ArrayList<String>());

        int parallelism = AttributeUtil.evaluatePositiveInt(this, attributes.get(ATTR_PARALLEL), beans, ATTR_PARALLEL, 1);

        boolean useMsd = AttributeUtil.evaluateBoolean(this, attributes.get(ATTR_USE_MSD), beans, false);

        RichTextString rtsRollup = attributes.get(ATTR_ROLLUP);
        RichTextString rtsRollups = attributes.get(ATTR_ROLLUPS);
        RichTextString rtsCube = attributes.get(ATTR_CUBE);
        RichTextString rtsGroupingSets = attributes.get(ATTR_GROUPING_SETS);
        AttributeUtil.ensureAtMostOneExists(this, Arrays.asList(rtsRollup, rtsRollups, rtsCube, rtsGroupingSets),
                Arrays.asList(ATTR_ROLLUP, ATTR_ROLLUPS, ATTR_CUBE, ATTR_GROUPING_SETS));
        List<Integer> rollup = AttributeUtil.evaluateIntegerArray(this, rtsRollup, beans, null);
        List<Integer> cube = AttributeUtil.evaluateIntegerArray(this, attributes.get(ATTR_CUBE), beans, null);
        List<List<Integer>> rollups = AttributeUtil.evaluateIntegerArrayArray(this, rtsRollups, beans, null);
        List<List<Integer>> groupingSets = AttributeUtil.evaluateIntegerArrayArray(this, rtsGroupingSets, beans, null);

        Aggregation.Builder builder = new Aggregation.Builder()
                .setAggregators(myAggs)
                .setParallelism(parallelism)
                .setUseMsd(useMsd);
        if (groupByProps != null)
            builder.setProperties(groupByProps);

        if (rollup != null)
            builder.setRollup(rollup);
        else if (cube != null)
            builder.setCube(cube);
        else if (rollups != null)
            builder.setRollups(rollups);
        else if (groupingSets != null)
            builder.setGroupingSets(groupingSets);

        try
        {
            myAggregation = builder.build();
        }
        catch (JaggException e)
        {
            throw new TagParseException("AggTag: Problem executing jAgg call: " + getLocation()
                    + ": " + e.getMessage(), e);
        }
        catch (RuntimeException e)
        {
            throw new TagParseException("AggTag: RuntimeException caught during jAgg execution" + getLocation()
                    + ": " + e.getMessage(), e);
        }
    }

    /**
     * Run a "group by" operation on the specified <code>AggregateFunctions</code>, get
     * the results, and expose the aggregate values and the
     * <code>AggregateFunctions</code> used.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        List<AggregateValue<Object>> aggValues = myAggregation.groupBy(myList);
        beans.put(myValuesVar, aggValues);
        if (myAggsVar != null)
            beans.put(myAggsVar, myAggs);

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, getWorkbookContext());

        return true;
    }
}