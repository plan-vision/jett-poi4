package net.sf.jett.tag;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import net.sf.jagg.Analytic;
import net.sf.jagg.AnalyticAggregator;
import net.sf.jagg.model.AnalyticValue;
import net.sf.jagg.exception.JaggException;
import org.apache.poi.ss.usermodel.RichTextString;

import net.sf.jett.exception.TagParseException;
import net.sf.jett.transform.BlockTransformer;
import net.sf.jett.util.AttributeUtil;

/**
 * <p>An <code>AnaTag</code> represents analytic values calculated from a
 * <code>List</code> of values already exposed to the context.  It uses
 * <code>jAgg</code> functionality and exposes the results and
 * <code>AnalyticAggregators</code> used for display later.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>items (required): <code>List</code></li>
 * <li>analytics (required): <code>String</code></li>
 * <li>analyticsVar (optional): <code>String</code></li>
 * <li>valuesVar (required): <code>String</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.9.0
 */
public class AnaTag extends BaseTag
{
    /**
     * Attribute that specifies the <code>List</code> of items to analyze.
     */
    public static final String ATTR_ITEMS = "items";
    /**
     * Attribute that specifies the <code>List</code> of analytic functions to
     * use.
     */
    public static final String ATTR_ANALYTICS = "analytics";
    /**
     * Attribute that specifies the name of the <code>List</code> of exposed
     * AnalyticAggregators.
     */
    public static final String ATTR_ANALYTICS_VAR = "analyticsVar";
    /**
     * Attribute that specifies name of the <code>List</code> of exposed
     * analytic values.
     */
    public static final String ATTR_VALUES_VAR = "valuesVar";

    private static final List<String> REQ_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ITEMS, ATTR_ANALYTICS, ATTR_VALUES_VAR));
    private static final List<String> OPT_ATTRS =
            new ArrayList<>(Arrays.asList(ATTR_ANALYTICS_VAR));

    private List<Object> myList = null;
    private List<AnalyticAggregator> myAnalytics = null;
    private String myAnalyticsVar = null;
    private String myValuesVar = null;
    private Analytic myAnalytic;

    /**
     * Returns this <code>Tag's</code> name.
     * @return This <code>Tag's</code> name.
     */
    public String getName()
    {
        return "ana";
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
     * attribute must be a <code>List</code>.  The "analytics" attribute must be
     * a semicolon-separated list of valid analytic specification strings.
     * The "valuesVar" attribute must be a string that indicates the name to
     * which the analytic values will be exposed in the <code>Map</code> of
     * beans.  The "analyticsVar" attribute must be a string that indicates the
     * name of the <code>List</code> that contains all created
     * <code>AnalyticAggregators</code> and to which that will be exposed in the
     * <code>Map</code> of beans.  The "ana" tag must have a body.
     */
    @SuppressWarnings("unchecked")
    @Override
    public void validateAttributes() throws TagParseException
    {
        super.validateAttributes();
        if (isBodiless())
            throw new TagParseException("Ana tags must have a body.  Bodiless agg tag found" + getLocation());

        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();
        Map<String, RichTextString> attributes = getAttributes();

        myList = AttributeUtil.evaluateObject(this, attributes.get(ATTR_ITEMS), beans, ATTR_ITEMS, List.class, null);

        List<String> analyticsList = AttributeUtil.evaluateList(this, attributes.get(ATTR_ANALYTICS), beans, null);
        myAnalytics = new ArrayList<>(analyticsList.size());
        for (String anaSpec : analyticsList)
            myAnalytics.add(AnalyticAggregator.getAnalytic(anaSpec));

        myAnalyticsVar = AttributeUtil.evaluateString(this, attributes.get(ATTR_ANALYTICS_VAR), beans, null);

        myValuesVar = AttributeUtil.evaluateString(this, attributes.get(ATTR_VALUES_VAR), beans, null);

        Analytic.Builder builder = new Analytic.Builder()
                .setAnalytics(myAnalytics);

        try
        {
            myAnalytic = builder.build();
        }
        catch (JaggException e)
        {
            throw new TagParseException("AnaTag: Problem executing jAgg call: " + getLocation()
                    + ": " + e.getMessage(), e);
        }
        catch (RuntimeException e)
        {
            throw new TagParseException("AnaTag: RuntimeException caught during jAgg execution" + getLocation()
                    + ": " + e.getMessage(), e);
        }
    }

    /**
     * Run an "analyze" operation on the specified <code>AnalyticFunctions</code>, get
     * the results, and expose the analytic values and the
     * <code>AnalyticAggregators</code> used.
     * @return Whether the first <code>Cell</code> in the <code>Block</code>
     *    associated with this <code>Tag</code> was processed.
     */
    @Override
    public boolean process()
    {
        TagContext context = getContext();
        Map<String, Object> beans = context.getBeans();

        List<AnalyticValue<Object>> aggValues = myAnalytic.analyze(myList);
        beans.put(myValuesVar, aggValues);
        if (myAnalyticsVar != null)
            beans.put(myAnalyticsVar, myAnalytics);

        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(context, getWorkbookContext());

        return true;
    }
}
