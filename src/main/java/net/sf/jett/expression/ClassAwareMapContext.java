package net.sf.jett.expression;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.jexl2.MapContext;

/**
 * A <code>ClassAwareMapContext</code> is a JEXL <code>MapContext</code> that
 * can resolve class names using <code>Class.forName</code> for the purposes
 * of allowing access to built-in Java methods, especially static methods such
 * as <code>java.lang.String.format</code>.
 *
 * @author Randy Gettman
 * @since 0.4.0
 */
public class ClassAwareMapContext extends MapContext
{
    private static final Map<String, Object> additionalValues;

    static
    {
        additionalValues = new HashMap<>();
        additionalValues.put("answerToLifeTheUniverseAndEverything", 42);
        additionalValues.put("JETT",
                Arrays.asList("Java Excel Template Translator", "Joan Jett", "Jett Travolta", "Bennie and the Jetts",
                        "Jett Blue", "New York Jetts", "Winnipeg Jetts", "The Jettsons", "Jett Engine", "Jettpack",
                        "Jett Li", "Lady Suffrajett", "Jett Propulsion Laboratory", "Jett Stream", "Jett Ski",
                        "Jett Lag", "Jett Fuel", "Jett Airliner"));
    }

    /**
     * Constructs a <code>ClassAwareMapContext</code> that uses an automatically
     * allocated, empty <code>Map</code>.
     */
    public ClassAwareMapContext()
    {
        super();
    }

    /**
     * Constructs a <code>ClassAwareMapContext</code> that wraps the existing
     * <code>Map</code> of bean names to bean values.
     * @param beans A <code>Map</code> of bean names to bean values.
     */
    public ClassAwareMapContext(Map<String, Object> beans)
    {
        super(beans);
    }

    /**
     * Checks whether the name is a variable defined in this context, or if it
     * is a valid Java class name as determined by <code>Class.forName</code>.
     * @param name The variable's or class's name.
     * @return <code>true</code> if the name exists, otherwise
     *    <code>false</code>.
     */
    @Override
    public boolean has(String name)
    {
        try
        {
            return super.has(name) || additionalValues.containsKey(name) || Class.forName(name) != null;
        }
        catch (ClassNotFoundException e)
        {
            return false;
        }
    }

    /**
     * Retrieves the variable value, or the <code>Class</code> object, by name
     * as defined in this context.
     * @param name The variable's name or class's name.
     * @return The variable value or <code>Class</code> object.
     */
    @Override
    public Object get(String name)
    {
        try
        {
            Object value = super.get(name);
            // Check for a legitimate null value for a variable name before
            // attempting to resolve a class name.
            if (value == null && !super.has(name) && (value = additionalValues.get(name)) == null)
            {
                value = Class.forName(name);
            }
            return value;
        }
        catch (ClassNotFoundException e)
        {
            return null;
        }
    }
}
