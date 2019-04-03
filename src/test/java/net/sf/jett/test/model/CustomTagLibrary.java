package net.sf.jett.test.model;

import java.util.HashMap;
import java.util.Map;

import net.sf.jett.tag.Tag;
import net.sf.jett.tag.TagLibrary;

/**
 * The <code>CustomTagLibrary</code> is for testing the custom tag libraries
 * feature.
 *
 * @author Randy Gettman
 */
public class CustomTagLibrary implements TagLibrary
{
    private static CustomTagLibrary theLibrary = new CustomTagLibrary();

    private Map<String, Class<? extends Tag>> myTagMap;

    /**
     * Singleton constructor.
     */
    private CustomTagLibrary()
    {
        myTagMap = new HashMap<>();
        myTagMap.put("addOne", AddOneTag.class);
    }

    /**
     * Returns the singleton instance of a <code>CustomTagLibrary</code>.
     * @return The <code>CustomTagLibrary</code>.
     */
    public static CustomTagLibrary getCustomTagLibrary()
    {
        return theLibrary;
    }

    /**
     * Returns the <code>Map</code> of tag names to tag <code>Class</code>
     * objects.
     * @return A <code>Map</code> of tag names to tag <code>Class</code>
     *    objects.
     */
    @Override
    public Map<String, Class<? extends Tag>> getTagMap()
    {
        return myTagMap;
    }
}