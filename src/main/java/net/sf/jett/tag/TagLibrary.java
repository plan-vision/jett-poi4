package net.sf.jett.tag;

import java.util.Map;

/**
 * A <code>TagLibrary</code> is a map of tag names to tag classes for a
 * particular namespace.
 *
 * @author Randy Gettman
 */
public interface TagLibrary
{
    /**
     * Returns the <code>Map</code> of tag names to tag <code>Class</code>
     * objects, e.g. <code>"if" =&gt; IfTag.class</code>.
     * @return A <code>Map</code> of tag names to tag <code>Class</code>
     *    objects.
     */
    public Map<String, Class<? extends Tag>> getTagMap();
}
