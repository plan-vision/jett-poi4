package net.sf.jett.tag;

import java.util.HashMap;
import java.util.Map;

/**
 * A <code>JtTagLibrary</code> represents the built-in <code>TagLibrary</code>.
 *
 * @author Randy Gettman
 */
public class JtTagLibrary implements TagLibrary
{
    private static JtTagLibrary theLibrary = new JtTagLibrary();

    private Map<String, Class<? extends Tag>> myTagMap;

    /**
     * Singleton constructor.
     */
    private JtTagLibrary()
    {
        myTagMap = new HashMap<>();
        myTagMap.put("agg"         , AggTag.class);
        myTagMap.put("ana"         , AnaTag.class);
        myTagMap.put("comment"     , CommentTag.class);
        myTagMap.put("for"         , ForTag.class);
        myTagMap.put("forEach"     , ForEachTag.class);
        myTagMap.put("formula"     , FormulaTag.class);
        myTagMap.put("group"       , GroupTag.class);
        myTagMap.put("hideCols"    , HideColsTag.class);
        myTagMap.put("hideRows"    , HideRowsTag.class);
        myTagMap.put("hideSheet"   , HideSheetTag.class);
        myTagMap.put("hyperlink"   , HyperlinkTag.class);
        myTagMap.put("if"          , IfTag.class);
        myTagMap.put("image"       , ImageTag.class);
        myTagMap.put("multiForEach", MultiForEachTag.class);
        myTagMap.put("name"        , NameTag.class);
        myTagMap.put("null"        , NullTag.class);
        myTagMap.put("pageBreak"   , PageBreakTag.class);
        myTagMap.put("rickroll"    , RickrollTag.class);
        myTagMap.put("set"         , SetTag.class);
        myTagMap.put("span"        , SpanTag.class);
        myTagMap.put("style"       , StyleTag.class);
        myTagMap.put("total"       , TotalTag.class);
    }

    /**
     * Returns the singleton instance of a <code>JtTagLibrary</code>.
     * @return The <code>JtTagLibrary</code>.
     */
    public static JtTagLibrary getJtTagLibrary()
    {
        return theLibrary;
    }

    /**
     * Returns the <code>Map</code> of tag names to tag <code>Class</code>
     * objects, e.g. <code>"if" =&gt; IfTag.class</code>.
     * @return A <code>Map</code> of tag names to tag <code>Class</code>
     *    objects.
     */
    @Override
    public Map<String, Class<? extends Tag>> getTagMap()
    {
        return myTagMap;
    }
}

