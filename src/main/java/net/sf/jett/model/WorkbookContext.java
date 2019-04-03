package net.sf.jett.model;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import net.sf.jett.event.CellListener;
import net.sf.jett.event.SheetListener;
import net.sf.jett.expression.ExpressionFactory;
import net.sf.jett.formula.CellRef;
import net.sf.jett.formula.Formula;
import net.sf.jett.tag.TagLibraryRegistry;

/**
 * A <code>WorkbookContext</code> object holds data relevant to the context of
 * a <code>Workbook</code>.
 *
 * @author Randy Gettman
 */
public class WorkbookContext
{
    private TagLibraryRegistry myRegistry;
    private List<CellListener> myCellListeners;
    private List<SheetListener> mySheetListeners;
    private List<String> myFixedSizeCollectionNames;
    private List<String> myNoImplicitProcessingCollectionNames;
    private Map<String, Formula> myFormulaMap;
    private Map<String, String> myTagLocationsMap;
    private Map<String, List<CellRef>> myCellRefMap;
    private int mySequenceNbr;
    private CellStyleCache myCellStyleCache;
    private FontCache myFontCache;
    private Map<String, Style> myStyleMap;
    private List<String> myTemplateSheetNames;
    private List<String> mySheetNames;
    private ExpressionFactory myExpressionFactory;
    private List<Map<String, Object>> myBeansMaps;

    /**
     * Initializes things to null/0.
     */
    public WorkbookContext()
    {
        myRegistry = null;
        myCellListeners = null;
        mySheetListeners = null;
        myFixedSizeCollectionNames = null;
        myNoImplicitProcessingCollectionNames = null;
        myFormulaMap = null;
        myCellRefMap = null;
        mySequenceNbr = 0;
        myCellStyleCache = null;
        myFontCache = null;
        myStyleMap = null;
        myTemplateSheetNames = null;
        mySheetNames = null;
        myExpressionFactory = null;
        myBeansMaps = null;
    }

    /**
     * Returns the <code>TagLibraryRegistry</code>.
     * @return The <code>TagLibraryRegistry</code>.
     */
    public TagLibraryRegistry getRegistry()
    {
        return myRegistry;
    }

    /**
     * Sets the <code>TagLibraryRegistry</code>.
     * @param registry The <code>TagLibraryRegistry</code>.
     */
    public void setRegistry(TagLibraryRegistry registry)
    {
        myRegistry = registry;
    }

    /**
     * Returns the <code>CellListeners</code>.
     * @return The <code>CellListeners</code>.
     */
    public List<CellListener> getCellListeners()
    {
        return myCellListeners;
    }

    /**
     * Returns the <code>SheetListeners</code>.
     * @return The <code>SheetListeners</code>.
     * @since 0.8.0
     */
    public List<SheetListener> getSheetListeners()
    {
        return mySheetListeners;
    }

    /**
     * Sets the <code>SheetListeners</code>.
     * @param sheetListeners The <code>SheetListeners</code>.
     * @since 0.8.0
     */
    public void setSheetListeners(List<SheetListener> sheetListeners)
    {
        mySheetListeners = sheetListeners;
    }

    /**
     * Sets the <code>CellListeners</code>.
     * @param cellListeners The <code>CellListeners</code>.
     */
    public void setCellListeners(List<CellListener> cellListeners)
    {
        myCellListeners = cellListeners;
    }

    /**
     * These named <code>Collections</code> have a known size and do not need to
     * have other <code>Cells</code> shifted out of the way for its contents;
     * space is already allocated.
     * @param collNames A <code>List</code> of <code>Collection</code> names
     *    that don't need other <code>Cells</code> shifted out of the way for
     *    its contents.
     */
    public void setFixedSizeCollectionNames(List<String> collNames)
    {
        myFixedSizeCollectionNames = collNames;
    }

    /**
     * Returns the <code>List</code> of "fixed size" collection names.
     * @return The <code>List</code> of "fixed size" collection names.
     */
    public List<String> getFixedSizedCollectionNames()
    {
        return myFixedSizeCollectionNames;
    }

    /**
     * Turn off implicit collections processing for the given
     * <code>Collections</code> specified by the given collection names.
     * @param collNames The names of the <code>Collections</code> on which NOT
     *    to perform implicit collections processing.
     */
    public void setNoImplicitCollectionProcessingNames(List<String> collNames)
    {
        myNoImplicitProcessingCollectionNames = collNames;
    }

    /**
     * Returns the <code>List</code> of collection names on which NOT to perform
     * implicit collections processing.
     * @return The <code>List</code> of collection names on which NOT to perform
     *    implicit collections processing.
     */
    public List<String> getNoImplicitProcessingCollectionNames()
    {
        return myNoImplicitProcessingCollectionNames;
    }

    /**
     * Returns the formula map, a <code>Map</code> of formula keys to
     * <code>Formulas</code>, with the keys of the format "sheetName!formula".
     * @return A <code>Map</code> of formula keys to <code>Formulas</code>.
     */
    public Map<String, Formula> getFormulaMap()
    {
        return myFormulaMap;
    }

    /**
     * Sets the formula map, a <code>Map</code> of formula keys to
     * <code>Formulas</code>, with the keys of the format "sheetName!formula".
     * @param formulaMap A <code>Map</code> of formula keys to
     *    <code>Formulas</code>.
     */
    public void setFormulaMap(Map<String, Formula> formulaMap)
    {
        myFormulaMap = formulaMap;
    }

    /**
     * Returns the tag locations map, a <code>Map</code> of current tag location
     * cell references to original tag location cell references, with the cell
     * references being in the format "Sheet!B1".  This is currently used only to
     * identify original tag locations for exception messages.
     * @return A <code>Map</code> of current tag location cell references to
     *    original tag location cell references.
     * @since 0.9.0
     */
    public Map<String, String> getTagLocationsMap()
    {
        return myTagLocationsMap;
    }

    /**
     * Sets the tag locations map, a <code>Map</code> of current tag location
     * cell references to original tag location cell references, with the cell
     * references being in the format "Sheet!B1".  This is currently used only to
     * identify original tag locations for exception messages.
     * @param tagLocationsMap A <code>Map</code> of current tag location cell
     *    references to original tag location cell references.
     * @since 0.9.0
     */
    public void setTagLocationsMap(Map<String, String> tagLocationsMap)
    {
        myTagLocationsMap = tagLocationsMap;
    }

    /**
     * Returns the cell reference map, a <code>Map</code> of cell key strings to
     * <code>Lists</code> of <code>CellRefs</code>.  The cell key strings are
     * original cell references, and the <code>Lists</code> contain translated
     * <code>CellRefs</code>, e.g. "Sheet1!C2" =&gt; [C2, C3, C4]
     * @return A <code>Map</code> of cell key strings to <code>Lists</code> of
     *    <code>CellRefs</code>.
     */
    public Map<String, List<CellRef>> getCellRefMap()
    {
        return myCellRefMap;
    }

    /**
     * Sets the cell reference map, a <code>Map</code> of cell key strings to
     * <code>Lists</code> of <code>CellRefs</code>.  The cell key strings are
     * original cell references, and the <code>Lists</code> contain translated
     * <code>CellRefs</code>, e.g. "Sheet1!C2" =&gt; [C2, C3, C4]
     * @param cellRefMap A <code>Map</code> of cell key strings to
     * <code>Lists</code> of <code>CellRefs</code>.
     */
    public void setCellRefMap(Map<String, List<CellRef>> cellRefMap)
    {
        myCellRefMap = cellRefMap;
    }

    /**
     * Returns the current sequence number.
     * @return The current sequence number.
     */
    public int getSequenceNbr()
    {
        return mySequenceNbr;
    }

    /**
     * Increments the current sequence number.
     */
    public void incrSequenceNbr()
    {
        mySequenceNbr++;
    }

    /**
     * Returns the <code>CellStyleCache</code>.
     * @return The <code>CellStyleCache</code>.
     * @since 0.5.0
     */
    public CellStyleCache getCellStyleCache()
    {
        return myCellStyleCache;
    }

    /**
     * Sets the <code>CellStyleCache</code>.
     * @param cache A <code>CellStyleCache</code>.
     * @since 0.5.0
     */
    public void setCellStyleCache(CellStyleCache cache)
    {
        myCellStyleCache = cache;
    }

    /**
     * Returns the <code>FontCache</code>.
     * @return The <code>FontCache</code>.
     * @since 0.5.0
     */
    public FontCache getFontCache()
    {
        return myFontCache;
    }

    /**
     * Sets the <code>FontCache</code>.
     * @param cache The <code>FontCache</code>.
     * @since 0.5.0
     */
    public void setFontCache(FontCache cache)
    {
        myFontCache = cache;
    }

    /**
     * Returns the <code>Map</code> of style names to <code>Styles</code>.
     * @return The <code>Map</code> of style names to <code>Styles</code>.
     * @since 0.5.0
     */
    public Map<String, Style> getStyleMap()
    {
        return myStyleMap;
    }

    /**
     * Sets <code>Map</code> of style names to <code>Styles</code>.
     * @param styleMap The <code>Map</code> of style names to
     *    <code>Styles</code>.
     * @since 0.5.0
     */
    public void setStyleMap(Map<String, Style> styleMap)
    {
        myStyleMap = styleMap;
    }

    /**
     * Returns a <code>List</code> of template sheet names.
     * @return A <code>List</code> of template sheet names.
     * @since 0.8.0
     */
    public List<String> getTemplateSheetNames()
    {
        return myTemplateSheetNames;
    }

    /**
     * Stores a copy of the given <code>List</code> of template sheet names.
     * @param templateSheetNames A <code>List</code> of template sheet names.
     * @since 0.8.0
     */
    public void setTemplateSheetNames(List<String> templateSheetNames)
    {
        myTemplateSheetNames = new ArrayList<>(templateSheetNames);
    }

    /**
     * Returns a <code>List</code> of sheet names.
     * @return A <code>List</code> of sheet names.
     * @since 0.8.0
     */
    public List<String> getSheetNames()
    {
        return mySheetNames;
    }

    /**
     * Stores a copy of the given <code>List</code> of sheet names.
     * @param sheetNames A <code>List</code> of sheet names.
     * @since 0.8.0
     */
    public void setSheetNames(List<String> sheetNames)
    {
        mySheetNames = new ArrayList<>(sheetNames);
    }

    /**
     * Returns the <code>ExpressionFactory</code>.
     * @return The <code>ExpressionFactory</code>.
     * @since 0.9.0
     */
    public ExpressionFactory getExpressionFactory()
    {
        return myExpressionFactory;
    }

    /**
     * Sets the <code>ExpressionFactory</code>.
     * @param factory The <code>ExpressionFactory</code>.
     * @since 0.9.0
     */
    public void setExpressionFactory(ExpressionFactory factory)
    {
        myExpressionFactory = factory;
    }

    /**
     * Returns a <code>List</code> of beans maps.
     * @return A <code>List</code> of beans maps.
     * @since 0.9.1
     */
    public List<Map<String, Object>> getBeansMaps()
    {
        return myBeansMaps;
    }

    /**
     * Stores a copy of the given <code>List</code> of beans maps.
     * @param beansMaps A <code>List</code> of beans maps.
     * @since 0.9.1
     */
    public void setBeansMaps(List<Map<String, Object>> beansMaps)
    {
        myBeansMaps = new ArrayList<>(beansMaps);
    }
}
