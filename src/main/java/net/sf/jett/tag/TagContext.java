package net.sf.jett.tag;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import net.sf.jett.model.Block;

/**
 * A <code>TagContext</code> object represents the context associated with a
 * <code>Tag</code>.
 *
 * @author Randy Gettman
 */
public class TagContext
{
    private Sheet mySheet;
    private Block myBlock;
    private Map<String, Object> myBeans;
    private Map<String, Cell> myProcessedCells;
    private Drawing myDrawing;
    private List<CellRangeAddress> myMergedRegions;
    private List<List<CellRangeAddress>> myConditionalFormattingRegions;
    private Tag myCurrTag;
    private String myFormulaSuffix;

    /**
     * Construct a <code>TagContext</code>, initializing things to null.
     */
    public TagContext()
    {
        mySheet = null;
        myBlock = null;
        myBeans = null;
        myDrawing = null;
        myMergedRegions = null;
        myConditionalFormattingRegions = null;
        myFormulaSuffix = "";
    }

    /**
     * Returns the <code>Sheet</code> on which a tag is found.
     * @return A <code>Sheet</code>.
     */
    public Sheet getSheet()
    {
        return mySheet;
    }

    /**
     * Sets the <code>Sheet</code> on which a tag is found.
     * @param sheet A <code>Sheet</code>.
     */
    public void setSheet(Sheet sheet)
    {
        this.mySheet = sheet;
    }

    /**
     * Returns the <code>Block</code> that applies to a tag.
     * @return A <code>Block</code>.
     */
    public Block getBlock()
    {
        return myBlock;
    }

    /**
     * Sets the <code>Block</code> that applies to a tag.
     * @param block A <code>Block</code>.
     */
    public void setBlock(Block block)
    {
        this.myBlock = block;
    }

    /**
     * Returns the <code>Map</code> of beans.
     * @return A <code>Map</code> of bean names and objects.
     */
    public Map<String, Object> getBeans()
    {
        return myBeans;
    }

    /**
     * Sets the <code>Map</code> of beans.
     * @param beans A <code>Map</code> of bean names and objects.
     */
    public void setBeans(Map<String, Object> beans)
    {
        this.myBeans = beans;
    }

    /**
     * Returns the <code>Map</code> of <code>Cells</code> that have already been
     * processed.
     * @return A <code>Map</code> of <code>Cells</code>.
     */
    public Map<String, Cell> getProcessedCellsMap()
    {
        return myProcessedCells;
    }

    /**
     * Sets the <code>Map</code> of <code>Cells</code> that have already been
     * processed.
     * @param processedCells A <code>Map</code> of <code>Cells</code>.
     */
    public void setProcessedCellsMap(Map<String, Cell> processedCells)
    {
        myProcessedCells = processedCells;
    }

    /**
     * Returns the <code>Sheet's</code> <code>Drawing</code> object, creating it
     * if it doesn't exist.  To avoid clobbering existing drawings, replace a
     * call to this method with a call to <code>getDrawingPatriarch</code> in
     * the POI "ss" package, because that call will NOT corrupt drawings, charts, etc.
     * @return A <code>Drawing</code>.
     * @since 0.2.0
     */
    public Drawing createDrawing()
    {
        if (myDrawing == null)
        {
            myDrawing = mySheet.createDrawingPatriarch();
        }
        return myDrawing;
    }

    /**
     * Returns the <code>Sheet's</code> <code>Drawing</code>.  Creates the
     * <code>Drawing</code> if it doesn't exist yet.
     * @return A <code>Drawing</code>.
     * @since 0.11.0
     */
    public Drawing getOrCreateDrawing()
    {
        if (myDrawing == null)
        {
            myDrawing = getDrawing();
            if (myDrawing == null)
            {
                myDrawing = createDrawing();
            }
        }
        return myDrawing;
    }

    /**
     * Returns the <code>Sheet's</code> <code>Drawing</code> object, if it
     * exists yet.
     * @return A <code>Drawing</code>, or <code>null</code> if it doesn't exist
     *    yet.
     * @since 0.2.0
     */
    public Drawing getDrawing()
    {
        if (myDrawing == null)
        {
            myDrawing = mySheet.getDrawingPatriarch();
        }
        return myDrawing;
    }

    /**
     * Sets the <code>Sheet's</code> <code>Drawing</code> object.  This is
     * usually used to initialize a <code>TagContext</code> from another
     * <code>TagContext</code>, copying the <code>Drawing</code> object.
     * @param drawing A <code>Drawing</code>.
     * @since 0.2.0
     */
    public void setDrawing(Drawing drawing)
    {
        myDrawing = drawing;
    }

    /**
     * Sets the <code>List</code> of <code>CellRangeAddress</code> objects to be
     * manipulated through this <code>TagContext</code>.  All merged region
     * manipulation for a <code>Sheet</code> goes through this list, instead of
     * the <code>Sheet</code> itself, for performance reasons.
     * @param mergedRegions A <code>List</code> of
     *    <code>CellRangeAddress</code>es.
     * @since 0.8.0
     */
    public void setMergedRegions(List<CellRangeAddress> mergedRegions)
    {
        myMergedRegions = mergedRegions;
    }

    /**
     * Returns the <code>List</code> of <code>CellRangeAddress</code> objects on
     * the current <code>Sheet</code>.  For performance reasons, the
     * <code>SheetTransformer</code> reads all merged regions into this list
     * before transformation, all manipulations are done to this list, and after
     * transformation, the list is re-applied to the <code>Sheet</code>.
     * @return A <code>List</code> of <code>CellRangeAddress</code>es.
     * @since 0.8.0
     */
    public List<CellRangeAddress> getMergedRegions()
    {
        return myMergedRegions;
    }

    /**
     * Sets the <code>List</code> of <code>Lists</code> of
     * <code>CellRangeAddress</code> objects to be manipulated through this
     * <code>TagContext</code>.  All conditional formatting range manipulation
     * for a <code>Sheet</code> goes through this list.  Each individual list
     * inside the outer list represents the updated cell range addresses for the
     * <code>ConditionalFormatting</code> at the corresponding index.
     * @param conditionalFormattingRegions A <code>List</code> of <code>Lists</code> of
     *    <code>CellRangeAddress</code>es.
     * @since 0.9.0
     */
    public void setConditionalFormattingRegions(List<List<CellRangeAddress>> conditionalFormattingRegions)
    {
        myConditionalFormattingRegions = conditionalFormattingRegions;
    }

    /**
     * Returns the <code>List</code> of <code>Lists</code> of
     * <code>CellRangeAddress</code> objects to be manipulated through this
     * <code>TagContext</code>.  All conditional formatting range manipulation
     * for a <code>Sheet</code> goes through this list.  Each individual list
     * inside the outer list represents the updated cell range addresses for the
     * <code>ConditionalFormatting</code> at the corresponding index.
     * @return A <code>List</code> of <code>Lists</code> of
     *    <code>CellRangeAddress</code>es.
     * @since 0.9.0
     */
    public List<List<CellRangeAddress>> getConditionalFormattingRegions()
    {
        return myConditionalFormattingRegions;
    }

    /**
     * Returns the current <code>Tag</code> for this context.
     * @return The current <code>Tag</code> for this context.
     * @since 0.9.0
     */
    public Tag getCurrentTag()
    {
        return myCurrTag;
    }

    /**
     * Sets the current <code>Tag</code> for this context.
     * @param tag The current <code>Tag</code> for this context.
     * @since 0.9.0
     */
    public void setCurrentTag(Tag tag)
    {
        myCurrTag = tag;
    }

    /**
     * Returns the formula suffix.  This is useful for formulas, where keeping
     * track of adjusted cell references in all loop levels is vital.
     * @return The formula suffix, e.g. "[1,0][2,1]".  This would
     *    be read as loop 1, iteration 0, subloop 2, iteration 1.
     * @since 0.10.0
     */
    public String getFormulaSuffix()
    {
        return myFormulaSuffix;
    }

    /**
     * Sets the formula suffix.  This is useful for formulas, where keeping
     * track of adjusted cell references in all loop levels is vital.
     * @param formulaSuffix The formula suffix, e.g. "[1,0][2,1]".  This would
     *    be read as loop 1, iteration 0, subloop 2, iteration 1.
     * @since 0.10.0
     */
    public void setFormulaSuffix(String formulaSuffix)
    {
        myFormulaSuffix = formulaSuffix;
    }
}