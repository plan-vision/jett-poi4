package net.sf.jett.event;

import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.model.Block;

/**
 * A <code>TagEvent</code> represents data associated with a "tag processed"
 * event.  It contains a reference to the <code>Block</code> of
 * <code>Cells</code> that was processed and the <code>Map</code> of bean names
 * to values used to process it.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class TagEvent
{
    private Sheet mySheet;
    private Block myBlock;
    private Map<String, Object> myBeans;

    /**
     * Constructs a <code>TagEvent</code> built using the given
     * <code>TagContext</code>.
     * @param sheet A <code>Sheet</code>.
     * @param block A <code>Block</code>.
     * @param beans A <code>Map</code> of bean names to values.
     */
    public TagEvent(Sheet sheet, Block block, Map<String, Object> beans)
    {
        mySheet = sheet;
        myBlock = block;
        myBeans = beans;
    }

    /**
     * Returns the <code>Sheet</code> on which the block of cells was processed.
     * @return The <code>Sheet</code> on which the block of cells was processed.
     */
    public Sheet getSheet()
    {
        return mySheet;
    }

    /**
     * Returns the <code>Block</code> of cells that was processed.
     * @return The <code>Block</code> of cells that was processed.
     */
    public Block getBlock()
    {
        return myBlock;
    }

    /**
     * Returns the <code>Map</code> of bean names to values used to process the
     * block of cells.
     * @return The <code>Map</code> of bean names to values.
     */
    public Map<String, Object> getBeans()
    {
        return myBeans;
    }
}
