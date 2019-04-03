package net.sf.jett.event;

import java.util.Map;

import net.sf.jett.model.Block;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * A <code>TagLoopEvent</code> represents data associated with a "tag loop
 * processed" event.  It contains the same data as a <code>TagEvent</code>,
 * plus the current 0-based looping index.
 *
 * @see TagEvent
 * @author Randy Gettman
 * @since 0.3.0
 */
public class TagLoopEvent extends TagEvent
{
    private int myLoopIndex;

    /**
     * Constructs a <code>TagLoopEvent</code> built using the given
     * <code>TagContext</code> and zero-based loop index.
     * @param sheet A <code>Sheet</code>.
     * @param block A <code>Block</code>.
     * @param beans A <code>Map</code> of bean names to values.
     * @param index The loop index.
     */
    public TagLoopEvent(Sheet sheet, Block block, Map<String, Object> beans, int index)
    {
        super(sheet, block, beans);
        myLoopIndex = index;
    }

    /**
     * Returns the current loop index (zero-based).
     * @return The current loop index (zero-based).
     */
    public int getLoopIndex()
    {
        return myLoopIndex;
    }
}
