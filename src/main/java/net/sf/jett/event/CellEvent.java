package net.sf.jett.event;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;

/**
 * A <code>CellEvent</code> represents data associated with a "cell processed"
 * event.  It contains a reference to the <code>Cell</code> that was processed,
 * the <code>Map</code> of bean names and values used to process it, the old
 * value, and the new value.
 *
 * @author Randy Gettman
 */
public class CellEvent
{
    private Cell myCell;
    private Map<String, Object> myBeans;
    private Object myOldValue;
    private Object myNewValue;

    /**
     * Creates a <code>CellEvent</code>.
     * @param cell The <code>Cell</code> that was processed.
     * @param beans The <code>Map</code> of bean names and values that was used
     *    to process <code>cell</code>.
     * @param oldValue The old cell value.
     * @param newValue The new cell value.
     */
    public CellEvent(Cell cell, Map<String, Object> beans, Object oldValue, Object newValue)
    {
        myCell = cell;
        myBeans = beans;
        myOldValue = oldValue;
        myNewValue = newValue;
    }

    /**
     * Returns the <code>Cell</code> that was processed.
     * @return The <code>Cell</code> that was processed.
     */
    public Cell getCell()
    {
        return myCell;
    }

    /**
     * Returns the <code>Map</code> of bean names and values that was used to
     * process the <code>Cell</code>.
     * @return The <code>Map</code> of bean names and values.
     */
    public Map<String, Object> getBeans()
    {
        return myBeans;
    }

    /**
     * Returns the old cell value.
     * @return The old cell value.
     */
    public Object getOldValue()
    {
        return myOldValue;
    }

    /**
     * Returns the new cell value.
     * @return The new cell value.
     */
    public Object getNewValue()
    {
        return myNewValue;
    }
}
