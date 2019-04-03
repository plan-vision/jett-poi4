package net.sf.jett.event;

import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * A <code>SheetEvent</code> represents data associated with a "sheet
 * processed" event.  It contains a reference to the <code>Sheet</code> that
 * was processed and the <code>Map</code> of bean names and values used to process
 * it.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class SheetEvent
{
    private Sheet mySheet;
    private Map<String, Object> myBeans;

    /**
     * Creates a <code>SheetEvent</code>.
     * @param sheet The <code>Sheet</code> that was processed.
     * @param beans The <code>Map</code> of bean names and values that was used
     *    to process <code>cell</code>.
     */
    public SheetEvent(Sheet sheet, Map<String, Object> beans)
    {
        mySheet = sheet;
        myBeans = beans;
    }

    /**
     * Returns the <code>Sheet</code> that was processed.
     * @return The <code>Sheet</code> that was processed.
     */
    public Sheet getSheet()
    {
        return mySheet;
    }

    /**
     * Returns the <code>Map</code> of bean names and values that was used to
     * process the <code>Sheet</code>.
     * @return The <code>Map</code> of bean names and values.
     */
    public Map<String, Object> getBeans()
    {
        return myBeans;
    }
}
