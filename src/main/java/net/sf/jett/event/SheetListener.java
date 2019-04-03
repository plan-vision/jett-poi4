package net.sf.jett.event;

/**
 * A <code>SheetListener</code> is an object that has an opportunity to inspect
 * a <code>Sheet</code> as it's being transformed, with access to the
 * <code>Sheet</code> and the current <code>Map</code> of beans.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public interface SheetListener
{
    /**
     * Called immediately before a <code>Sheet</code> is about to be processed.
     * The given <code>SheetEvent</code> contains the following related data: a
     * reference to the <code>Sheet</code> that is about to be processed and a
     * <code>Map</code> of bean names to bean values that will be used.
     * @param event A <code>SheetEvent</code>.
     * @return A <code>boolean</code> that indicates whether the
     *    <code>Sheet</code> should be processed.  <code>true</code> to process
     *    the <code>Sheet</code> as normal, and <code>false</code> to skip the
     *    processing of the <code>Sheet</code>.  Note that only one
     *    <code>SheetListener</code> needs to return <code>false</code> to
     *    prevent the processing of the <code>Sheet</code>.
     */
    public boolean beforeSheetProcessed(SheetEvent event);

    /**
     * Called when a <code>Sheet</code> has been processed.  The given
     * <code>SheetEvent</code> contains the following related data: a reference
     * to the <code>Sheet</code> that was processed and a <code>Map</code> of
     * bean names to bean values that was used.
     *
     * @param event The <code>SheetEvent</code>.
     */
    public void sheetProcessed(SheetEvent event);
}