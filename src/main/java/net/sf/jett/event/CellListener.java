package net.sf.jett.event;

/**
 * A <code>CellListener</code> is an object that has an opportunity to inspect
 * a <code>Cell</code> as it's being transformed, with access to the
 * <code>Cell</code>, the current <code>Map</code> of beans, and the old and
 * new values for the <code>Cell</code>.
 *
 * @author Randy Gettman
 */
public interface CellListener
{
    /**
     * Called immediately before a <code>Cell</code> is about to be processed.
     * The given <code>CellEvent</code> contains the following related data: a
     * reference to the <code>Cell</code> that was processed, a <code>Map</code>
     * of bean names to bean values that was used, and the old (current) value
     * of the <code>Cell</code>.  The new value of the <code>Cell</code> is not
     * yet available, so it is supplied as <code>null</code>.
     * @param event A <code>SheetEvent</code>.
     * @return A <code>boolean</code> that indicates whether the
     *    <code>Cell</code> should be processed.  <code>true</code> to process
     *    the <code>Cell</code> as normal, and <code>false</code> to skip the
     *    processing of the <code>Cell</code>.  Note that only one
     *    <code>CellListener</code> needs to return <code>false</code> to
     *    prevent the processing of the <code>Cell</code>.
     * @since 0.8.0
     */
    public boolean beforeCellProcessed(CellEvent event);

    /**
     * Called when a <code>Cell</code> has been processed.  The given
     * <code>CellEvent</code> contains the following related data: a reference
     * to the <code>Cell</code> that was processed, a <code>Map</code> of bean
     * names to bean values that was used, the old value of the
     * <code>Cell</code>, and the new value of the <code>Cell</code> after
     * processing.
     *
     * @param event The <code>CellEvent</code>.
     */
    public void cellProcessed(CellEvent event);
}
