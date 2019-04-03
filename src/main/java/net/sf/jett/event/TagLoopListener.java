package net.sf.jett.event;

/**
 * A <code>TagLoopListener</code> is an object that has an opportunity to
 * manipulate a block of <code>Cells</code> as they're being transformed while
 * in a looping tag, with access to the block of <code>Cells</code>, the
 * current <code>Map</code> of beans, and the looping index.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public interface TagLoopListener
{
    /**
     * Called immediately before a block of <code>Cells</code> is about to be
     * processed in a looping tag.  The given <code>TagEvent</code> contains the
     * following related data: a reference to the <code>Block</code> of
     * <code>Cells</code> that is about to be processed and the <code>Map</code>
     * of bean names to bean values that is about to be used.
     * @param event A <code>TagLoopEvent</code>.
     * @return A <code>boolean</code> that indicates whether the
     *    block of <code>Cells</code> should be processed.  <code>true</code> to
     *    process the block as normal, and <code>false</code> to skip the
     *    processing of the block.
     * @since 0.8.0
     */
    public boolean beforeTagLoopProcessed(TagLoopEvent event);


    /**
     * Called when a block of <code>Cells</code> has been processed in a looping
     * tag.  The given <code>TagLoopEvent</code> contains the following related
     * data: the <code>Block</code> of <code>Cells</code> that was processed, a
     * <code>Map</code> of bean names to bean values that was used, and the
     * zero-based looping index.
     *
     * @param event The <code>TagLoopEvent</code>.
     */
    public void onTagLoopProcessed(TagLoopEvent event);
}
