package net.sf.jett.model;

/**
 * A <code>LoopTagStatus</code> object gives information about the current
 * loop's status.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public interface LoopTagStatus
{
    /**
     * Returns the current 0-based index of the current iteration.
     * @return The current 0-based index of the current iteration.
     */
    public int getIndex();

    /**
     * Returns whether the current iteration is the first iteration.
     * @return Whether the current iteration is the first iteration.
     */
    public boolean isFirst();

    /**
     * Returns whether the current iteration is the last iteration.
     * @return Whether the current iteration is the last iteration.
     */
    public boolean isLast();
}
