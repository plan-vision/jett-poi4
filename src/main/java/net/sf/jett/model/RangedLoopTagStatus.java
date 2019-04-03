package net.sf.jett.model;

/**
 * A <code>RangedLoopTagStatus</code> is a <code>LoopTagStatus</code> with a
 * range (start/end) and a step (change per iteration).
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public interface RangedLoopTagStatus extends LoopTagStatus
{
    /**
     * Returns the starting value of the range.
     * @return The starting value of the range.
     */
    public int getStart();

    /**
     * Returns the ending value of the range.
     * @return The ending value of the range.
     */
    public int getEnd();

    /**
     * Returns the change per iteration.
     * @return The change per iteration.
     */
    public int getStep();
}
