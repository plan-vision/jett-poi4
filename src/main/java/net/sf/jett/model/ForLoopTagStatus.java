package net.sf.jett.model;

import net.sf.jett.tag.Tag;

/**
 * A <code>ForLoopTagStatus</code> is a <code>BaseLoopTagStatus</code> that is
 * also a <code>RangedLoopTagStatus</code>.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class ForLoopTagStatus extends BaseLoopTagStatus implements RangedLoopTagStatus
{
    private int myStart;
    private int myEnd;
    private int myStep;

    /**
     * Constructs a <code>ForLoopTagStatus</code> with the given number of
     * iterations, a current index of 0, and the given start, end, and step
     * values.
     * @param tag The parent <code>Tag</code>.  This is only used to protect
     * the {@link #incrementIndex} method so only the parent tag can call
     *     it, not code in templates.
     * @param numIterations The total number of iterations.
     * @param start The start value of the range.
     * @param end The end value of the range.
     * @param step The step amount of the range.
     */
    public ForLoopTagStatus(Tag tag, int numIterations, int start, int end, int step)
    {
        super(tag, numIterations);
        myStart = start;
        myEnd = end;
        myStep = step;
    }

    /**
     * Returns the starting value of the range.
     * @return The starting value of the range.
     */
    @Override
    public int getStart()
    {
        return myStart;
    }

    /**
     * Returns the ending value of the range.
     * @return The ending value of the range.
     */
    @Override
    public int getEnd()
    {
        return myEnd;
    }

    /**
     * Returns the step amount of the range.
     * @return The step amount of the range.
     */
    @Override
    public int getStep()
    {
        return myStep;
    }
}
