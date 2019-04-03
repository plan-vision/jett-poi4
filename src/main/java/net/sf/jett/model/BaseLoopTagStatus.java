package net.sf.jett.model;

import net.sf.jett.tag.Tag;

/**
 * A <code>BaseLoopTagStatus</code> represents information about the current iteration
 * of a looping tag.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class BaseLoopTagStatus implements LoopTagStatus
{
    private int myNumIterations;
    private int myCurrIndex;
    private Tag myParentTag;

    /**
     * Constructs a <code>BaseLoopTagStatus</code> with the given number of
     * iterations and a current index of 0.
     * @param tag The parent <code>Tag</code>.  This is only used to protect
     *     the {@link #incrementIndex} method so only the parent tag can call
     *     it, not code in templates.
     * @param numIterations The total number of iterations.
     */
    public BaseLoopTagStatus(Tag tag, int numIterations)
    {
        if (tag == null)
        {
            throw new IllegalArgumentException("The tag parameter must not be null!");
        }
        myParentTag = tag;
        myNumIterations = numIterations;
        myCurrIndex = 0;
    }

    /**
     * Constructs a <code>BaseLoopTagStatus</code> with the given number of
     * iterations and the current iteration number.  Because there is no parent
     * tag, the current iteration number cannot be changed.
     * @param currIteration The current iteration number.  This must be at
     *     least 0 and it must be less than <code>numIterations</code>.
     * @param numIterations The total number of iterations.
     */
    public BaseLoopTagStatus(int currIteration, int numIterations)
    {
        if (currIteration < 0 || currIteration >= numIterations)
        {
            throw new IllegalArgumentException("The current iteration number (" + currIteration +
               ") must be at least 0 and less than the number of iterations (" + numIterations + ")!");
        }
        myParentTag = null;
        myNumIterations = numIterations;
        myCurrIndex = currIteration;
    }

    /**
     * Returns the current index.
     * @return The current index.
     */
    @Override
    public int getIndex()
    {
        return myCurrIndex;
    }

    /**
     * Returns whether the current iteration is the first iteration.
     * @return Whether the current iteration is the first iteration.
     */
    @Override
    public boolean isFirst()
    {
        return myCurrIndex == 0;
    }

    /**
     * Returns whether the current iteration is the last iteration.
     * @return Whether the current iteration is the last iteration.
     */
    @Override
    public boolean isLast()
    {
        return myCurrIndex + 1 == myNumIterations;
    }

    /**
     * Returns the number of iterations.
     * @return The number of iterations.
     */
    public int getNumIterations()
    {
        return myNumIterations;
    }

    /**
     * Increments the current index, if the given tag is this object's parent
     * tag.  This check is so only the <code>Tag</code> that owns this object
     * can change the current iteration number, and not some code in a template.
     * @param tag The parent tag.
     */
    public void incrementIndex(Tag tag)
    {
        if (tag == null || tag != myParentTag)
        {
            throw new IllegalArgumentException("Tag given is not this object's parent tag!");
        }
        myCurrIndex++;
    }
}
