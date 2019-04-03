package net.sf.jett.test.model;

import org.junit.Ignore;

/**
 * This class is to test the ability to pass through custom functions to the
 * internal JEXL Engine.
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
@Ignore
public class TestFuncs
{
    /**
     * Just a test answer.
     */
    public static final int THE_ANSWER = 42;

    private static int numCalls = 0;

    /**
     * A test method that keeps track of the number of times it's been called.
     * @return 42
     */
    public int testMethod()
    {
        numCalls++;
        return THE_ANSWER;
    }

    /**
     * Returns the number of times that <code>testMethod</code> has been called.
     * @return The number of times that <code>testMethod</code> has been called.
     */
    public int getCalls()
    {
        return numCalls;
    }

    /**
     * Resets the number of times that <code>testMethod</code> has been called
     * to zero.
     */
    public void resetCalls()
    {
        numCalls = 0;
    }
}
