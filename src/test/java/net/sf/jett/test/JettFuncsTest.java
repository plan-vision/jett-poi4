package net.sf.jett.test;

import org.junit.Test;
import static org.junit.Assert.*;

import net.sf.jett.expression.JettFuncs;

/**
 * This JUnit Test class tests the <code>JettFuncs</code> utility class.
 *
 * @author Randy Gettman
 * @since 0.4.0
 */
public class JettFuncsTest
{
    /**
     * Tests the two-arg version of the "cellRef" method.
     */
    @Test
    public void testCellRef()
    {
        assertEquals("E2", JettFuncs.cellRef(1, 4));
        assertEquals("Z100", JettFuncs.cellRef(99, 25));
    }

    /**
     * Tests the four-arg version of the "cellRef" method.
     */
    @Test
    public void testCellRefRange()
    {
        assertEquals("B2:C4", JettFuncs.cellRef(1, 1, 3, 2));
        assertEquals("K2:K32", JettFuncs.cellRef(1, 10, 31, 1));
    }
}
