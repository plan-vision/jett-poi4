package net.sf.jett.test.model;

import java.util.List;

/**
 * This class helps test a grid-like summation.
 *
 * @author Randy Gettman
 * @since 0.10.0
 */
public class Region
{
    private String myRegionName;
    private List<Integer> mySalesFigures;

    /**
     * Constructs a <code>Region</code> with the given name and the given list
     * of sales figures.
     * @param name The region's name.
     * @param salesFigures The list of sales figures.
     */
    public Region(String name, List<Integer> salesFigures)
    {
        myRegionName = name;
        mySalesFigures = salesFigures;
    }

    /**
     * Returns the region name.
     * @return The region name.
     */
    public String getName()
    {
        return myRegionName;
    }

    /**
     * Returns a <code>List</code> of sales figures.
     * @return A <code>List</code> of sales figures.
     */
    public List<Integer> getSalesFigures()
    {
        return mySalesFigures;
    }
}