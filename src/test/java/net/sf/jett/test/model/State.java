package net.sf.jett.test.model;

import java.util.ArrayList;
import java.util.List;

/**
 * A <code>State</code> represents a State of the United States of America with
 * many <code>Counties</code>.
 *
 * @author Randy Gettman
 */
public class State
{
    private String myName;
    private List<County> myCounties;

    /**
     * Initializes things to null/empty.
     */
    public State()
    {
        myName = null;
        myCounties = new ArrayList<>();
    }

    /**
     * Returns the name.
     * @return The name.
     */
    public String getName()
    {
        return myName;
    }

    /**
     * Sets the name.
     * @param name The name.
     */
    public void setName(String name)
    {
        myName = name;
    }

    /**
     * Returns the <code>List</code> of <code>Counties</code>.
     * @return The <code>List</code> of <code>Counties</code>.
     */
    public List<County> getCounties()
    {
        return myCounties;
    }

    /**
     * Sets the <code>List</code> of <code>Counties</code>.
     * @param counties The <code>List</code> of <code>Counties</code>.
     */
    public void setCounties(List<County> counties)
    {
        myCounties = counties;
    }

    /**
     * Adds a <code>County</code> to the list.
     * @param county A <code>County</code>.
     */
    public void addCounty(County county)
    {
        myCounties.add(county);
        county.setStateName(myName);
    }
}
