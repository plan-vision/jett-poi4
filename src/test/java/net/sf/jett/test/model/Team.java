package net.sf.jett.test.model;

/**
 * A <code>Team</code> represents a Team city, name, wins and losses.
 *
 * @author Randy Gettman
 */
public class Team
{
    private String myCity;
    private String myName;
    private int myWins;
    private int myLosses;
    private Division myDivision;

    /**
     * Construct a <code>Team</code>, initializing things to empty/0.
     */
    public Team()
    {
        this(null);
    }

    /**
     * Construct a <code>Team</code>, initializing things to empty/0, but it's a
     * member of the given <code>Division</code>.
     * @param division A <code>Division</code>.
     * @since 0.3.0
     */
    public Team(Division division)
    {
        myCity = "";
        myName = "";
        myWins = 0;
        myLosses = 0;
        myDivision = division;
    }

    /**
     * Returns the city name.
     * @return The city name.
     */
    public String getCity()
    {
        return myCity;
    }

    /**
     * Sets the city name.
     * @param city The city name.
     */
    public void setCity(String city)
    {
        myCity = city;
    }

    /**
     * Returns the team name.
     * @return The team name.
     */
    public String getName()
    {
        return myName;
    }

    /**
     * Sets the team name.
     * @param name The team name.
     */
    public void setName(String name)
    {
        myName = name;
    }

    /**
     * Returns the number of wins.
     * @return The number of wins.
     */
    public int getWins()
    {
        return myWins;
    }

    /**
     * Sets the number of wins.
     * @param wins The number of wins.
     */
    public void setWins(int wins)
    {
        myWins = wins;
    }

    /**
     * Returns the number of losses.
     * @return The number of losses.
     */
    public int getLosses()
    {
        return myLosses;
    }

    /**
     * Sets the number of losses.
     * @param losses The number of losses.
     */
    public void setLosses(int losses)
    {
        myLosses = losses;
    }

    /**
     * Returns the division name.
     * @return The division name.
     * @since 0.3.0
     */
    public String getDivisionName()
    {
        return myDivision.getName();
    }

    /**
     * Returns the number of games above even (0.500).
     * @return The number of games above even (0.500).
     * @since 0.4.0
     */
    public int getNumGamesAboveEven()
    {
        return myWins - myLosses;
    }

    /**
     * Returns the winning percentage.
     * @return The winning percentage, or 0 if wins + losses &lt;= 0.
     */
    public double getPct()
    {
        if (myWins + myLosses <= 0)
            return 0;
        return (double) myWins / ((double) myWins + myLosses);
    }

    /**
     * Returns a dynamic property.  For now, only "division_name" is recognized,
     * although others could be recognized.  This is for dynamic property
     * testing.
     * @param key The key.
     * @return The dynamic property value, or <code>null</code> if the dynamic
     *    property name was not recognized.
     * @since 0.8.0
     */
    public Object get(String key)
    {
        if ("division_name".equals(key))
            return myDivision.getName();

        return null;
    }

    /**
     * The string representation.
     * @return The string representation.
     * @since 0.2.0
     */
    public String toString()
    {
        StringBuilder buf = new StringBuilder();
        buf.append(myCity);
        buf.append(" ");
        buf.append(myName);
        buf.append(" (");
        buf.append(myWins);
        buf.append("-");
        buf.append(myLosses);
        if (myDivision != null)
        {
            buf.append(") -- in the \"");
            buf.append(myDivision.getName());
            buf.append("\" division");
        }
        else
        {
            buf.append(")");
        }
        return buf.toString();

    }
}
