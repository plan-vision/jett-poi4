package net.sf.jett.test.model;

/**
 * A <code>County</code> represents the county, with name and some statistics.
 *
 * @author Randy Gettman
 */
public class County
{
    private String myName;
    private int myPopulation;
    private double myArea;    // in square kilometers
    private int myEstablishedYear;
    private String myCountySeat;
    private String myFipsCode;
    private String myStateName;

    /**
     * Initializes things to 0/null.
     */
    public County()
    {
        myName = null;
        myPopulation = 0;
        myArea = 0;
        myEstablishedYear = 0;
        myCountySeat = null;
        myFipsCode = null;
        myStateName = null;
    }

    /**
     * Create a <code>County</code> with the given attributes.
     * @param name The county name.
     * @param population The population.
     * @param area The area, in square kilometers.
     * @param establishedYear The year in which it was established.
     * @param countySeat The name of the city that is the county seat.
     * @param fipsCode The FIPS code.
     */
    public County(String name, int population, double area, int establishedYear, String countySeat,
                  String fipsCode)
    {
        myName = name;
        myPopulation = population;
        myArea = area;
        myEstablishedYear = establishedYear;
        myCountySeat = countySeat;
        myFipsCode = fipsCode;
    }

    /**
     * Returns the county name.
     * @return The county name.
     */
    public String getName()
    {
        return myName;
    }

    /**
     * Returns the first letter of the county name.
     * @return The first letter of the county name.
     */
    public String getNameFirstLetter()
    {
        return myName.substring(0, 1);
    }

    /**
     * Sets the county name.
     * @param name The county name.
     */
    public void setName(String name)
    {
        myName = name;
    }

    /**
     * Returns the population.
     * @return The population.
     */
    public int getPopulation()
    {
        return myPopulation;
    }

    /**
     * Sets the population.
     * @param population The population.
     */
    public void setPopulation(int population)
    {
        myPopulation = population;
    }

    /**
     * Returns the area, in square kilometers.
     * @return The area, in square kilometers.
     */
    public double getArea()
    {
        return myArea;
    }

    /**
     * Sets the area, in square kilometers.
     * @param area The area, in square kilometers.
     */
    public void setArea(double area)
    {
        myArea = area;
    }

    /**
     * Returns the established year.
     * @return The established year.
     */
    public int getEstablishedYear()
    {
        return myEstablishedYear;
    }

    /**
     * Sets the established year.
     * @param establishedYear The established year.
     */
    public void setEstablishedYear(int establishedYear)
    {
        myEstablishedYear = establishedYear;
    }

    /**
     * Returns the county seat city.
     * @return The county seat city.
     */
    public String getCountySeat()
    {
        return myCountySeat;
    }

    /**
     * Sets the county seat city.
     * @param countySeat The county seat city.
     */
    public void setCountySeatCity(String countySeat)
    {
        myCountySeat = countySeat;
    }

    /**
     * Returns the FIPS code.
     * @return The FIPS code.
     */
    public String getFipsCode()
    {
        return myFipsCode;
    }

    /**
     * Sets the FIPS code.
     * @param fipsCode The FIPS code.
     */
    public void setFipsCode(String fipsCode)
    {
        myFipsCode = fipsCode;
    }

    /**
     * Returns the population density in people per square kilometer.
     * @return The population density.
     */
    public double getPopulationDensity()
    {
        return (double) myPopulation / myArea;
    }

    /**
     * Sets the state name.
     * @param stateName The state name.
     */
    public void setStateName(String stateName)
    {
        myStateName = stateName;
    }

    /**
     * Returns the state name.
     * @return The state name.
     * @since 0.9.0
     */
    public String getStateName()
    {
        return myStateName;
    }
}
