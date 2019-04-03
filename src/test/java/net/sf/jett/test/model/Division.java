package net.sf.jett.test.model;

import java.util.ArrayList;
import java.util.List;

/**
 * A <code>Division</code> represents a division with many teams.
 *
 * @author Randy Gettman
 */
public class Division
{
    private String myName;
    private List<Team> myTeams;

    /**
     * Initializes things to null/empty.
     */
    public Division()
    {
        myName = null;
        myTeams = new ArrayList<>();
    }

    /**
     * Returns the name of the division.
     * @return The name of the division.
     */
    public String getName()
    {
        return myName;
    }

    /**
     * Sets the name of the division.
     * @param name The name of the division.
     */
    public void setName(String name)
    {
        myName = name;
    }

    /**
     * Returns the <code>List</code> of </code>Teams</code>.
     * @return The <code>List</code> of </code>Teams</code>.
     */
    public List<Team> getTeams()
    {
        return myTeams;
    }

    /**
     * Sets the <code>List</code> of </code>Teams</code>.
     * @param teams The <code>List</code> of </code>Teams</code>.
     */
    public void setTeams(List<Team> teams)
    {
        myTeams = teams;
    }

    /**
     * Adds the given <code>Team</code> to the list.
     * @param team The <code>Team</code>.
     */
    public void addTeam(Team team)
    {
        myTeams.add(team);
    }
}

