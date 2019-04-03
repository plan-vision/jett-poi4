package net.sf.jett.test.model;

/**
 * An <code>Employee</code> represents an employee with first and last name,
 * salary, and an optional manager, who is another <code>Employee</code>.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class Employee
{
    private String myFirstName;
    private String myLastName;
    private double mySalary;
    private String myTitle;
    private Employee myManager;
    private String myCatchPhrase;
    private boolean amIManager;

    /**
     * Constructs an <code>Employee</code> with null/0/false attributes.
     */
    public Employee()
    {
        myFirstName = null;
        myLastName = null;
        mySalary = 0;
        myTitle = null;
        myManager = null;
        myCatchPhrase = null;
        amIManager = false;
    }

    /**
     * Constructs an <code>Employee</code> out of an integer seed.  Leaves
     * the manager and isManager properties false/null.
     * @param seed The seed.
     * @since 0.5.0
     */
    public Employee(int seed)
    {
        myFirstName = "First" + seed;
        myLastName = "Last" + seed;
        mySalary = 1000 + seed;
        myTitle = "Title" + (seed / 100);
        myCatchPhrase = "Generic Catch Phrase #" + seed;
    }

    /**
     * Returns the first name.
     * @return The first name.
     */
    public String getFirstName()
    {
        return myFirstName;
    }

    /**
     * Sets the first name.
     * @param firstName The first name.
     */
    public void setFirstName(String firstName)
    {
        myFirstName = firstName;
    }

    /**
     * Returns the last name.
     * @return The last name.
     */
    public String getLastName()
    {
        return myLastName;
    }

    /**
     * Sets the last name.
     * @param lastName The last name.
     */
    public void setLastName(String lastName)
    {
        myLastName = lastName;
    }

    /**
     * Returns the full name, which is the first name, followed by a space, then
     * the last name.
     * @return The full name.
     */
    public String getFullName()
    {
        return myFirstName + " " + myLastName;
    }

    /**
     * Returns the salary.
     * @return The salary.
     */
    public double getSalary()
    {
        return mySalary;
    }

    /**
     * Sets the salary.
     * @param salary The salary.
     */
    public void setSalary(double salary)
    {
        mySalary = salary;
    }

    /**
     * Returns the employee title.
     * @return The employee title.
     */
    public String getTitle()
    {
        return myTitle;
    }

    /**
     * Sets the employee title.
     * @param title The employee title.
     */
    public void setTitle(String title)
    {
        myTitle = title;
    }

    /**
     * Returns the manager, an <code>Employee</code>.
     * @return The manager, an <code>Employee</code>.
     */
    public Employee getManager()
    {
        return myManager;
    }

    /**
     * Sets the manager, an <code>Employee</code>.
     * @param manager The manager, an <code>Employee</code>.
     */
    public void setManager(Employee manager)
    {
        myManager = manager;
    }

    /**
     * Returns the catch phrase.
     * @return The catch phrase.
     */
    public String getCatchPhrase()
    {
        return myCatchPhrase;
    }

    /**
     * Sets the catch phrase.
     * @param catchPhrase The catch phrase.
     */
    public void setCatchPhrase(String catchPhrase)
    {
        myCatchPhrase = catchPhrase;
    }

    /**
     * Returns whether this <code>Employee</code> is a manager.
     * @return Whether this <code>Employee</code> is a manager.
     * @since 0.4.0
     */
    public boolean isAManager()
    {
        return amIManager;
    }

    /**
     * Sets whether this <code>Employee</code> is a manager.
     * @param isAManager Whether this <code>Employee</code> is a manager.
     * @since 0.4.0
     */
    public void setAManager(boolean isAManager)
    {
        amIManager = isAManager;
    }

    /**
     * Returns the name in "Last, First" format.
     * @return The name in "Last, First" format.
     * @since 0.6.0
     */
    @Override
    public String toString()
    {
        return getLastName() + ", " + getFirstName();
    }
}
