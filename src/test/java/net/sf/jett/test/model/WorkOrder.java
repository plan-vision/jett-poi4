package net.sf.jett.test.model;

/**
 * A <code>WorkOrder</code> is a test class used to test the bug reported in
 * Ticket 57.
 *
 * @author Randy Gettman
 * @since 0.10.0
 */
public class WorkOrder
{
    private String myDepartment;
    private String myLocation;
    private String myInstallation;
    private String myDate;
    private double myJobAmt;
    private double myMatAmt;

    /**
     * Constructs a <code>WorkOrder</code> with the given attributes.
     * @param department The department name.
     * @param location The location.
     * @param installation What is being installed.
     * @param date When the installation takes place.
     * @param jobAmt The cost of the job.
     * @param matAmt The materials cost.
     */
    public WorkOrder(String department, String location, String installation, String date, double jobAmt, double matAmt)
    {
        myDepartment = department;
        myLocation = location;
        myInstallation = installation;
        myDate = date;
        myJobAmt = jobAmt;
        myMatAmt = matAmt;
    }

    /**
     * Returns the department name.
     * @return The department name.
     */
    public String getDepartment()
    {
        return myDepartment;
    }

    /**
     * Returns the location of the job.
     * @return The location of the job.
     */
    public String getLocation()
    {
        return myLocation;
    }

    /**
     * Returns what is being installed.
     * @return What is being installed.
     */
    public String getInstallation()
    {
        return myInstallation;
    }

    /**
     * Returns the date the installation takes place.
     * @return The date the installation takes place.
     */
    public String getDate()
    {
        return myDate;
    }

    /**
     * Returns the cost of the job.
     * @return The cost of the job.
     */
    public double getJobAmt()
    {
        return myJobAmt;
    }

    /**
     * Returns the materials amount.
     * @return The materials amount.
     */
    public double getMatAmt()
    {
        return myMatAmt;
    }

    /**
     * Returns the sum of the job amount and the material amount.
     * @return The sum of the job amount and the material amount.
     */
    public double getTotAmt()
    {
        return myJobAmt + myMatAmt;
    }
}
