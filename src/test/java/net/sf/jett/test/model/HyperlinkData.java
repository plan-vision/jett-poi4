package net.sf.jett.test.model;

/**
 * A <code>HyperlinkData</code> represents all the data necessary to produce a
 * hyperlink in Excel.
 *
 * @author Randy Gettman
 * @since 0.2.0
 */
public class HyperlinkData
{
    private String myType;
    private String myAddress;
    private String myLabel;

    /**
     * Constructs a <code>HyperlinkData</code> with invalid (null) data.
     */
    public HyperlinkData()
    {
        myType = myAddress = myLabel = null;
    }

    /**
     * Constructs a <code>HyperlinkData</code>.
     * @param type The type of the hyperlink, legal values are given as
     *    constants in <code>HyperlinkTag.java</code>.
     * @param address The address of the hyperlink.
     * @param label The label of the hyperlink.
     */
    public HyperlinkData(String type, String address, String label)
    {
        myType = type;
        myAddress = address;
        myLabel = label;
    }

    /**
     * Returns the hyperlink type.
     * @return The hyperlink type.
     */
    public String getType()
    {
        return myType;
    }

    /**
     * Sets the hyperlink type.
     * @param type The hyperlink type.
     */
    public void setType(String type)
    {
        myType = type;
    }

    /**
     * Returns the hyperlink address.
     * @return The hyperlink address.
     */
    public String getAddress()
    {
        return myAddress;
    }

    /**
     * Sets the hyperlink address.
     * @param address The hyperlink address.
     */
    public void setAddress(String address)
    {
        myAddress = address;
    }

    /**
     * Returns the hyperlink label.
     * @return The hyperlink label.
     */
    public String getLabel()
    {
        return myLabel;
    }

    /**
     * Sets the hyperlink label.
     * @param label The hyperlink label.
     */
    public void setLabel(String label)
    {
        myLabel = label;
    }
}
