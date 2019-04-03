package net.sf.jett.test.model;

import java.util.List;

/**
 * An <code>Element</code> is a test class used to test the bug reported in
 * Ticket 46.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class Element
{
    private String myTitle;
    private List<Element> mySubElements;

    /**
     * Constructs an <code>Element</code> from the given title and a list
     * of sub-<code>Elements</code>.
     * @param title The title.
     * @param subElements The list of sub-<code>Elements</code>.
     */
    public Element(String title, List<Element> subElements)
    {
        myTitle = title;
        mySubElements = subElements;
    }

    /**
     * Returns the title.
     * @return The title.
     */
    public String getTitle()
    {
        return myTitle;
    }

    /**
     * Returns a <code>List</code> of <code>Elements</code> that are sub-elements.
     * @return A <code>List</code> of <code>Elements</code>.
     */
    public List<Element> getSubElements()
    {
        return mySubElements;
    }
}
