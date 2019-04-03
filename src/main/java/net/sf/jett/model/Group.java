package net.sf.jett.model;

import java.util.List;

/**
 * A <code>Group</code> is a group of objects that shares some common values
 * for some properties.  One of the objects is designated as "the object" to
 * represent all "items" with the same common values for some properties.  JETT
 * will not be able to determine the type of the objects at compile time, so
 * there's no point to making this class generic.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class Group
{
    private Object myObj;
    private List<?> myItems;

    /**
     * Constructs a <code>Group</code> without a representative object or a list
     * of all items in the group.
     */
    public Group()
    {
        myObj = null;
        myItems = null;
    }

    /**
     * Returns the object that is representative of all objects in the group.
     * @return The object that is representative of all objects in the group.
     */
    public Object getObj()
    {
        return myObj;
    }

    /**
     * Sets the object that is representative of all objects in the group.  This
     * object should be an element in the list of items.
     * @param obj An object that is representative of all objects in the group.
     */
    public void setObj(Object obj)
    {
        myObj = obj;
    }

    /**
     * Returns the <code>List</code> of items in the group.
     * @return The <code>List</code> of items in the group.
     */
    public List<?> getItems()
    {
        return myItems;
    }

    /**
     * Sets the <code>List</code> of items in the group.
     * @param items The <code>List</code> of items in the group.
     */
    public void setItems(List<?> items)
    {
        myItems = items;
    }

    /**
     * Returns the string representation.
     * @return The string representation.
     */
    public String toString()
    {
        StringBuilder buf = new StringBuilder();
        buf.append("Group(");
        buf.append(myObj.toString());
        buf.append(",[");
        for (Object item : myItems)
        {
            buf.append(item.toString());
            buf.append(",");
        }
        return buf.toString();
    }
}
