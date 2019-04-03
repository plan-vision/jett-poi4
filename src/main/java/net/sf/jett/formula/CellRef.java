package net.sf.jett.formula;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

/**
 * A <code>CellRef</code> is a subclass of the POI class
 * <code>CellReference</code> that provides a correct <code>equals</code>
 * method and adds the concept of a default value.
 *
 * @author Randy Gettman
 */
public class CellRef extends CellReference implements Comparable<CellRef>
{
    /**
     * If no default value is given, this the default default value.
     */
    public static final String DEF_DEFAULT_VALUE = "0";

    /**
     * The token expected in a cell reference with a default value.
     */
    public static final String DEFAULT_VALUE_IND = "||";

    private String myDefaultValue = null;

    public CellRef(String cellRef)
    {
        super(cellRef);
    }

    public CellRef(int pRow, int pCol)
    {
        super(pRow, pCol);
    }

    public CellRef(int pRow, short pCol)
    {
        super(pRow, pCol);
    }

    // The CellReference constructor (Cell) exists in POI 3.7.
    public CellRef(Cell cell)
    {
        super(cell);
    }

    public CellRef(int pRow, int pCol, boolean pAbsRow, boolean pAbsCol)
    {
        super(pRow, pCol, pAbsRow, pAbsCol);
    }

    public CellRef(String pSheetName, int pRow, int pCol, boolean pAbsRow, boolean pAbsCol)
    {
        super(pSheetName, pRow, pCol, pAbsRow, pAbsCol);
    }

    /**
     * Checks whether this cell reference is equal to another object.
     * They are equal if their row and column indexes are equal, the absolute
     * markers are present in the same way, and either their sheet names are
     * equal or they are both <code>null</code>.
     * @param o The other object.
     * @return <code>true</code> if equal, <code>false</code> otherwise.
     */
    @Override
    public boolean equals(Object o)
    {
        if(!(o instanceof CellRef))
            return false;
        CellRef cr = (CellRef) o;
        return getRow() == cr.getRow()
                && getCol() == cr.getCol()
                && isRowAbsolute() == cr.isRowAbsolute()
                && isColAbsolute() == cr.isColAbsolute()
                && ((getSheetName() == null && cr.getSheetName() == null) ||
                (getSheetName() != null && getSheetName().equals(cr.getSheetName())));
    }

    /**
     * Sets the default value.  If this is not called, then the default value
     * itself defaults to <code>null</code>.
     * @param defaultValue The default value.
     */
    public void setDefaultValue(String defaultValue)
    {
        myDefaultValue = defaultValue;
    }

    /**
     * Returns the default value.  If this is not called, then the default value
     * itself defaults to <code>null</code>.
     * @return The default value.
     */
    public String getDefaultValue()
    {
        return myDefaultValue;
    }

    /**
     * Include any default value in the formatted string.
     * @return The formatted string, with a possible default value appended.
     */
    public String formatAsStringWithDef()
    {
        String formatted = formatAsString();
        if (myDefaultValue != null)
            formatted = formatted + DEFAULT_VALUE_IND + myDefaultValue;
        return formatted;
    }

    /**
     * Compares this <code>CellRef</code> to another <code>CellRef</code>.
     * Comparison order: sheet names, row indexes, then column indexes.
     * @param other Another <code>CellRef</code>.
     * @return An integer less than zero, equal to zero, or greater than zero
     *    if this <code>CellRef</code> compares less than, equal to, or greater
     *    than the other <code>CellRef</code>.
     */
    @Override
    public int compareTo(CellRef other)
    {
        String sheetName1 = "", sheetName2 = "";
        if (getSheetName() != null)
            sheetName1 = getSheetName();
        if (other.getSheetName() != null)
            sheetName2 = other.getSheetName();
        int comp = sheetName1.compareTo(sheetName2);
        if (comp != 0)
            return comp;
        comp = getRow() - other.getRow();
        if (comp != 0)
            return comp;
        return getCol() - other.getCol();
    }
}