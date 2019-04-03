package net.sf.jett.formula;

import org.apache.poi.ss.usermodel.Cell;

/**
 * A <code>CellRefRange</code> is a <code>CellRef</code>, representing a range.
 * The inherited attributes represent the upper-left corner of a block of
 * cells.  An additional internal <code>CellRef</code> represents the bottom-
 * right corner of the block of cells.
 *
 * @author Randy Gettman
 */
public class CellRefRange extends CellRef
{
    private CellRef myRangeEndCellRef = null;

    public CellRefRange(String cellRef)
    {
        super(cellRef);
    }

    public CellRefRange(int pRow, int pCol)
    {
        super(pRow, pCol);
    }

    public CellRefRange(int pRow, short pCol)
    {
        super(pRow, pCol);
    }

    public CellRefRange(Cell cell)
    {
        super(cell);
    }

    public CellRefRange(int pRow, int pCol, boolean pAbsRow, boolean pAbsCol)
    {
        super(pRow, pCol, pAbsRow, pAbsCol);
    }

    public CellRefRange(String pSheetName, int pRow, int pCol, boolean pAbsRow, boolean pAbsCol)
    {
        super(pSheetName, pRow, pCol, pAbsRow, pAbsCol);
    }

    /**
     * Sets the end of the cell range with another <code>CellRef</code>.  Copies
     * the given <code>CellRef</code> so it can set its own internal copy's
     * sheet name to <code>null</code>.
     * @param rangeEnd The <code>CellRef</code> indicating the end of the range
     *    of cells.
     */
    public void setRangeEndCellRef(CellRef rangeEnd)
    {
        myRangeEndCellRef = new CellRef(rangeEnd.getRow(), rangeEnd.getCol(),
                rangeEnd.isRowAbsolute(), rangeEnd.isColAbsolute());
    }

    /**
     * Returns the end of the cell range.  It has no sheet name.
     * @return The end of the cell range <code>CellRef</code>.
     */
    public CellRef getRangeEndCellRef()
    {
        return myRangeEndCellRef;
    }

    /**
     * Checks whether this cell reference range is equal to another object.
     * They are equal if:
     * <ul>
     * <li>The other object is also a <code>CellRefRange</code>, the superclass
     *    method <code>equals</code> returns <code>true</code>, and either both
     *    have no range end <code>CellRefs</code>, or both range end
     *    <code>CellRefs</code> compare equal.
     * <li>The other object is a <code>CellRef</code> but not a
     *    <code>CellRefRange</code>, the superclass method <code>equals</code>
     *    returns <code>true</code>, and this object does not have a range end
     *    <code>CellRef</code>.
     * </ul>
     * @param o The other object.
     */
    @Override
    public boolean equals(Object o)
    {
        if (o instanceof CellRef)
        {
            if (o instanceof CellRefRange)
            {
                CellRefRange crr = (CellRefRange) o;
                return (super.equals(o) &&
                        (myRangeEndCellRef == null && crr.myRangeEndCellRef == null) ||
                        (myRangeEndCellRef != null && myRangeEndCellRef.equals(crr.myRangeEndCellRef)));
            }
            else
            {
                return (myRangeEndCellRef == null && super.equals(o));
            }
        }
        return false;
    }

    /**
     * If there is a range end <code>CellRef</code>, then append a colon ":"
     * character followed by the range end formatted string.
     * @return The string representation of the range of cells.
     */
    @Override
    public String formatAsString()
    {
        String superString = super.formatAsString();
        if (myRangeEndCellRef == null)
            return superString;
        else
            return superString + ":" + myRangeEndCellRef.formatAsString();
    }

    // No need to override other methods.
}

