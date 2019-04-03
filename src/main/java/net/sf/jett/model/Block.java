package net.sf.jett.model;

import org.apache.poi.ss.usermodel.Cell;

/**
 * A <code>Block</code> object represents a rectangular block of
 * <code>Cells</code>.
 *
 * @author Randy Gettman
 */
public class Block
{
    /**
     * Possible directionalities of the <code>Block</code>.
     */
    public enum Direction
    {
        /**
         * This <code>Block</code> has vertical directionality.
         */
        VERTICAL,
        /**
         * This <code>Block</code> has horizontal directionality.
         */
        HORIZONTAL,
        /**
         * This <code>Block</code> has no directionality.
         */
        NONE
    }
    private int myLeftColNum;
    private int myRightColNum;
    private int myTopRowNum;
    private int myBottomRowNum;
    private Direction myDirection;
    private Block myParent;
    private int myIterationNbr;

    /**
     * Construct a <code>Block</code> that lies in between the given start and
     * end tags.
     * @param parent The <code>Block's</code> parent.
     * @param startTag The <code>Cell</code> containing the start tag.
     * @param endTag The <code>Cell</code> containing the end tag.
     */
    public Block(Block parent, Cell startTag, Cell endTag)
    {
        myTopRowNum = startTag.getRowIndex();
        myBottomRowNum = endTag.getRowIndex();
        myLeftColNum = startTag.getColumnIndex();
        myRightColNum = endTag.getColumnIndex();
        myDirection = Direction.VERTICAL;
        myParent = parent;
        myIterationNbr = 0;
    }

    /**
     * Construct a <code>Block</code> that coincides with the given
     * <code>Cell</code>.  This is used for bodiless tags.
     * @param parent The <code>Block's</code> parent.
     * @param tag The <code>Cell</code> containing the bodiless tag.
     */
    public Block(Block parent, Cell tag)
    {
        // Block area.
        myTopRowNum = tag.getRowIndex();
        myBottomRowNum = tag.getRowIndex();
        myLeftColNum = tag.getColumnIndex();
        myRightColNum = tag.getColumnIndex();
        myDirection = Direction.NONE;
        myParent = parent;
        myIterationNbr = 0;
    }

    /**
     * Construct a <code>Block</code> at the given boundary column and row
     * numbers.
     * @param parent The <code>Block's</code> parent.
     * @param left The left-most column number (0-based).
     * @param right The right-most column index (0-based).
     * @param top The top-most row number (0-based).
     * @param bottom The bottom-most row number (0-based).
     */
    public Block(Block parent, int left, int right, int top, int bottom)
    {
        this(parent, left, right, top, bottom, 0);
    }

    /**
     * Construct a <code>Block</code> at the given boundary column and row
     * numbers, with a 0-based iteration number.
     * @param parent The <code>Block's</code> parent.
     * @param left The left-most column number (0-based).
     * @param right The right-most column index (0-based).
     * @param top The top-most row number (0-based).
     * @param bottom The bottom-most row number (0-based).
     * @param iterationNbr The 0-based iteration number, which will be non-zero
     *                     only if created as part of processing a looping tag.
     * @since 0.11.0
     */
    public Block(Block parent, int left, int right, int top, int bottom, int iterationNbr)
    {
        myLeftColNum = left;
        myRightColNum = right;
        myTopRowNum = top;
        myBottomRowNum = bottom;
        myDirection = Direction.VERTICAL;
        myParent = parent;
        myIterationNbr = iterationNbr;
    }

    /**
     * Returns the left column number (0-based).
     * @return The left column number.
     */
    public int getLeftColNum()
    {
        return myLeftColNum;
    }

    /**
     * Returns the right column number (0-based).
     * @return The right column number.
     */
    public int getRightColNum()
    {
        return myRightColNum;
    }

    /**
     * Returns the top row number (0-based).
     * @return The top row number.
     */
    public int getTopRowNum()
    {
        return myTopRowNum;
    }

    /**
     * Returns the bottom row number (0-based).
     * @return The bottom row number.
     */
    public int getBottomRowNum()
    {
        return myBottomRowNum;
    }

    /**
     * Returns the specific <code>Direction</code> of this <code>Block</code>.
     * @return An enumerated type that determines the directionality of this
     *    <code>Block</code>.
     * @see Direction
     */
    public Direction getDirection()
    {
        return myDirection;
    }

    /**
     * Sets the specific <code>Direction</code> of this <code>Block</code>.
     * @param dir An enumerated type that determines the directionality of this
     *    <code>Block</code>.
     * @see Direction
     */
    public void setDirection(Direction dir)
    {
        myDirection = dir;
    }

    /**
     * Translates the block the given number of columns and rows.
     * @param cols The number of columns to move the block (can be negative).
     * @param rows The number of rows to move the block (can be negative).
     */
    public void translate(int cols, int rows)
    {
        myLeftColNum += cols;
        myRightColNum += cols;
        myTopRowNum += rows;
        myBottomRowNum += rows;
    }

    /**
     * Expands the block the given number of columns and rows.  The "left" and
     * "top" properties are unchanged.  The "bottom" property is changed.
     * @param cols The number of columns to expand the block (can be negative).
     * @param rows The number of rows to expand the block (can be negative).
     */
    public void expand(int cols, int rows)
    {
        myRightColNum += cols;
        myBottomRowNum += rows;
    }

    /**
     * Collapses the block to zero size in columns and rows.
     * @since 0.3.0
     */
    public void collapse()
    {
        myRightColNum = myLeftColNum - 1;
        myBottomRowNum = myTopRowNum - 1;
    }

    /**
     * Returns a <code>String</code> representation of this <code>Block</code>.
     * @return A <code>String</code> representation of this <code>Block</code>.
     */
    public String toString()
    {
        StringBuilder buf = new StringBuilder();
        buf.append("Block: left=");
        buf.append(myLeftColNum);
        buf.append(", right=");
        buf.append(myRightColNum);
        buf.append(", top=");
        buf.append(myTopRowNum);
        buf.append(", bottom=");
        buf.append(myBottomRowNum);
        buf.append(", direction=");
        buf.append(myDirection);
        buf.append(", iteration=");
        buf.append(myIterationNbr);
        return buf.toString();
    }

    /**
     * Returns this <code>Block's</code> parent <code>Block</code>.
     * @return This <code>Block's</code> parent <code>Block</code>.
     */
    public Block getParent()
    {
        return myParent;
    }

    /**
     * Returns this <code>Block's</code> iteration number, which is non-zero
     * only for blocks copied for processing a looping tag.
     * @return The 0-based iteration number, i.e. "I'm twin number 0" or
     *     "I'm twin number 9".
     * @since 0.11.0
     */
    public int getIterationNbr() { return myIterationNbr; }

    /**
     * When this <code>Block</code> is a copy, it may have to react to "sibling"
     * <code>Blocks</code> when the "sibling" is processed first and it grows or
     * shrinks during processing.  Here are the cases:
     * <ol>
     * <li> Vertical direction.  Expand to the right to match the sibling's
     *    column growth.  Translate downward to move out of the way of the
     *    sibling <code>Block</code>.
     * <li> Horizontal direction.  Translate to the right to move out of the way
     *    of the sibling <code>Block</code>.  Expand downward to match the
     *    sibling's row growth.
     * </ol>
     * @param sibling The sibling <code>Block</code> that grew or shrank.
     * @param colGrowth The number of columns the sibling <code>Block</code>
     *    grew (or shrank if negative).
     * @param rowGrowth The number of rows the sibling <code>Block</code> grew
     *    (or shrank if negative).
     */
    public void reactToGrowth(Block sibling, int colGrowth, int rowGrowth)
    {
        switch(myDirection)
        {
        case VERTICAL:
            // Expand right to match original's growth.  Do not shrink!
            if (colGrowth > 0)
            {
                int diff = sibling.myRightColNum - sibling.myLeftColNum - myRightColNum + myLeftColNum;
                if (diff > 0)
                    expand(diff, 0);
            }
            // Move down to get out of the way of the original's growth.
            translate(0, rowGrowth);
            break;
        case HORIZONTAL:
            // Move right to get out of the way of the original's growth.
            translate(colGrowth, 0);
            // Expand down to match original's growth.  Do not shrink!
            if (rowGrowth > 0)
            {
                // Expand this block so it is as tall as the sibling.
                int diff = sibling.myBottomRowNum - sibling.myTopRowNum - myBottomRowNum + myTopRowNum;
                if (diff > 0)
                    expand(0, diff);
            }
            break;
        }
    }
}

