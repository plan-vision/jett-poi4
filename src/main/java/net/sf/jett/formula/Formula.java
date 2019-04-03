package net.sf.jett.formula;

import java.util.List;

import net.sf.jett.util.FormulaUtil;

/**
 * A <code>Formula</code> represents an Excel-like formula inside "$[" and "]"
 * delimiters.
 *
 * @author Randy Gettman
 */
public class Formula
{
    /**
     * The beginning delimiter for a <code>Formula</code>.
     */
    public static final String BEGIN_FORMULA = "$[";
    /**
     * The ending delimiter for a <code>Formula</code>.
     */
    public static final String END_FORMULA = "]";

    private String myFormulaText;
    private List<CellRef> myCellRefs;

    /**
     * Creates a <code>Formula</code> with the given formula text and the given
     * <code>List</code> of <code>CellRefs</code>.
     * @param formulaText The formula text, as it was entered into the template.
     * @param cellRefs A <code>List</code> of <code>CellRefs</code>.
     */
    public Formula(String formulaText, List<CellRef> cellRefs)
    {
        myFormulaText = FormulaUtil.formatSheetNames(formulaText, cellRefs);
        myCellRefs = cellRefs;
    }

    /**
     * Returns the formula text.
     * @return The formula text.
     */
    public String getFormulaText()
    {
        return myFormulaText;
    }

    /**
     * Returns the <code>List</code> of <code>CellRefs</code>.
     * @return The <code>List</code> of <code>CellRefs</code>.
     */
    public List<CellRef> getCellRefs()
    {
        return myCellRefs;
    }

    /**
     * Returns the string representation.
     * @return The string representation.
     */
    @Override
    public String toString()
    {
        StringBuilder buf = new StringBuilder();
        buf.append("Formula{");
        buf.append(myFormulaText);
        buf.append(",[");
        for (CellRef cellRef : myCellRefs)
        {
            buf.append(cellRef.formatAsString());
            buf.append(",");
        }
        buf.append("]}");
        return buf.toString();
    }
}