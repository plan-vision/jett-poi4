package net.sf.jett.model;

import org.apache.poi.ss.util.CellRangeAddress;

/**
 * <p>As it turns out, <code>Workbook#cloneSheet</code> doesn't clone all the
 * properties of a <code>Sheet</code>, specifically missing things such as
 * Print Setup properties.  The coverage is higher for HSSF than for XSSF, but
 * still not complete.</p>
 * <p>In addition, it seems that calling <code>Workbook#setSheetOrder</code>
 * completely messes up XSSF, including nulling out any Repeating Rows and
 * setting other Print Setup properties to defaults.</p>
 * <p>This object stores all such incorrectly cloned sheet properties.  These
 * properties will be read and assigned to this object prior to any cloning.
 * After cloning and sheet moving, these properties will be the source data to
 * put the data back.</p>
 * <p>There are other properties that have to do with Print Setup, but if they
 * are correctly copied, then there is no reason for them to be here.</p>
 *
 * @author Randy Gettman
 * @since 0.7.0
 */
public class MissingCloneSheetProperties
{
    // Directly on Sheet.  These properties are not copied on any cloned Excel
    // spreadsheet, HSSFSheet or XSSFSheet, as of POI 3.10.
    CellRangeAddress myRepeatingColumns;
    // In addition, this somehow gets nulled out on XSSFSheets when
    // setSheetOrder is called, as of POI 3.10.
    CellRangeAddress myRepeatingRows;

    // On the PrintSetup object on the Sheet object.  These properties appear
    // not to be copied on XSSFSheets, but they are copied on HSSFSheets, as of
    // POI 3.10.
    short myCopies;
    boolean amIDraft;
    short myFitHeight;
    short myFitWidth;
    short myHResolution;
    boolean amILandscape;
    boolean amILeftToRight;
    boolean amINoColor;
    boolean amINotes;
    short myPageStart;
    short myPaperSize;
    short myScale;
    boolean amIUsePage;
    boolean amIValidSettings;
    short myVResolution;

    /**
     * Default constructor to set things to default values.
     */
    public MissingCloneSheetProperties()
    {
    }

    /**
     * When a <code>Sheet</code> is cloned, we will of course need to clone the
     * missing clone sheet properties as well.  This is the copy constructor.
     * @param other Another <code>MissingCloneSheetProperties</code>.
     */
    public MissingCloneSheetProperties(MissingCloneSheetProperties other)
    {
        myRepeatingColumns = other.myRepeatingColumns;
        myRepeatingRows = other.myRepeatingRows;
        myCopies = other.myCopies;
        amIDraft = other.amIDraft;
        myFitHeight = other.myFitHeight;
        myFitWidth = other.myFitWidth;
        myHResolution = other.myHResolution;
        amILandscape = other.amILandscape;
        amILeftToRight = other.amILeftToRight;
        amINoColor = other.amINoColor;
        amINotes = other.amINotes;
        myPageStart = other.myPageStart;
        myPaperSize = other.myPaperSize;
        myScale = other.myScale;
        amIUsePage = other.amIUsePage;
        amIValidSettings = other.amIValidSettings;
        myVResolution = other.myVResolution;
    }

    /**
     * Returns the range of columns to repeat at the left of every page.
     * @return The range of columns to repeat at the left of every page.
     */
    public CellRangeAddress getRepeatingColumns()
    {
        return myRepeatingColumns;
    }

    /**
     * Sets the range of columns to repeat at the left of every page.
     * @param repeatingColumns The range of columns to repeat at the left of every page.
     */
    public void setRepeatingColumns(CellRangeAddress repeatingColumns)
    {
        this.myRepeatingColumns = repeatingColumns;
    }

    /**
     * Returns the range of rows to repeat at the top of every page.
     * @return The range of rows to repeat at the top of every page.
     */
    public CellRangeAddress getRepeatingRows()
    {
        return myRepeatingRows;
    }

    /**
     * Sets the range of rows to repeat at the top of every page.
     * @param repeatingRows The range of rows to repeat at the top of every page.
     */
    public void setRepeatingRows(CellRangeAddress repeatingRows)
    {
        this.myRepeatingRows = repeatingRows;
    }

    /**
     * Returns the number of copies.
     * @return The number of copies.
     */
    public short getCopies()
    {
        return myCopies;
    }

    /**
     * Sets the number of copies.
     * @param copies The number of copies.
     */
    public void setCopies(short copies)
    {
        this.myCopies = copies;
    }

    /**
     * Returns whether it's draft quality.
     * @return Whether it's draft quality.
     */
    public boolean isDraft()
    {
        return amIDraft;
    }

    /**
     * Sets whether it's draft quality.
     * @param draft Whether it's draft quality.
     */
    public void setDraft(boolean draft)
    {
        this.amIDraft = draft;
    }

    /**
     * Returns the number of pages tall to fit the sheet.
     * @return The number of pages tall to fit the sheet.
     */
    public short getFitHeight()
    {
        return myFitHeight;
    }

    /**
     * Sets the number of pages tall to fit the sheet.
     * @param fitHeight The number of pages tall to fit the sheet.
     */
    public void setFitHeight(short fitHeight)
    {
        this.myFitHeight = fitHeight;
    }

    /**
     * Returns the number of pages wide to fit the sheet.
     * @return The number of pages wide to fit the sheet.
     */
    public short getFitWidth()
    {
        return myFitWidth;
    }

    /**
     * Sets the number of pages wide to fit the sheet.
     * @param fitWidth The number of pages wide to fit the sheet.
     */
    public void setFitWidth(short fitWidth)
    {
        this.myFitWidth = fitWidth;
    }

    /**
     * Returns the "H Resolution".
     * @return The "H Resolution".
     */
    public short getHResolution()
    {
        return myHResolution;
    }

    /**
     * Sets the "H Resolution".
     * @param hResolution The "H Resolution".
     */
    public void setHResolution(short hResolution)
    {
        this.myHResolution = hResolution;
    }

    /**
     * Returns whether it's landscape.
     * @return Whether it's landscape.
     */
    public boolean isLandscape()
    {
        return amILandscape;
    }

    /**
     * Sets whether it's landscape.
     * @param landscape Whether it's landscape.
     */
    public void setLandscape(boolean landscape)
    {
        this.amILandscape = landscape;
    }

    /**
     * Returns whether the page print order should be left to right before up to down.
     * @return Whether the page print order should be left to right before up to down.
     */
    public boolean isLeftToRight()
    {
        return amILeftToRight;
    }

    /**
     * Sets whether the page print order should be left to right before up to down.
     * @param leftToRight Whether the page print order should be left to right before up to down.
     */
    public void setLeftToRight(boolean leftToRight)
    {
        this.amILeftToRight = leftToRight;
    }

    /**
     * Returns whether to print with no color (b/w).
     * @return Whether to print with no color (b/w).
     */
    public boolean isNoColor()
    {
        return amINoColor;
    }

    /**
     * Sets whether to print with no color (b/w).
     * @param noColor Whether to print with no color (b/w).
     */
    public void setNoColor(boolean noColor)
    {
        this.amINoColor = noColor;
    }

    /**
     * Returns whether to use "notes".
     * @return Whether to use "notes".
     */
    public boolean isNotes()
    {
        return amINotes;
    }

    /**
     * Sets whether to use "notes".
     * @param notes Whether to use "notes".
     */
    public void setNotes(boolean notes)
    {
        this.amINotes = notes;
    }

    /**
     * Returns the starting page number.
     * @return The starting page number.
     */
    public short getPageStart()
    {
        return myPageStart;
    }

    /**
     * Sets the starting page number.
     * @param pageStart The starting page number.
     */
    public void setPageStart(short pageStart)
    {
        this.myPageStart = pageStart;
    }

    /**
     * Returns the paper size.
     * @return The paper size.
     */
    public short getPaperSize()
    {
        return myPaperSize;
    }

    /**
     * Sets the paper size.
     * @param paperSize The paper size.
     */
    public void setPaperSize(short paperSize)
    {
        this.myPaperSize = paperSize;
    }

    /**
     * Returns the scale.
     * @return The scale.
     */
    public short getScale()
    {
        return myScale;
    }

    /**
     * Sets the scale.
     * @param scale The scale.
     */
    public void setScale(short scale)
    {
        this.myScale = scale;
    }

    /**
     * Returns whether to "use page".
     * @return Whether to "use page".
     */
    public boolean isUsePage()
    {
        return amIUsePage;
    }

    /**
     * Sets whether to "use page".
     * @param usePage Whether to "use page".
     */
    public void setUsePage(boolean usePage)
    {
        this.amIUsePage = usePage;
    }

    /**
     * Returns whether the settings are "valid".
     * @return Whether the settings are "valid".
     */
    public boolean isValidSettings()
    {
        return amIValidSettings;
    }

    /**
     * Sets whether the settings are "valid".
     * @param validSettings Whether the settings are "valid".
     */
    public void setValidSettings(boolean validSettings)
    {
        this.amIValidSettings = validSettings;
    }

    /**
     * Returns the "V resolution".
     * @return The "V resolution".
     */
    public short getVResolution()
    {
        return myVResolution;
    }

    /**
     * Sets the "V resolution".
     * @param vResolution The "V resolution".
     */
    public void setVResolution(short vResolution)
    {
        this.myVResolution = vResolution;
    }
}
