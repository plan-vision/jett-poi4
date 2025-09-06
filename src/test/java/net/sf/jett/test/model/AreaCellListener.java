package net.sf.jett.test.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.event.CellEvent;
import net.sf.jett.event.CellListener;

/**
 * An <code>AreaCellListener</code> is a <code>CellListener</code> that takes
 * state area figures of 10,000+ and italicizes the text in that cell.
 *
 * @author Randy Gettman
 */
public class AreaCellListener implements CellListener
{
    private int myAreaThreshold;

    /**
     * Creates an <code>AreaCellListener</code> that turns text in all
     * <code>Cells</code> that in the template contained the word "area"
     * italic if the resultant number in the <code>Cell</code> is the given
     * area or greater.
     * @param area The area threshold.
     */
    public AreaCellListener(int area)
    {
        myAreaThreshold = area;
    }

    /**
     * In B1, sets a new expression.  Prevents B2 from being processed.
     * @param event A <code>CellEvent</code>.
     * @return <code>true</code>.
     * @since 0.8.0
     */
    @Override
    public boolean beforeCellProcessed(CellEvent event)
    {
        Cell cell = event.getCell();
        int row = cell.getRowIndex();
        int col = cell.getColumnIndex();
        if (row == 0 && col == 1) // B1
        {
            cell.setCellValue("${california.name}");
        }
        return !(row == 1 && col == 1); // B2; Don't process this cell.
    }

    /**
     * Turn area cells with area over a threshold italic!
     *
     * @param event The <code>CellEvent</code>.
     */
    @Override
    public void cellProcessed(CellEvent event)
    {
        Cell cell = event.getCell();
        Object oldValue = event.getOldValue();
        Object newValue = event.getNewValue();

        if (oldValue != null && oldValue.toString().contains("area") &&
                newValue != null && newValue instanceof Number)
        {
            double population = ((Number) newValue).doubleValue();
            if (population >= myAreaThreshold)
            {
                Workbook workbook = cell.getSheet().getWorkbook();
                CellStyle style = workbook.createCellStyle();
                style.cloneStyleFrom(cell.getCellStyle());
                int fontIdx = style.getFontIndex();
                Font font = workbook.getFontAt(fontIdx);
                Font italicFont = workbook.findFont(font.getBold(), font.getColor(), font.getFontHeight(),
                        font.getFontName(), true, font.getStrikeout(), font.getTypeOffset(),
                        font.getUnderline());
                if (italicFont == null)
                {
                    italicFont = workbook.createFont();
                    italicFont.setBold(font.getBold());
                    italicFont.setColor(font.getColor());
                    italicFont.setFontHeight(font.getFontHeight());
                    italicFont.setFontName(font.getFontName());
                    italicFont.setItalic(true);
                    italicFont.setStrikeout(font.getStrikeout());
                    italicFont.setTypeOffset(font.getTypeOffset());
                    italicFont.setUnderline(font.getUnderline());
                }
                style.setFont(italicFont);
                cell.setCellStyle(style);
            }
        }
    }
}