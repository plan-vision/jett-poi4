package net.sf.jett.test.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.event.CellEvent;
import net.sf.jett.event.CellListener;

/**
 * A <code>PopulationCellListener</code> is a <code>CellListener</code> that
 * takes population figures over a certain threshold and bolds the text in
 * that cell.
 *
 * @author Randy Gettman
 */
public class PopulationCellListener implements CellListener
{
    private int myPopThreshold;

    /**
     * Creates a <code>PopulationCellListener</code> that turns text in all
     * <code>Cells</code> that in the template contained the word "population"
     * bold if the resultant number in the <code>Cell</code> is the given
     * population or greater.
     * @param population The population threshold.
     */
    public PopulationCellListener(int population)
    {
        myPopThreshold = population;
    }

    /**
     * Doesn't do anything.
     * @param event A <code>SheetEvent</code>.
     * @return <code>true</code>.
     * @since 0.8.0
     */
    @Override
    public boolean beforeCellProcessed(CellEvent event)
    {
        return true;
    }

    /**
     * Turn cell text with populations over a certain threshold bold!
     *
     * @param event The <code>CellEvent</code>.
     */
    @Override
    public void cellProcessed(CellEvent event)
    {
        Cell cell = event.getCell();
        Object oldValue = event.getOldValue();
        Object newValue = event.getNewValue();

        if (oldValue != null && oldValue.toString().contains("population") &&
                newValue != null && newValue instanceof Number)
        {
            double population = ((Number) newValue).doubleValue();
            if (population >= myPopThreshold)
            {
                Workbook workbook = cell.getSheet().getWorkbook();
                CellStyle style = workbook.createCellStyle();
                style.cloneStyleFrom(cell.getCellStyle());
                int fontIdx = style.getFontIndex();
                Font font = workbook.getFontAt(fontIdx);
                Font boldFont = workbook.findFont(true, font.getColor(), font.getFontHeight(),
                        font.getFontName(), font.getItalic(), font.getStrikeout(), font.getTypeOffset(),
                        font.getUnderline());
                if (boldFont == null)
                {
                    boldFont = workbook.createFont();
                    boldFont.setBold(true);
                    boldFont.setColor(font.getColor());
                    boldFont.setFontHeight(font.getFontHeight());
                    boldFont.setFontName(font.getFontName());
                    boldFont.setItalic(font.getItalic());
                    boldFont.setStrikeout(font.getStrikeout());
                    boldFont.setTypeOffset(font.getTypeOffset());
                    boldFont.setUnderline(font.getUnderline());
                }
                style.setFont(boldFont);
                cell.setCellStyle(style);
            }
        }
    }
}
