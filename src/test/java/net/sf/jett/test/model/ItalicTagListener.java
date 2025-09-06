package net.sf.jett.test.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.event.TagEvent;
import net.sf.jett.event.TagListener;
import net.sf.jett.model.Block;

/**
 * An <code>ItalicTagListener</code> is a <code>TagListener</code> that turns
 * all text within the block italic.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class ItalicTagListener implements TagListener
{
    /**
     * Doesn't do anything.
     * @param event A <code>TagEvent</code>.
     * @return <code>true</code>.
     * @since 0.8.0
     */
    @Override
    public boolean beforeTagProcessed(TagEvent event)
    {
        return true;
    }

    /**
     * Turns all cell text italic!
     * @param event The <code>TagEvent</code>.
     */
    @Override
    public void onTagProcessed(TagEvent event)
    {
        Block block = event.getBlock();
        Sheet sheet = event.getSheet();
        for (int r = block.getTopRowNum(); r <= block.getBottomRowNum(); r++)
        {
            Row row = sheet.getRow(r);
            if (row != null)
            {
                for (int c = block.getLeftColNum(); c <= block.getRightColNum(); c++)
                {
                    Cell cell = row.getCell(c);
                    if (cell != null)
                    {
                        Workbook workbook = sheet.getWorkbook();
                        CellStyle style = cell.getCellStyle();
                        int fontIdx = style.getFontIndex();
                        Font font = workbook.getFontAt(fontIdx);
                        if (!font.getItalic())
                        {
                            Font italicFont = workbook.findFont(font.getBold(), font.getColor(), font.getFontHeight(),
                                    font.getFontName(), true /*italic*/, font.getStrikeout(), font.getTypeOffset(),
                                    font.getUnderline());
                            CellStyle newStyle = workbook.createCellStyle();
                            newStyle.cloneStyleFrom(style);
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
                            newStyle.setFont(italicFont);
                            cell.setCellStyle(newStyle);
                        }
                    }
                }
            }
        }
    }
}