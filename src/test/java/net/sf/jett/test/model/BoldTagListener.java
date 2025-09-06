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
 * A <code>BoldTagListener</code> is a <code>TagListener</code> that turns all
 * text within the block bold.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class BoldTagListener implements TagListener
{
    /**
     * In B1, sets a new expression.  Prevents B2 from being processed.
     * @param event A <code>TagEvent</code>.
     * @return <code>true</code>.
     * @since 0.8.0
     */
    @Override
    public boolean beforeTagProcessed(TagEvent event)
    {
        Sheet sheet = event.getSheet();
        String sheetName = sheet.getSheetName();
        Block block = event.getBlock();
        int row = block.getTopRowNum();
        int col = block.getLeftColNum();
        if ("before".equals(sheetName) && row == 0 && col == 1) // B1
        {
            Row r = sheet.getRow(0);
            Cell cell = r.getCell(1);
            cell.setCellValue("${employees.size()}");
        }
        return !("before".equals(sheetName) && row == 1 && col == 1); // B2; Don't process this cell.
    }

    /**
     * Turns all cell text bold!
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
                        if (!font.getBold())
                        {
                            Font boldFont = workbook.findFont(true, font.getColor(), font.getFontHeight(),
                                    font.getFontName(), font.getItalic(), font.getStrikeout(), font.getTypeOffset(),
                                    font.getUnderline());
                            CellStyle newStyle = workbook.createCellStyle();
                            newStyle.cloneStyleFrom(style);
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
                            newStyle.setFont(boldFont);
                            cell.setCellStyle(newStyle);
                        }
                    }
                }
            }
        }
    }
}
