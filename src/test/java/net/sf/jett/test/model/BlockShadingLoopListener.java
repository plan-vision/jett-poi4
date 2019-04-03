package net.sf.jett.test.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.event.TagLoopEvent;
import net.sf.jett.event.TagLoopListener;
import net.sf.jett.model.Block;

/**
 * A <code>BlockShadingLoopListener</code> is a <code>TagLoopListener</code>
 * that shades alternating blocks light gray.
 *
 * @author Randy Gettman
 * @since 0.3.0
 */
public class BlockShadingLoopListener implements TagLoopListener
{
    /**
     * In D1, sets a new expression.  Prevents B2 from being processed.
     * @param event A <code>TagLoopEvent</code>.
     * @return <code>true</code>.
     * @since 0.8.0
     */
    @Override
    public boolean beforeTagLoopProcessed(TagLoopEvent event)
    {
        Block block = event.getBlock();
        Sheet sheet = event.getSheet();
        int loopIndex = event.getLoopIndex();
        int row = block.getTopRowNum();
        int col = block.getLeftColNum();
        if (loopIndex == 2) // B1 + 2 cols to the right = D1
        {
            Row r = sheet.getRow(row);
            Cell cell = r.getCell(col);
            cell.setCellValue("Three!");
        }
        return !(row == 1 && col == 1); // B2; Don't process this cell.
    }

    /**
     * Shade alternating blocks light gray.
     * @param event The <code>TagLoopEvent</code>.
     */
    @Override
    public void onTagLoopProcessed(TagLoopEvent event)
    {
        Sheet sheet = event.getSheet();
        Workbook workbook = sheet.getWorkbook();
        Block block = event.getBlock();
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        int index = event.getLoopIndex();

        if (index % 2 == 1)
        {
            for (int r = top; r <= bottom; r++)
            {
                Row row = sheet.getRow(r);
                for (int c = left; c <= right; c++)
                {
                    Cell cell = row.getCell(c);
                    CellStyle style = cell.getCellStyle();
                    CellStyle newStyle = workbook.createCellStyle();
                    newStyle.cloneStyleFrom(style);
                    newStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                    newStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(newStyle);
                }
            }
        }
    }
}
