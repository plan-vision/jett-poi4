package net.sf.jett.transform;

import java.util.Map;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.tag.TagContext;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>BlockTransformer</code> knows how to transform a <code>Block</code>
 * of <code>Cells</code> that reside on a <code>Sheet</code> (given with a
 * <code>TagContext</code>).
 *
 * @author Randy Gettman
 */
public class BlockTransformer
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Transforms the given <code>Sheet</code>, using the given <code>Map</code>
     * of bean names to bean objects.
     * @param context Contains the <code>Sheet</code>, the <code>Block</code>,
     *    and the <code>Map</code> of bean names to values.
     * @param workbookContext The <code>WorkbookContext</code>.
     */
    public void transform(TagContext context, WorkbookContext workbookContext)
    {
        transform(context, workbookContext, true);
    }

    /**
     * Transforms the given <code>Sheet</code>, using the given <code>Map</code>
     * of bean names to bean objects.
     * @param context Contains the <code>Sheet</code>, the <code>Block</code>,
     *    and the <code>Map</code> of bean names to values.
     * @param workbookContext The <code>WorkbookContext</code>.
     * @param process Whether to process the <code>Cells</code>; regardless,
     *    they are added to the processed cells <code>Map</code>.
     */
    public void transform(TagContext context, WorkbookContext workbookContext, boolean process)
    {
        Sheet sheet = context.getSheet();
        Block block = context.getBlock();
        Map<String, Cell> processedCells = context.getProcessedCellsMap();
        CellTransformer transformer = new CellTransformer();

        logger.trace("Transforming block: {}", block);

        for (int rowNum = block.getTopRowNum(); rowNum <= block.getBottomRowNum(); rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            if (row != null)
            {
                boolean cellProcessed;
                int startCellNum, endCellNum;
                startCellNum = block.getLeftColNum();
                endCellNum = block.getRightColNum();

                for (int cellNum = startCellNum; cellNum <= endCellNum && rowNum <= block.getBottomRowNum(); cellNum++)
                {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null)
                    {
                        if (process)
                            cellProcessed = transformer.transform(cell, workbookContext, context);
                        else
                        {
                            // Don't process, but place it in the Map to mark it as
                            // processed anyway.
                            cellProcessed = true;
                            processedCells.put(SheetUtil.getCellKey(cell), cell);
                        }
                    }
                    else
                        cellProcessed = true;

                    // It's possible that the block shrank during processing.
                    // Don't run off the Block!
                    endCellNum = block.getRightColNum();

                    if (!cellProcessed)
                    {
                        // Try the cell again next loop.  This may happen if the
                        // transformation resulted in the removal of the block.
                        cellNum--;
                    }
                }
            }
        }

        logger.trace("End: {}", block);
    }
}

