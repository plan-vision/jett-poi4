package net.sf.jett.tag;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.model.Block;

/**
 * <p>A <code>HideRowsTag</code> is a <code>BaseHideTag</code> that hides a
 * range of rows.</p>
 *
 * <br>Attributes:
 * <ul>
 * <li><em>Inherits all attributes from {@link BaseTag}.</em></li>
 * <li>test (optional): <code>boolean</code></li>
 * </ul>
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class HideRowsTag extends BaseHideTag
{
    /**
    * Returns this <code>Tag's</code> name.
    * @return This <code>Tag's</code> name.
    */
    @Override
    public String getName()
    {
        return "hideRows";
    }

    /**
     * Hide/show the rows in this tag's block.
     * @param hide Whether to hide or show.
     */
    @Override
    public void setHidden(boolean hide)
    {
        TagContext context = getContext();
        Block block = context.getBlock();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Sheet sheet = context.getSheet();

        for (int r = top; r <= bottom; r++)
        {
            Row row = sheet.getRow(r);
            if (row == null)
                row = sheet.createRow(r);
            row.setZeroHeight(hide);
        }
    }
}
