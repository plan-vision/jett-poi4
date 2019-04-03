package net.sf.jett.tag;

import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.model.Block;

/**
 * <p>A <code>HideColsTag</code> is a <code>BaseHideTag</code> that hides a
 * range of columns.</p>
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
public class HideColsTag extends BaseHideTag
{
    /**
    * Returns this <code>Tag's</code> name.
    * @return This <code>Tag's</code> name.
    */
    @Override
    public String getName()
    {
        return "hideCols";
    }

    /**
     * Hide/show the columns in this tag's block.
     * @param hide Whether to hide or show.
     */
    @Override
    public void setHidden(boolean hide)
    {
        TagContext context = getContext();
        Block block = context.getBlock();
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        Sheet sheet = context.getSheet();

        for (int c = left; c <= right; c++)
        {
            boolean isHidden = sheet.isColumnHidden(c);
            if (isHidden != hide)
            {
                sheet.setColumnHidden(c, hide);
            }
            if (isHidden && !hide && sheet.getColumnWidth(c) == 0)
            {
                sheet.setColumnWidth(c, 256 * sheet.getDefaultColumnWidth());
            }
        }
    }
}
