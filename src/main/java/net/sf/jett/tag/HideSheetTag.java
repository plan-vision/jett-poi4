package net.sf.jett.tag;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * <p>A <code>HideSheetTag</code> is a <code>BaseHideTag</code> that hides an
 * entire sheet.</p>
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
public class HideSheetTag extends BaseHideTag
{
    /**
    * Returns this <code>Tag's</code> name.
    * @return This <code>Tag's</code> name.
    */
    @Override
    public String getName()
    {
        return "hideSheet";
    }

    /**
     * Hide/show the entire sheet where this tag is located.
     * @param hide Whether to hide or show.
     */
    @Override
    public void setHidden(boolean hide)
    {
        TagContext context = getContext();
        Sheet sheet = context.getSheet();
        Workbook workbook = sheet.getWorkbook();
        int index = workbook.getSheetIndex(sheet);

        workbook.setSheetHidden(index, hide);
    }
}
