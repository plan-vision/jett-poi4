package net.sf.jett.test.model;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import net.sf.jett.event.SheetEvent;
import net.sf.jett.event.SheetListener;

/**
 * A test <code>SheetListener</code> class to demonstrate the sheet listener
 * functionality in JETT.
 *
 * @author Randy Gettman
 * @since 0.8.0
 */
public class DemoSheetListener implements SheetListener
{
    /**
     * Changes B2 on "First" to "${message2}".  Prevents processing on "Second".
     * @param event A <code>SheetEvent</code>.
     * @return <code>false</code> for Sheet "Second", else <code>true</code>.
     */
    @Override
    public boolean beforeSheetProcessed(SheetEvent event)
    {
        Sheet sheet = event.getSheet();
        String sheetName = sheet.getSheetName();
        Row r = sheet.getRow(1);
        if (r == null)
            r = sheet.createRow(1);
        Cell c = r.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        c.setCellValue("${message2}");      // Which will be evaluated later.
        return !"Second".equals(sheetName); // Don't process the Sheet named "Second".
    }

    /**
     * Changes F1 to "Changed by DemoSheetListener!".
     * @param event The <code>SheetEvent</code>.
     */
    @Override
    public void sheetProcessed(SheetEvent event)
    {
        Sheet sheet = event.getSheet();
        Row r = sheet.getRow(0);
        if (r == null)
            r = sheet.createRow(0);
        Cell c = r.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        c.setCellValue("Changed by DemoSheetListener!");
    }
}
