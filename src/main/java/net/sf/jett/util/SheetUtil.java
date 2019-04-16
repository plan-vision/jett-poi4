package net.sf.jett.util;

import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Stack;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
//import org.apache.poi.ss.usermodel.ConditionalFormatting;
//import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
//import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.expression.Expression;
import net.sf.jett.formula.Formula;
import net.sf.jett.model.Block;
import net.sf.jett.model.ExcelColor;
import net.sf.jett.model.PastEndAction;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.tag.Tag;
import net.sf.jett.tag.TagContext;

/**
 * The <code>SheetUtil</code> utility class provides methods for
 * <code>Sheet</code>, <code>Row</code>, and <code>Cell</code> manipulation.
 *
 * @author Randy Gettman
 */
public class SheetUtil
{
    private static final Logger logger = LoggerFactory.getLogger(SheetUtil.class);

    private static final BigDecimal BD_MAX_DOUBLE = new BigDecimal(Double.MAX_VALUE);
    private static final BigDecimal BD_MIN_DOUBLE = new BigDecimal(Double.MIN_NORMAL);
    private static final BigInteger BI_MAX_DOUBLE = BD_MAX_DOUBLE.toBigInteger();

    // Won't catch when begins with a number; test for it when used.
    private static final Pattern POSSIBLE_VARIABLES = Pattern.compile("[A-Za-z0-9_]+");
    // Allow for "variable.n".
    private static final Pattern POSSIBLE_VARIABLES2 = Pattern.compile("[A-Za-z0-9_]+\\.[0-9]+");

    /**
     * Copy only the column widths in the given range of column indexes left by
     * the given number of columns.
     *
     * @param sheet    The <code>Sheet</code> on which to copy column widths.
     * @param colStart The 0-based column index on which to start.
     * @param colEnd   The 0-based column index on which to end.
     * @param numCols  The number of columns to copy column widths left.
     */
    private static void copyColumnWidthsLeft(Sheet sheet, int colStart, int colEnd, int numCols)
    {
        logger.trace("    cCWL: colStart = {}, colEnd = {}, numCols = {}", colStart, colEnd, numCols);
        int newColNum;
        for (int colNum = colStart; colNum <= colEnd; colNum++)
        {
            newColNum = colNum - numCols;
            logger.debug("    Setting column width on col {} to col {}'s width: {}.",
                    newColNum, colNum, sheet.getColumnWidth(colNum));
            sheet.setColumnWidth(newColNum, sheet.getColumnWidth(colNum));
        }
    }

    /**
     * Copy only the column widths in the given range of column indexes right by
     * the given number of columns.
     *
     * @param sheet    The <code>Sheet</code> on which to copy column widths.
     * @param colStart The 0-based column index on which to start.
     * @param colEnd   The 0-based column index on which to end.
     * @param numCols  The number of columns to copy column widths left.
     */
    private static void copyColumnWidthsRight(Sheet sheet, int colStart, int colEnd, int numCols)
    {
        logger.trace("    cCWR: colStart = {}, colEnd = {}, numCols = {}", colStart, colEnd, numCols);
        int newColNum;
        for (int colNum = colEnd; colNum >= colStart; colNum--)
        {
            newColNum = colNum + numCols;
            logger.debug("    Setting column width on col {} to col {}'s width: {}.",
                    newColNum, colNum, sheet.getColumnWidth(colNum));
            sheet.setColumnWidth(newColNum, sheet.getColumnWidth(colNum));
        }
    }

    /**
     * Determine the last populated column and return its 0-based index.
     *
     * @param sheet The <code>Sheet</code> on which to determine the last
     *              populated column.
     * @return The 0-based index of the last populated column (-1 if the
     * <code>Sheet</code> is empty).
     */
    public static int getLastPopulatedColIndex(Sheet sheet)
    {
        int maxCol = -1;
        int lastCol;
        for (Row row : sheet)
        {
            // For some reason, "getLastCellNum()" returns the last cell index "PLUS ONE".
            lastCol = row.getLastCellNum() - 1;
            if (lastCol > maxCol)
                maxCol = lastCol;
        }
        return maxCol;
    }

    /**
     * Copy only the row heights in the given range of row indexes up by
     * the given number of columns.
     *
     * @param sheet    The <code>Sheet</code> on which to copy row heights.
     * @param rowStart The 0-based row index on which to start.
     * @param rowEnd   The 0-based row index on which to end.
     * @param numRows  The number of row to copy row heights up.
     */
    private static void copyRowHeightsUp(Sheet sheet, int rowStart, int rowEnd, int numRows)
    {
        logger.trace("    cRHU: rowStart = {}, rowEnd = {}, numRows = {}", rowStart, rowEnd, numRows);
        int newRowNum;
        Row row, newRow;
        for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++)
        {
            newRowNum = rowNum - numRows;
            row = sheet.getRow(rowNum);
            newRow = sheet.getRow(newRowNum);
            if (row == null && newRow != null)
            {
                newRow.setHeight(sheet.getDefaultRowHeight());  // "standard" height
                logger.debug("      Setting row height on row {} to \"standard\" row height of {}. (Row {} does not exist.)",
                        newRowNum, newRow.getHeight(), rowNum);
            }
            else if (row != null)
            {
                if (newRow == null)
                    newRow = sheet.createRow(newRowNum);
                logger.debug("      Setting row height on row {} to row {}'s height: {}",
                        newRowNum, rowNum, row.getHeight());
                newRow.setHeight(row.getHeight());
            }
        }
    }

    /**
     * Copy only the row heights in the given range of row indexes down by
     * the given number of columns.
     *
     * @param sheet    The <code>Sheet</code> on which to copy row heights.
     * @param rowStart The 0-based row index on which to start.
     * @param rowEnd   The 0-based row index on which to end.
     * @param numRows  The number of row to copy row heights down.
     */
    private static void copyRowHeightsDown(Sheet sheet, int rowStart, int rowEnd, int numRows)
    {
        logger.trace("    cRHD: rowStart = {}, rowEnd = {}, numRows = {}", rowStart, rowEnd, numRows);
        int newRowNum;
        Row row, newRow;
        for (int rowNum = rowEnd; rowNum >= rowStart; rowNum--)
        {
            newRowNum = rowNum + numRows;
            row = sheet.getRow(rowNum);
            newRow = sheet.getRow(newRowNum);
            if (row == null && newRow != null)
            {
                newRow.setHeight(sheet.getDefaultRowHeight());  // "standard" height
                logger.debug("      Setting row height on row {} to \"standard\" row height of {}. (Row {} does not exist.)",
                        newRowNum, newRow.getHeight(), rowNum);
            }
            else if (row != null)
            {
                if (newRow == null)
                    newRow = sheet.createRow(newRowNum);
                logger.debug("      Setting row height on row {} to row {}'s height: {}",
                        newRowNum, rowNum, row.getHeight());
                newRow.setHeight(row.getHeight());
            }
        }
    }

    /**
     * Shift all <code>Cells</code> in the given range of row and column indexes
     * left by the given number of columns.  This will replace any
     * <code>Cells</code> that are "in the way".  Shifts merged regions also.
     *
     * @param sheet           The <code>Sheet</code> on which to move <code>Cells</code>.
     * @param context         A <code>TagContext</code>.
     * @param workbookContext A <code>WorkbookContext</code>.
     * @param colStart        The 0-based column index on which to start moving cells.
     * @param colEnd          The 0-based column index on which to end moving cells.
     * @param rowStart        The 0-based row index on which to start moving cells.
     * @param rowEnd          The 0-based row index on which to end moving cells.
     * @param numCols         The number of columns to move <code>Cells</code> left.
     */
    private static void shiftCellsLeft(Sheet sheet, TagContext context, WorkbookContext workbookContext,
                                       int colStart, int colEnd, int rowStart, int rowEnd, int numCols)
    {
        logger.trace("    Shifting cells left in rows {} to {}, cells {} to {} by {} columns.",
                    rowStart, rowEnd, colStart, colEnd, numCols);
        Row row;
        Cell cell, newCell;
        int newColIndex;
        Map<String, String> tagLocationsMap = workbookContext.getTagLocationsMap();
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)
        {
            row = sheet.getRow(rowIndex);
            if (row != null)
            {
                for (int colIndex = colStart; colIndex <= colEnd; colIndex++)
                {
                    cell = row.getCell(colIndex);
                    newColIndex = colIndex - numCols;
                    newCell = row.getCell(newColIndex);
                    if (cell == null && newCell != null)
                        removeCell(row, newCell);
                    else if (cell != null)
                    {
                        if (newCell == null)
                            newCell = row.createCell(newColIndex);
                        copyCell(cell, newCell);

                        String cellRef = getCellKey(cell);
                        String newCellRef = getCellKey(newCell);
                        String origCellRef = tagLocationsMap.get(cellRef);
                        if (origCellRef != null)
                        {
                            tagLocationsMap.remove(cellRef);
                            tagLocationsMap.put(newCellRef, origCellRef);
                            logger.debug("sCL: Replacing {} => {} with {} => {}",
                                    cellRef, origCellRef, newCellRef, origCellRef);
                        }

                        // Remove the just copied Cell if we detect that it won't be
                        // overwritten by future loops.
                        if (colIndex > colEnd - numCols && colIndex <= colEnd)
                            removeCell(row, cell);
                    }
                }
            }
        }

        shiftMergedRegionsInRange(context, colStart, colEnd, rowStart, rowEnd, -numCols, 0, true, true);
        //shiftConditionalFormattingRegionsInRange(sheet, colStart, colEnd,
        //         rowStart, rowEnd, -numCols, 0);
    }

    /**
     * Shift all <code>Cells</code> in the given range of row and column indexes
     * right by the given number of columns.  This will leave empty
     * <code>Cells</code> behind.  Shifts merged regions also.
     *
     * @param sheet           The <code>Sheet</code> on which to move <code>Cells</code>.
     * @param context         A <code>TagContext</code>.
     * @param workbookContext A <code>WorkbookContext</code>.
     * @param colStart        The 0-based column index on which to start moving cells.
     * @param colEnd          The 0-based column index on which to end moving cells.
     * @param rowStart        The 0-based row index on which to start moving cells.
     * @param rowEnd          The 0-based row index on which to end moving cells.
     * @param numCols         The number of columns to move <code>Cells</code> right.
     */
    private static void shiftCellsRight(Sheet sheet, TagContext context, WorkbookContext workbookContext,
                                        int colStart, int colEnd, int rowStart, int rowEnd, int numCols)
    {
        logger.trace("    Shifting cells right in rows {} to {}, cells {} to {} by {} columns.",
                rowStart, rowEnd, colStart, colEnd, numCols);
        Row row;
        Cell cell, newCell;
        int newColIndex;
        Map<String, String> tagLocationsMap = workbookContext.getTagLocationsMap();
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)
        {
            row = sheet.getRow(rowIndex);
            if (row != null)
            {
                for (int colIndex = colEnd; colIndex >= colStart; colIndex--)
                {
                    cell = row.getCell(colIndex);
                    newColIndex = colIndex + numCols;
                    newCell = row.getCell(newColIndex);
                    if (cell == null && newCell != null)
                        removeCell(row, newCell);
                    else if (cell != null)
                    {
                        if (newCell == null)
                            newCell = row.createCell(newColIndex);
                        copyCell(cell, newCell);

                        String cellRef = getCellKey(cell);
                        String newCellRef = getCellKey(newCell);
                        String origCellRef = tagLocationsMap.get(cellRef);
                        if (origCellRef != null)
                        {
                            tagLocationsMap.remove(cellRef);
                            tagLocationsMap.put(newCellRef, origCellRef);
                            logger.debug("sCR: Replacing {} => {} with {} => {}",
                                    cellRef, origCellRef, newCellRef, origCellRef);
                        }

                        // Remove the just copied Cell if we detect that it won't be
                        // overwritten by future loops.
                        if (colIndex < colStart + numCols && colIndex <= colEnd)
                            removeCell(row, cell);
                    }
                }
            }
        }

        shiftMergedRegionsInRange(context, colStart, colEnd, rowStart, rowEnd, numCols, 0, true, true);
        //shiftConditionalFormattingRegionsInRange(sheet, colStart, colEnd,
        //         rowStart, rowEnd, numCols, 0);
    }

    /**
     * Shift all <code>Cells</code> in the given range of row and column indexes
     * up by the given number of rows.  This will leave empty
     * <code>Cells</code> behind.  Shifts merged regions also.
     *
     * @param sheet           The <code>Sheet</code> on which to move <code>Cells</code>.
     * @param context         A <code>TagContext</code>.
     * @param workbookContext A <code>WorkbookContext</code>.
     * @param colStart        The 0-based column index on which to start moving cells.
     * @param colEnd          The 0-based column index on which to end moving cells.
     * @param rowStart        The 0-based row index on which to start moving cells.
     * @param rowEnd          The 0-based row index on which to end moving cells.
     * @param numRows         The number of columns to move <code>Cells</code> up.
     */
    private static void shiftCellsUp(Sheet sheet, TagContext context, WorkbookContext workbookContext,
                                     int colStart, int colEnd, int rowStart, int rowEnd, int numRows)
    {
        logger.trace("    Shifting cells up in rows {} to {}, cells {} to {} by {} rows.",
                rowStart, rowEnd, colStart, colEnd, numRows);
        int newRowIndex;
        Row oldRow, newRow;
        Cell cell, newCell;
        Map<String, String> tagLocationsMap = workbookContext.getTagLocationsMap();
        for (int colIndex = colStart; colIndex <= colEnd; colIndex++)
        {
            for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++)
            {
                newRowIndex = rowIndex - numRows;
                oldRow = sheet.getRow(rowIndex);
                newRow = sheet.getRow(newRowIndex);
                cell = null;
                if (oldRow != null)
                    cell = oldRow.getCell(colIndex);
                newCell = null;
                if (newRow != null)
                    newCell = newRow.getCell(colIndex);

                if (cell == null && newRow != null && newCell != null)
                    removeCell(newRow, newCell);
                else if (cell != null)
                {
                    if (newRow == null)
                        newRow = sheet.createRow(newRowIndex);
                    if (newCell == null)
                        newCell = newRow.createCell(colIndex);
                    copyCell(cell, newCell);

                    String cellRef = getCellKey(cell);
                    String newCellRef = getCellKey(newCell);
                    String origCellRef = tagLocationsMap.get(cellRef);
                    if (origCellRef != null)
                    {
                        tagLocationsMap.remove(cellRef);
                        tagLocationsMap.put(newCellRef, origCellRef);
                        logger.debug("sCU: Replacing {} => {} with {} => {}",
                                cellRef, origCellRef, newCellRef, origCellRef);
                    }

                    // Remove the just copied Cell if we detect that it won't be
                    // overwritten by future loops.
                    if (rowIndex > rowEnd - numRows && rowIndex <= rowEnd)
                        removeCell(oldRow, cell);
                }
            }
        }

        shiftMergedRegionsInRange(context, colStart, colEnd, rowStart, rowEnd, 0, -numRows, true, true);
        //shiftConditionalFormattingRegionsInRange(sheet, colStart, colEnd,
        //         rowStart, rowEnd, 0, -numRows);
    }

    /**
     * Shift all <code>Cells</code> in the given range of row and column indexes
     * down by the given number of rows.  This will leave empty
     * <code>Cells</code> behind.
     *
     * @param sheet           The <code>Sheet</code> on which to move <code>Cells</code>.
     * @param context         A <code>TagContext</code>.
     * @param workbookContext A <code>WorkbookContext</code>.
     * @param colStart        The 0-based column index on which to start moving cells.
     * @param colEnd          The 0-based column index on which to end moving cells.
     * @param rowStart        The 0-based row index on which to start moving cells.
     * @param rowEnd          The 0-based row index on which to end moving cells.
     * @param numRows         The number of columns to move <code>Cells</code> down.
     */
    private static void shiftCellsDown(Sheet sheet, TagContext context, WorkbookContext workbookContext,
                                       int colStart, int colEnd, int rowStart, int rowEnd, int numRows)
    {
        logger.trace("    Shifting cells down in rows {} to {}, cells {} to {} by {} rows.",
                rowStart, rowEnd, colStart, colEnd, numRows);
        int newRowIndex;
        Row oldRow, newRow;
        Cell cell, newCell;
        Map<String, String> tagLocationsMap = workbookContext.getTagLocationsMap();
        for (int rowIndex = rowEnd; rowIndex >= rowStart; rowIndex--)
        {
            newRowIndex = rowIndex + numRows;
            oldRow = sheet.getRow(rowIndex);
            if (oldRow == null)
                oldRow = sheet.createRow(rowIndex);
            newRow = sheet.getRow(newRowIndex);
            if (newRow == null)
                newRow = sheet.createRow(newRowIndex);
            for (int colIndex = colStart; colIndex <= colEnd; colIndex++)
            {
                cell = oldRow.getCell(colIndex);
                newCell = newRow.getCell(colIndex);
                if (cell == null && newCell != null)
                    removeCell(newRow, newCell);
                else if (cell != null)
                {
                    if (newCell == null)
                        newCell = newRow.createCell(colIndex);
                    copyCell(cell, newCell);

                    String cellRef = getCellKey(cell);
                    String newCellRef = getCellKey(newCell);
                    String origCellRef = tagLocationsMap.get(cellRef);
                    if (origCellRef != null)
                    {
                        tagLocationsMap.remove(cellRef);
                        tagLocationsMap.put(newCellRef, origCellRef);

                        logger.debug("sCD: Replacing {} => {} with {} => {}",
                                cellRef, origCellRef, newCellRef, origCellRef);
                    }
                }

                // Remove the just copied Cell if we detect that it won't be
                // overwritten by future loops.
                if (rowIndex < rowStart + numRows && rowIndex <= rowEnd && cell != null)
                    removeCell(oldRow, cell);
            }
        }

        shiftMergedRegionsInRange(context, colStart, colEnd, rowStart, rowEnd, 0, numRows, true, true);
        //shiftConditionalFormattingRegionsInRange(sheet, colStart, colEnd,
        //         rowStart, rowEnd, 0, numRows);
    }

    /**
     * Removes the given <code>Cell</code> from the given <code>Row</code>.
     * Also removes any <code>Comment</code>.
     *
     * @param row  The <code>Row</code> on which to remove a <code>Cell</code>.
     * @param cell The <code>Cell</code> to remove.
     */
    private static void removeCell(Row row, Cell cell)
    {
        cell.removeCellComment();
        cell.removeHyperlink();
        row.removeCell(cell);
    }

    /**
     * Copy the contents of the old <code>Cell</code> to the new
     * <code>Cell</code>, including borders, cell styles, etc.
     *
     * @param oldCell The source <code>Cell</code>.
     * @param newCell The destination <code>Cell</code>.
     */
    private static void copyCell(Cell oldCell, Cell newCell)
    {
        logger.trace("      cC: oldCell({}) to newCell({})",
            oldCell.getAddress().formatAsString(), newCell.getAddress().formatAsString());

        newCell.setCellStyle(oldCell.getCellStyle());
        Hyperlink h = oldCell.getHyperlink();
        if (h != null)
        {
            CreationHelper helper = oldCell.getSheet().getWorkbook().getCreationHelper();
            Hyperlink hyperlink = helper.createHyperlink(h.getType());
            hyperlink.setAddress(h.getAddress());
            // Insert into sheet.
            newCell.setHyperlink(hyperlink);
        }

        switch (oldCell.getCellType())
        {
        case STRING:
            newCell.setCellValue(oldCell.getRichStringCellValue());
            break;
        case NUMERIC:
            newCell.setCellValue(oldCell.getNumericCellValue());
            break;
        case BLANK:
            newCell.setCellType(CellType.BLANK);
            break;
        case FORMULA:
            newCell.setCellFormula(oldCell.getCellFormula());
            break;
        case BOOLEAN:
            newCell.setCellValue(oldCell.getBooleanCellValue());
            break;
        case ERROR:
            newCell.setCellErrorValue(oldCell.getErrorCellValue());
            break;
        default:
            break;
        }
        // Copy the Comment (if any).
//      Comment comment = oldCell.getCellComment();
//      if (comment != null)
//      {
//         Sheet sheet = newCell.getSheet();
//         Drawing drawing;
//         if (sheet instanceof HSSFSheet)
//         {
//            // The POI documentation warns of corrupting other "drawings" such
//            // as charts and "complex" drawings!!!
//            drawing = ((HSSFSheet) sheet).getDrawingPatriarch();
//            if (drawing == null)
//               drawing = sheet.createDrawingPatriarch();
//         }
//         else if (sheet instanceof XSSFSheet)
//         {
//            drawing = sheet.createDrawingPatriarch();
//         }
//         else
//            throw new IllegalArgumentException("Don't know how to copy a Cell Comment on a " +
//               sheet.getClass().getName());
//         CreationHelper helper = sheet.getWorkbook().getCreationHelper();
//         ClientAnchor newAnchor = helper.createClientAnchor();
//
//         Comment newComment = drawing.createCellComment(newAnchor);
//         newComment.setString(comment.getString());
//         newComment.setAuthor(comment.getAuthor());
//         newCell.setCellComment(newComment);
//      }
    }

    /**
     * Sets the cell value on the given <code>Cell</code> to the given
     * <code>value</code>, regardless of data type.
     *
     * @param context The <code>WorkbookContext</code>; access to the
     *                <code>CellStyleCache</code> and <code>FontCache</code> is used.
     * @param cell    The <code>Cell</code> on which to set the value.
     * @param value   The value.
     * @return The actual value set in the <code>Cell</code>.
     */
    public static Object setCellValue(WorkbookContext context, Cell cell, Object value)
    {
        return setCellValue(context, cell, value, null);
    }

    /**
     * Sets the cell value on the given <code>Cell</code> to the given
     * <code>value</code>, regardless of data type.
     *
     * @param context        The <code>WorkbookContext</code>; access to the
     *                       <code>CellStyleCache</code> and <code>FontCache</code> is used.
     * @param cell           The <code>Cell</code> on which to set the value.
     * @param value          The value.
     * @param origRichString The original <code>RichTextString</code>, to be
     *                       used to set the <code>CellStyle</code> if the value isn't some kind of
     *                       string (<code>String</code> or <code>RichTextString</code>).
     * @return The actual value set in the <code>Cell</code>.
     */
    public static Object setCellValue(WorkbookContext context, Cell cell, Object value, RichTextString origRichString)
    {
        CreationHelper helper = cell.getSheet().getWorkbook().getCreationHelper();
        Object newValue = value;
        boolean applyStyle = true;
        if (value == null)
        {
            newValue = helper.createRichTextString("");
            cell.setCellValue((RichTextString) newValue);
            cell.setCellType(CellType.BLANK);
        }
        else if (value instanceof String)
        {
            newValue = helper.createRichTextString(value.toString());
            cell.setCellValue((RichTextString) newValue);
            applyStyle = false;
        }
        else if (value instanceof RichTextString)
        {
            cell.setCellValue((RichTextString) value);
            applyStyle = false;
        }
        else if (value instanceof Double)
            cell.setCellValue((Double) value);
        else if (value instanceof Integer)
            cell.setCellValue((Integer) value);
        else if (value instanceof Float)
            cell.setCellValue((Float) value);
        else if (value instanceof Long)
            cell.setCellValue((Long) value);
        else if (value instanceof Date)
            cell.setCellValue((Date) value);
        else if (value instanceof Calendar)
            cell.setCellValue((Calendar) value);
        else if (value instanceof Short)
            cell.setCellValue((Short) value);
        else if (value instanceof Byte)
            cell.setCellValue((Byte) value);
        else if (value instanceof Boolean)
            cell.setCellValue((Boolean) value);
        else if (value instanceof BigInteger)
        {
            // Use the double value if it makes sense.
            BigInteger bi = (BigInteger) value;
            BigInteger abs = bi.abs();
            if (abs.compareTo(BI_MAX_DOUBLE) <= 0)
            {
                cell.setCellValue(bi.doubleValue());
            }
            else
            {
                cell.setCellValue(bi.toString());
            }
        }
        else if (value instanceof BigDecimal)
        {
            // Use the double value if it makes sense.
            BigDecimal bd = (BigDecimal) value;
            BigDecimal abs = bd.abs();
            if (abs.compareTo(BigDecimal.ZERO) == 0 ||
                    (abs.compareTo(BD_MIN_DOUBLE) >= 0 && abs.compareTo(BD_MAX_DOUBLE) <= 0))
            {
                cell.setCellValue(bd.doubleValue());
            }
            else
            {
                cell.setCellValue(bd.toString());
            }
        }
        else
        {
            newValue = helper.createRichTextString(value.toString());
            cell.setCellValue((RichTextString) newValue);
            applyStyle = false;
        }
        if (applyStyle)
        {
            RichTextStringUtil.applyFont(origRichString, cell, context.getCellStyleCache(), context.getFontCache());
        }
        return newValue;
    }

    /**
     * Determines whether the <code>Cell</code> on the given <code>Sheet</code>
     * at the given row and column indexes is immaterial: either it doesn't
     * exist, or it exists and the cell type is blank.  That is, whether the
     * cell doesn't exist, is blank, or is empty, and its cell style is the
     * default.
     *
     * @param sheet  The <code>Sheet</code>.
     * @param rowNum The 0-based row index.
     * @param colNum The 0-based column index.
     * @return Whether the <code>Cell</code> is blank.
     * @since 0.9.1
     */
    public static boolean isCellImmaterial(Sheet sheet, int rowNum, int colNum)
    {
        Row r = sheet.getRow(rowNum);
        if (r == null)
            return true;
        Cell c = r.getCell(colNum);
        return (c == null ||
                ((c.getCellType() == CellType.BLANK ||
                        (c.getCellType() == CellType.STRING && "".equals(c.getStringCellValue()))) &&
                        c.getCellStyle().getIndex() == 0
                )
        );
    }

    /**
     * Determines whether the <code>Cell</code> on the given <code>Sheet</code>
     * at the given row and column indexes is blank: either it doesn't exist, or
     * it exists and the cell type is blank.  That is, whether the cell doesn't
     * exist, is blank, or is empty.
     *
     * @param sheet  The <code>Sheet</code>.
     * @param rowNum The 0-based row index.
     * @param colNum The 0-based column index.
     * @return Whether the <code>Cell</code> is blank.
     */
    public static boolean isCellBlank(Sheet sheet, int rowNum, int colNum)
    {
        Row r = sheet.getRow(rowNum);
        if (r == null)
            return true;
        Cell c = r.getCell(colNum);
        return (c == null ||
                c.getCellType() == CellType.BLANK ||
                (c.getCellType() == CellType.STRING && "".equals(c.getStringCellValue())));
    }

    /**
     * Returns a <code>String</code> that can reference the given
     * <code>Cell</code>.
     *
     * @param cell The <code>Cell</code>.
     * @return A string in the format "sheet!A1".
     */
    public static String getCellKey(Cell cell)
    {
        StringBuilder buf = new StringBuilder();
        buf.append(cell.getSheet().getSheetName());
        buf.append("!");
        buf.append(CellReference.convertNumToColString(cell.getColumnIndex()));
        buf.append(cell.getRowIndex() + 1);
        return buf.toString();
    }

    /**
     * Shifts all merged regions found in the given range by the given number
     * of rows and columns (usually one of those two will be zero).
     *
     * @param context A <code>TagContext</code> that supplies the merged regions
     *                to shift.
     * @param left    The 0-based index of the column on which to start shifting
     *                merged regions.
     * @param right   The 0-based index of the column on which to end shifting
     *                merged regions.
     * @param top     The 0-based index of the row on which to start shifting
     *                merged regions.
     * @param bottom  The 0-based index of the row on which to end shifting
     *                merged regions.
     * @param numCols The number of columns to shift the merged region (can be
     *                negative).
     * @param numRows The number of rows to shift the merged region (can be
     *                negative).
     * @param remove  Determines whether to remove the old merged region,
     *                resulting in a shift, or not to remove the old merged region,
     *                resulting in a copy.
     * @param add     Determines whether to add the new merged region, resulting in
     *                a copy, or not to add the new merged region, resulting in a shift.
     */
    private static void shiftMergedRegionsInRange(TagContext context,
                                                  int left, int right, int top, int bottom, int numCols, int numRows,
                                                  boolean remove, boolean add)
    {
        logger.trace("    sMRIR: left {}, right {}, top {}, bottom {}, numCols {}, numRows {}, remove {}, add {}",
                left, right, top, bottom, numCols, numRows, remove, add);
        if (numCols == 0 && numRows == 0 && remove && add)
            return;

        List<CellRangeAddress> sheetMergedRegions = context.getMergedRegions();
        if (add)
        {
            int numMergedRegions = sheetMergedRegions.size();
            for (int i = 0; i < numMergedRegions; i++)
            {
                CellRangeAddress region = sheetMergedRegions.get(i);
                if (isCellAddressWhollyContained(region, left, right, top, bottom))
                {
                    CellRangeAddress newRegion = new CellRangeAddress(
                            region.getFirstRow() + numRows,
                            region.getLastRow() + numRows,
                            region.getFirstColumn() + numCols,
                            region.getLastColumn() + numCols);
                    if (remove)
                    {
                        logger.debug("      Updating adjusted merged region from {} to {}.", region, newRegion);
                        sheetMergedRegions.set(i, newRegion);
                    }
                    else
                    {
                        logger.debug("      Copying merged region from {} to {}.", region, newRegion);
                        sheetMergedRegions.add(newRegion);
                    }
                }
            }
        }
        else if (remove)
        {
            int numMergedRegions = sheetMergedRegions.size();
            List<CellRangeAddress> regionsToRemove = new ArrayList<>();
            for (int i = 0; i < numMergedRegions; i++)
            for (CellRangeAddress region : sheetMergedRegions)
            {
                if (isCellAddressWhollyContained(region, left, right, top, bottom))
                {
                    logger.debug("      Removing merged region: {}");
                    regionsToRemove.add(region);
                }
            }
            if (!regionsToRemove.isEmpty())
            {
                sheetMergedRegions.removeAll(regionsToRemove);
            }
        }
    }

    // TODO: Decide whether to even do this in this build.

//   /**
//    * Removes all conditional formatting regions found in the given range.
//    * @param sheet The <code>Sheet</code> on which to remove conditional
//    *    formatting regions.
//    * @param left The 0-based index of the column on which to start removing
//    *    conditional formatting regions.
//    * @param right The 0-based index of the column on which to end removing
//    *    conditional formatting regions.
//    * @param top The 0-based index of the row on which to start removing
//    *    conditional formatting regions.
//    * @param bottom The 0-based index of the row on which to end removing
//    *    conditional formatting regions.
//    * @since 0.7.0
//    */
//   private static void removeConditionalFormattingRegionsInRange(Sheet sheet,
//      int left, int right, int top, int bottom)
//   {
//      logger.trace("    rCFRIR: left {}, right {}, top {}, bottom {}",
//            left, right, top, bottom);
//
//      SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
//      int numConditionalFormattings = scf.getNumConditionalFormattings();
//      for (int i = 0; i < numConditionalFormattings; i++)
//      {
//         ConditionalFormatting cf = scf.getConditionalFormattingAt(i);
//         CellRangeAddress[] regions = cf.getFormattingRanges();
//         boolean[] keep = new boolean[regions.length];
//         int numToKeep = 0;
//         for (int j = 0; j < regions.length; j++)
//         {
//            CellRangeAddress region = regions[j];
//            if (!isCellAddressWhollyContained(region, left, right, top, bottom))
//            {
//               keep[j] = true;
//               numToKeep++;
//            }
//            if (!keep[j])
//            {
//               logger.debug("      Removing Conditional Formatting at region: {}");
//            }
//         }
//
//         if (numToKeep < regions.length)
//         {
//            ConditionalFormattingRule[] rules = null;
//            CellRangeAddress[] newRegions = null;
//
//            // Only bother to extract the current data if we are keeping at
//            // least one CellRangeAddress.
//            if (numToKeep > 0)
//            {
//               int numRules = cf.getNumberOfRules();
//               rules = new ConditionalFormattingRule[numRules];
//               for (int r = 0; r < numRules; r++)
//               {
//                  rules[r] = cf.getRule(r);
//               }
//
//               newRegions = new CellRangeAddress[numToKeep];
//               int idx = 0;
//               for (int j = 0; j < regions.length; j++)
//               {
//                  if (keep[j])
//                  {
//                     newRegions[idx++] = regions[j];
//                  }
//               }
//            }
//
//            // Either way, remove the current one.
//            scf.removeConditionalFormatting(i);
//            // Only if we need to keep at least one region.
//            if (numToKeep > 0)
//            {
//               scf.addConditionalFormatting(newRegions, rules);
//            }
//
//            // Either way, one less to look at.
//            numConditionalFormattings--;
//            // Look at current index again next loop.
//            i--;
//         }
//         // else no action to take.
//      }  // end for loop over ConditionalFormattings
//   }
//
//   /**
//    * Shifts all conditional formatting regions found in the given range by the
//    * given number of rows and columns (usually one of those two will be zero).
//    * @param sheet The <code>Sheet</code> on which to shift conditional
//    *    formatting regions.
//    * @param left The 0-based index of the column on which to start shifting
//    *    conditional formatting regions.
//    * @param right The 0-based index of the column on which to end shifting
//    *    conditional formatting regions.
//    * @param top The 0-based index of the row on which to start shifting
//    *    conditional formatting regions.
//    * @param bottom The 0-based index of the row on which to end shifting
//    *    conditional formatting regions.
//    * @param numCols The number of columns to shift the conditional formatting
//    *    region (can be negative).
//    * @param numRows The number of rows to shift the conditional formatting
//    *    region (can be negative).
//    * @since 0.7.0
//    */
//   private static void shiftConditionalFormattingRegionsInRange(Sheet sheet,
//      int left, int right, int top, int bottom, int numCols, int numRows)
//   {
//      logger.trace("    sCFRIR: left {}, right {}, top {}, bottom {}, numCols {}, numRows {}",
//            left, right, top, bottom, numCols, numRows);
//      if (numCols == 0 && numRows == 0)
//         return;
//
//      SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
//      for (int i = 0; i < scf.getNumConditionalFormattings(); i++)
//      {
//         ConditionalFormatting cf = scf.getConditionalFormattingAt(i);
//         CellRangeAddress[] regions = cf.getFormattingRanges();
//         for (int j = 0; j < regions.length; j++)
//         {
//            CellRangeAddress region = regions[j];
//            if (isCellAddressWhollyContained(region, left, right, top, bottom))
//            {
//               // Replace the region in the existing array with a new, shifted
//               // region.
//               int firstCol = region.getFirstColumn() + numCols;
//               int firstRow = region.getFirstRow() + numRows;
//               int lastCol = region.getLastColumn() + numCols;
//               int lastRow = region.getLastRow() + numRows;
//               CellRangeAddress shifted = new CellRangeAddress(
//                   firstRow, lastRow, firstCol, lastCol);
//
//               logger.debug("      Shifting Conditional Formatting at region: {} to: {}", region, shifted);
//
//               regions[j] = shifted;
//            }
//         }
//      }
//   }
//
//   /**
//    * Copies all conditional formatting regions found in the given range by the
//    * given number of rows and columns (usually one of those two will be zero).
//    * @param sheet The <code>Sheet</code> on which to copy conditional
//    *    formatting regions.
//    * @param left The 0-based index of the column on which to start copying
//    *    conditional formatting regions.
//    * @param right The 0-based index of the column on which to end copying
//    *    conditional formatting regions.
//    * @param top The 0-based index of the row on which to start copying
//    *    conditional formatting regions.
//    * @param bottom The 0-based index of the row on which to end copying
//    *    conditional formatting regions.
//    * @param numCols The number of columns to copy the conditional formatting
//    *    region (can be negative).
//    * @param numRows The number of rows to copy the conditional formatting
//    *    region (can be negative).
//    * @since 0.7.0
//    */
//   private static void copyConditionalFormattingRegionsInRange(Sheet sheet,
//      int left, int right, int top, int bottom, int numCols, int numRows)
//   {
//      logger.trace("    cCFRIR: left {}, right {}, top {}, bottom {}, numCols {}, numRows {}",
//            left, right, top, bottom, numCols, numRows);
//      if (numCols == 0 && numRows == 0)
//         return;
//
//      SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
//      int numConditionalFormattings = scf.getNumConditionalFormattings();
//      for (int i = 0; i < numConditionalFormattings; i++)
//      {
//         ConditionalFormatting cf = scf.getConditionalFormattingAt(i);
//         CellRangeAddress[] regions = cf.getFormattingRanges();
//         List<CellRangeAddress> newRegionsList = new ArrayList<CellRangeAddress>();
//         for (CellRangeAddress region : regions)
//         {
//            if (isCellAddressWhollyContained(region, left, right, top, bottom))
//            {
//               // Replace the region in the existing array with a new, shifted
//               // region.
//               int firstCol = region.getFirstColumn() + numCols;
//               int firstRow = region.getFirstRow() + numRows;
//               int lastCol = region.getLastColumn() + numCols;
//               int lastRow = region.getLastRow() + numRows;
//               CellRangeAddress copy = new CellRangeAddress(
//                   firstRow, lastRow, firstCol, lastCol);
//
//               newRegionsList.add(copy);
//
//               logger.debug("      Copying Conditional Formatting at region: {} to {}", region, copy);
//            }
//         }
//
//         // Only remove/add with more regions if any copies were made.
//         if (newRegionsList.size() > 0)
//         {
//            int numRules = cf.getNumberOfRules();
//            ConditionalFormattingRule[] rules = new ConditionalFormattingRule[numRules];
//            for (int r = 0; r < numRules; r++)
//            {
//               rules[r] = cf.getRule(r);
//            }
//
//            CellRangeAddress[] newRegions = new CellRangeAddress[regions.length + newRegionsList.size()];
//            System.arraycopy(regions, 0, newRegions, 0, regions.length);
//            System.arraycopy(newRegionsList.toArray(new CellRangeAddress[newRegionsList.size()]),
//                    0, newRegions, regions.length, newRegionsList.size());
//
//            scf.removeConditionalFormatting(i);
//            scf.addConditionalFormatting(newRegions, rules);
//
//            // Don't need to look at this one again at the end of the loop.
//            numConditionalFormattings--;
//            // Look at current index again next loop.
//            i--;
//         }
//      }
//   }

    // TODO: End of Decide whether to even do this in this build.

    /**
     * Determines whether the given <code>CellRangeAddress</code>, representing
     * a merged region, is wholly contained in the given area of
     * <code>Cells</code>.  If <code>left</code> &gt;= <code>right</code>, then
     * this will search the entire row(s).
     *
     * @param mergedRegion The <code>CellRangeAddress</code> merged region.
     * @param left         The 0-based column index on which to start searching for
     *                     merged regions.
     * @param right        The 0-based column index on which to stop searching for
     *                     merged regions.
     * @param top          The 0-based row index on which to start searching for
     *                     merged regions.
     * @param bottom       The 0-based row index on which to stop searching for
     *                     merged regions.
     * @return <code>true</code> if wholly contained, <code>false</code>
     * otherwise.
     */
    private static boolean isCellAddressWhollyContained(CellRangeAddress mergedRegion,
                                                        int left, int right, int top, int bottom)
    {
        return (mergedRegion.getFirstRow() >= top && mergedRegion.getLastRow() <= bottom &&
                mergedRegion.getFirstColumn() >= left && mergedRegion.getLastColumn() <= right);
    }

    /**
     * Removes all <code>Cells</code> found inside the given <code>Block</code>
     * on the given <code>Sheet</code>.
     *
     * @param sheet      The <code>Sheet</code> on which to delete a
     *                   <code>Block</code>.
     * @param tagContext A <code>TagContext</code>.
     * @param block      The <code>Block</code> of <code>Cells</code> to delete.
     * @param context    The <code>WorkbookContext</code>.
     */
    public static void deleteBlock(Sheet sheet, TagContext tagContext, Block block, WorkbookContext context)
    {
        logger.trace("  deleteBlock: {}: {}.", sheet.getSheetName(), block);
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Map<String, String> tagLocationsMap = context.getTagLocationsMap();

        // Blank out the Cells.
        for (int rowNum = top; rowNum <= bottom; rowNum++)
        {
            Row r = sheet.getRow(rowNum);
            if (r != null)
            {
                for (int cellNum = left; cellNum <= right; cellNum++)
                {
                    Cell c = r.getCell(cellNum);
                    if (c != null)
                    {
                        String cellRef = getCellKey(c);
                        removeCell(r, c);
                        tagLocationsMap.remove(cellRef);

                        logger.debug("dB: Removing {}", cellRef);
                    }
                }
            }
        }
        // Remove any merged regions in this Block.
        shiftMergedRegionsInRange(tagContext, left, right, top, bottom, 0, 0, true, false);
        // Remove any conditional formatting regions in this Block.
        //removeConditionalFormattingRegionsInRange(sheet, left, right, top, bottom);
        // Lose the current cell references.
        FormulaUtil.shiftCellReferencesInRange(sheet.getSheetName(), context,
                left, right, top, bottom,
                0, 0, true, false);
    }

    /**
     * Blanks out all <code>Cells</code> found inside the given
     * <code>Block</code> on the given <code>Sheet</code>.
     *
     * @param sheet   The <code>Sheet</code> on which to clear a
     *                <code>Block</code>
     * @param block   The <code>Block</code> of <code>Cells</code> to clear.
     * @param context The <code>WorkbookContext</code>.
     */
    public static void clearBlock(Sheet sheet, Block block, WorkbookContext context)
    {
        logger.trace("  clearBlock: {}: {}.", sheet.getSheetName(), block);
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Map<String, String> tagLocationsMap = context.getTagLocationsMap();

        // Blank out the Cells.
        for (int rowNum = top; rowNum <= bottom; rowNum++)
        {
            Row r = sheet.getRow(rowNum);
            if (r != null)
            {
                for (int cellNum = left; cellNum <= right; cellNum++)
                {
                    Cell c = r.getCell(cellNum);
                    if (c != null)
                    {
                        String cellRef = getCellKey(c);
                        c.setCellType(CellType.BLANK);
                        c.removeHyperlink();
                        tagLocationsMap.remove(cellRef);
                        logger.debug("cB: Removing {}", cellRef);
                    }
                }
            }
        }
        // Lose the current cell references.
        FormulaUtil.shiftCellReferencesInRange(sheet.getSheetName(), context,
                left, right, top, bottom,
                0, 0, true, false);
    }

    /**
     * Takes the "replace value" <code>PastEndAction</code> on the entire
     * <code>Sheet</code> - sheet name, header/footer, and all <code>Cells</code>
     * on it.
     *
     * @param sheet            The <code>Sheet</code> on which to replace expressions.
     * @param pastEndRefs      A <code>List</code> of strings identifying which
     *                         expressions represent collection access beyond the end of the
     *                         collection.
     * @param replacementValue The value with which to replace those expressions.
     * @return The sheet name after replacement.  It may change because the
     * sheet name can be changed by this method.
     * @since 0.9.1
     */
    public static String takePastEndAction(Sheet sheet, List<String> pastEndRefs, String replacementValue)
    {
        Workbook workbook = sheet.getWorkbook();
        String sheetName = sheet.getSheetName();
        String newSheetName = replaceValue(sheetName, pastEndRefs, replacementValue);
        if (!sheetName.equals(newSheetName))
        {
            newSheetName = SheetUtil.safeSetSheetName(workbook, workbook.getSheetIndex(sheet), newSheetName);
        }

        Header header = sheet.getHeader();
        String text = header.getLeft();
        String result = replaceValue(text, pastEndRefs, replacementValue);
        header.setLeft(result);
        text = header.getCenter();
        result = replaceValue(text, pastEndRefs, replacementValue);
        header.setCenter(result);
        text = header.getRight();
        result = replaceValue(text, pastEndRefs, replacementValue);
        header.setRight(result);

        Footer footer = sheet.getFooter();
        text = footer.getLeft();
        result = replaceValue(text, pastEndRefs, replacementValue);
        footer.setLeft(result);
        text = footer.getCenter();
        result = replaceValue(text, pastEndRefs, replacementValue);
        footer.setCenter(result);
        text = footer.getRight();
        result = replaceValue(text, pastEndRefs, replacementValue);
        footer.setRight(result);

        Block block = new Block(null, 0, SheetUtil.getLastPopulatedColIndex(sheet), 0, sheet.getLastRowNum());
        block.setDirection(Block.Direction.NONE);

        SheetUtil.takePastEndAction(sheet, block, pastEndRefs, PastEndAction.REPLACE_EXPR, replacementValue);

        return newSheetName;
    }

    /**
     * Helper method to perform the "replace value" past end action on a string.
     *
     * @param text             The text to change.
     * @param pastEndRefs      A <code>List</code> of strings identifying which
     *                         expressions represent collection access beyond the end of the
     *                         collection.
     * @param replacementValue The value with which to replace those expressions.
     * @return The modified string.
     * @since 0.9.1
     */
    private static String replaceValue(String text, List<String> pastEndRefs, String replacementValue)
    {
        int exprBegin = text.indexOf(Expression.BEGIN_EXPR);
        int exprEnd = text.indexOf(Expression.END_EXPR);
        logger.trace("  rV: exprBegin = {}, exprEnd = {}", exprBegin, exprEnd);

        while (exprBegin != -1 && exprEnd != -1 && exprEnd > exprBegin)
        {
            String expression = text.substring(exprBegin + Expression.BEGIN_EXPR.length(), exprEnd - 1 + Expression.END_EXPR.length());
            boolean replaceExpr = containsPastEndRef(expression, pastEndRefs);

            if (replaceExpr)
            {
                String remove = text.substring(exprBegin, exprEnd + Expression.END_EXPR.length());

                logger.debug("    removing \"{}\".", remove);
                text = text.replaceAll(Pattern.quote(remove), replacementValue);
                logger.debug("    text is now \"{}\".", text);
                exprBegin = text.indexOf(Expression.BEGIN_EXPR, exprBegin);
            }
            else
            {
                exprBegin = text.indexOf(Expression.BEGIN_EXPR, exprEnd + 1);
            }
            exprEnd = text.indexOf(Expression.END_EXPR, exprBegin);

            logger.trace("  tPEAOC: exprBegin = {}, exprEnd = {}", exprBegin, exprEnd);
        }
        return text;
    }

    /**
     * Takes the given <code>PastEndAction</code> on all <code>Cells</code>
     * found inside the given <code>Block</code> on the given <code>Sheet</code>
     * that contain any of the given expressions.
     *
     * @param sheet            The <code>Sheet</code> on which to take a
     *                         <code>PastEndAction</code> on a <code>Block</code>.
     * @param block            The <code>Block</code> of <code>Cells</code>.
     * @param pastEndRefs      A <code>List</code> of strings identifying which
     *                         expressions represent collection access beyond the end of the
     *                         collection.
     * @param pastEndAction    An enumerated value representing the action to take
     *                         on such a cell/expression that references collection access beyond the
     *                         end of the collection.
     * @param replacementValue If the past end action is to replace expressions,
     *                         then this is the value with which to replace those expressions, else
     *                         this is ignored.
     * @see PastEndAction
     */
    public static void takePastEndAction(Sheet sheet, Block block, List<String> pastEndRefs,
                                         PastEndAction pastEndAction, String replacementValue)
    {
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        // No past end refs, no past end actions to take.
        if (pastEndRefs == null || pastEndRefs.size() == 0)
            return;

        logger.trace("takePastEndAction: {}, action {}.", block, pastEndAction);
        logger.trace("  PastEndRefs: {}", pastEndRefs);

        for (int rowNum = top; rowNum <= bottom; rowNum++)
        {
            Row r = sheet.getRow(rowNum);
            if (r != null)
            {
                for (int cellNum = left; cellNum <= right; cellNum++)
                {
                    Cell c = r.getCell(cellNum);
                    if (c != null)
                        takePastEndActionOnCell(c, pastEndRefs, pastEndAction, replacementValue);
                }
            }
        }
    }

    /**
     * Take the given <code>PastEndAction</code> on the given <code>Cell</code>,
     * if its contents contains any of the given references.
     *
     * @param cell             The <code>Cell</code>.
     * @param pastEndRefs      A <code>List</code> of strings identifying which
     *                         expressions represent collection access beyond the end of the
     *                         collection.
     * @param pastEndAction    The <code>PastEndAction</code> to take.
     * @param replacementValue If the past end action is to replace expressions,
     *                         then this is the value with which to replace those expressions, else
     *                         this is ignored.
     */
    private static void takePastEndActionOnCell(Cell cell, List<String> pastEndRefs,
                                                PastEndAction pastEndAction, String replacementValue)
    {
        String strValue;
        boolean takeAction = false;
        if (cell.getCellType() == CellType.STRING)
        {
            for (String pastEndRef : pastEndRefs)
            {
                strValue = cell.getStringCellValue();
                if (strValue != null && strValue.contains(pastEndRef))
                {
                    takeAction = true;
                    break;
                }
            }
        }
        if (takeAction)
        {
            switch (pastEndAction)
            {
            case CLEAR_CELL:
                cell.setCellType(CellType.BLANK);
                break;
            case REMOVE_CELL:
                removeCell(cell.getRow(), cell);
                break;
            case REPLACE_EXPR:
                // Force any expressions containing collection references beyond
                // the end of the collection to "evaluate" to null.  (They're removed.)
                if (cell.getCellType() == CellType.STRING)
                {
                    CreationHelper helper = cell.getSheet().getWorkbook().getCreationHelper();
                    RichTextString rts = cell.getRichStringCellValue();
                    String cellText = rts.getString();
                    int exprBegin = cellText.indexOf(Expression.BEGIN_EXPR);
                    int exprEnd = cellText.indexOf(Expression.END_EXPR);
                    logger.trace("  tPEAOC: exprBegin = {}, exprEnd = {}", exprBegin, exprEnd);

                    while (exprBegin != -1 && exprEnd != -1 && exprEnd > exprBegin)
                    {
                        String expression = cellText.substring(exprBegin + Expression.BEGIN_EXPR.length(), exprEnd - 1 + Expression.END_EXPR.length());
                        boolean replaceExpr = SheetUtil.containsPastEndRef(expression, pastEndRefs);

                        if (replaceExpr)
                        {
                            int afterExprEnd = exprEnd + Expression.END_EXPR.length();
                            String remove = cellText.substring(exprBegin, afterExprEnd);
                            String value = replacementValue;
                            // It doesn't make sense to use the replacement value when the
                            // expression is in the "items" attribute and a Collection is
                            // expected.  Use an empty list instead.
                            // 7 for items=" before the expression
                            // 1 for " after the expression
                            // e.g. items="${remove}"
                            if (exprBegin >= 7 && "items=\"".equals(cellText.substring(exprBegin - 7, exprBegin)) &&
                                    afterExprEnd < cellText.length() - 1 && "\"".equals(cellText.substring(afterExprEnd, afterExprEnd + 1)))
                            {
                                // JEXL for a new, empty ArrayList.
                                value = Expression.BEGIN_EXPR + "new('java.util.ArrayList')" + Expression.END_EXPR;
                            }
                            logger.trace("    removing \"{}\".", remove);
                            rts = RichTextStringUtil.replaceAll(rts, helper, remove, value);
                            cell.setCellValue(rts);
                            cellText = rts.getString();
                            logger.debug("    cellText is now \"{}\".", cellText);
                            exprBegin = cellText.indexOf(Expression.BEGIN_EXPR, exprBegin);
                        }
                        else
                        {
                            exprBegin = cellText.indexOf(Expression.BEGIN_EXPR, exprEnd + 1);
                        }
                        exprEnd = cellText.indexOf(Expression.END_EXPR, exprBegin);

                        logger.trace("  tPEAOC: exprBegin = {}, exprEnd = {}", exprBegin, exprEnd);
                    }
                }
                break;
            default:
                throw new IllegalStateException("Unknown PastEndAction: " + pastEndAction);
            }
        }
    }

    /**
     * Helper method to determine if a "past end reference" is present in the
     * given expression.
     *
     * @param expression  An expression.
     * @param pastEndRefs A <code>List</code> of "past end references".
     * @return Whether a past end reference is present in the expression.
     * @since 0.9.1
     */
    private static boolean containsPastEndRef(String expression, List<String> pastEndRefs)
    {
        Matcher m = POSSIBLE_VARIABLES.matcher(expression);
        while (m.find())
        {
            String possibleVariable = m.group();
            logger.trace("    cPER: Found {}", possibleVariable);
            // Pattern doesn't eliminate possible variables that start
            // with a digit.  Check here.
            if (!Character.isDigit(possibleVariable.charAt(0)) && pastEndRefs.contains(possibleVariable))
            {
                logger.trace("    It's a past end ref!");
                return true;
            }
        }

        // Allow to match "variable.n".
        m = POSSIBLE_VARIABLES2.matcher(expression);
        while (m.find())
        {
            String possibleVariable = m.group();
            logger.trace("    Found {}", possibleVariable);
            // Pattern doesn't eliminate possible variables that start
            // with a digit.  Check here.
            if (!Character.isDigit(possibleVariable.charAt(0)) && pastEndRefs.contains(possibleVariable))
            {
                logger.trace("    It's a past end ref!");
                return true;
            }
        }
        return false;
    }

    /**
     * Removes the given <code>Block</code> of <code>Cells</code> from the given
     * <code>Sheet</code>.
     *
     * @param sheet      The <code>Sheet</code> on which to remove the block.
     * @param tagContext A <code>TagContext</code>.
     * @param block      The <code>Block</code> to remove.
     * @param context    The <code>WorkbookContext</code>.
     */
    public static void removeBlock(Sheet sheet, TagContext tagContext, Block block, WorkbookContext context)
    {
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Block ancestor;
        logger.trace("removeBlock: {}: {}.", sheet.getSheetName(), block);

        int numToShiftUp = bottom - top + 1;
        int numToShiftLeft = right - left + 1;
        int startCellNum, endCellNum, startRowNum, endRowNum;

        switch (block.getDirection())
        {
        case VERTICAL:
            // Cells will be shifted up.
            logger.trace("  Case: Vertical");
            // Shift up all Cells from leftmost to rightmost in Block, from just
            // below the Block all the way down to the bottom of the first Shift
            // Ending Ancestor.
            ancestor = getShiftEndingAncestor(block, -numToShiftUp, 0);
            startRowNum = bottom + 1;
            endRowNum = ancestor.getBottomRowNum();

            // Remove the contents of the Block.
            deleteBlock(sheet, tagContext, block, context);

            // If we reached the root parent, and our block is as wide as it, then
            // shrink it too.
            if (ancestor.getParent() == null &&
                    left == ancestor.getLeftColNum() && right == ancestor.getRightColNum())
            {
                logger.debug("  Shrinking ancestor block ({}) by {} rows!",
                        ancestor, numToShiftUp);
                ancestor.expand(0, -numToShiftUp);
                copyRowHeightsUp(sheet, startRowNum, endRowNum, numToShiftUp);
            }
            shiftCellsUp(sheet, tagContext, context, left, right, startRowNum, endRowNum, numToShiftUp);
            FormulaUtil.shiftCellReferencesInRange(sheet.getSheetName(), context,
                    left, right, startRowNum, endRowNum,
                    0, -numToShiftUp, true, true);
            break;
        case HORIZONTAL:
            // Cells will be shifted left.
            logger.trace("  Case: Horizontal");
            // Shift left all Cells from the top to the bottom in Block, from just
            // to the right of the Block all the way to the far right of the first
            // Shift Ending Ancestor.
            ancestor = getShiftEndingAncestor(block, 0, -numToShiftLeft);
            startCellNum = right + 1;
            endCellNum = ancestor.getRightColNum();

            // Remove the contents of the Block.
            deleteBlock(sheet, tagContext, block, context);

            // If we reached the root parent, and our block is as tall as it, then
            // shrink it too.
            if (ancestor.getParent() == null &&
                    top == ancestor.getTopRowNum() && bottom == ancestor.getBottomRowNum())
            {
                logger.debug("  Shrinking ancestor block ({}) by {} columns!",
                        ancestor, numToShiftLeft);
                ancestor.expand(-numToShiftLeft, 0);
                copyColumnWidthsLeft(sheet, startCellNum, endCellNum, numToShiftLeft);
            }

            shiftCellsLeft(sheet, tagContext, context, startCellNum, endCellNum, top, bottom, numToShiftLeft);
            FormulaUtil.shiftCellReferencesInRange(sheet.getSheetName(), context,
                    startCellNum, endCellNum, top, bottom,
                    -numToShiftLeft, 0, true, true);
            break;
        case NONE:
            logger.trace("  Case: None");
            // Remove the Block content.
            deleteBlock(sheet, tagContext, block, context);
            break;
        }
    }

    /**
     * Walk up the <code>Block</code> tree until a "shift ending" ancestor is
     * found, or until the tree has been exhausted.  Optionally, grow/shrink
     * parent blocks encountered until the "shift ending" ancestor is found.
     * (The "shift ending" ancestor is not grown/shrunk).  The "shift ending"
     * ancestor is defined as an ancestor <code>Block</code> that is either a
     * different direction than the original <code>Block</code> or is larger
     * than the original <code>Block</code> along the other direction (that is,
     * larger in height for Horizontal blocks, or larger in width for Vertical
     * blocks).
     *
     * @param block         The <code>Block</code> to search for ancestors.
     * @param numVertCells  The number of cells to grow each parent vertically
     *                      until the "shift ending" ancestor is found, or shrink if
     *                      <code>numCells</code> is negative.
     * @param numHorizCells The number of cells to grow each parent horizontally
     *                      until the "shift ending" ancestor is found, or shrink if
     *                      <code>numCells</code> is negative.
     * @return The closest "shift ending" ancestor <code>Block</code>.
     */
    public static Block getShiftEndingAncestor(Block block, int numVertCells, int numHorizCells)
    {
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Block ancestor = block.getParent();
        Block.Direction dir = block.getDirection();

        switch (dir)
        {
        case VERTICAL:
            while (ancestor != null)
            {
                if (ancestor.getDirection() != dir || left != ancestor.getLeftColNum() ||
                        right != ancestor.getRightColNum())
                    break;

                // Ancestors grow until the Shift Ending Ancestor is found.
                if (numVertCells != 0)
                {
                    logger.debug("    Growing ancestor block ({}) by {} rows!", ancestor, numVertCells);
                    ancestor.expand(0, numVertCells);
                }
                if (numHorizCells != 0)
                {
                    logger.debug("    Growing ancestor block ({}) by {} columns!", ancestor, numHorizCells);
                    ancestor.expand(numHorizCells, 0);
                }

                // Prepare for next loop.
                ancestor = ancestor.getParent();
            }
            break;
        case HORIZONTAL:
            while (ancestor != null)
            {
                if (ancestor.getDirection() != dir || top != ancestor.getTopRowNum() ||
                        bottom != ancestor.getBottomRowNum())
                    break;

                // Ancestors grow until the Shift Ending Ancestor is found.
                if (numVertCells != 0)
                {
                    logger.debug("    Growing ancestor block ({}) by {} rows!", ancestor, numVertCells);
                    ancestor.expand(0, numVertCells);
                }
                if (numHorizCells != 0)
                {
                    logger.debug("    Growing ancestor block ({}) by {} columns!", ancestor, numHorizCells);
                    ancestor.expand(numHorizCells, 0);
                }

                // Prepare for next loop.
                ancestor = ancestor.getParent();
            }
            break;
        }
        logger.trace("    gSEA: Ancestor of {} is {}", block, ancestor);
        return ancestor;
    }

    /**
     * Shifts <code>Cells</code> out of the way.
     *
     * @param sheet         The <code>Sheet</code> on which to shift.
     * @param tagContext    A <code>TagContext</code>.
     * @param block         The <code>Block</code> whose copies will occupy the
     *                      <code>Cells</code> that will move to make way for the copies.
     * @param context       The <code>WorkbookContext</code>.
     * @param numBlocksAway The number of blocks (widths or lengths, depending
     *                      on the case of <code>block</code> that defines the area of
     *                      <code>Cells</code> to shift.
     */
    public static void shiftForBlock(Sheet sheet, TagContext tagContext, Block block, WorkbookContext context, int numBlocksAway)
    {
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Block ancestor, prevAncestor;
        Block parent = block.getParent();

        logger.trace("shiftForBlock: {}: {}, numBlocksAway={}.", sheet.getSheetName(), block, numBlocksAway);
        logger.trace("  parent: " + parent);

        // Below this point!

        // If moving down...
        int height = bottom - top + 1;
        int translateDown = (numBlocksAway - 1) * height;  // Make room for n - 1 more Blocks.
        // If moving right...
        int width = right - left + 1;
        int translateRight = (numBlocksAway - 1) * width;  // Make room for n - 1 more Blocks.

        int startCellNum, endCellNum, startRowNum, endRowNum;
        Stack<Block> blocksToShift = new Stack<>();
        Stack<Integer> shiftAmounts = new Stack<>();

        switch (block.getDirection())
        {
        case VERTICAL:
            // Cells will be shifted down.
            logger.trace("  Case Vertical");
            // The number of shift operations could be as many as the number of
            // Shift Ending Ancestors + 1 (for the root parent of the Sheet).
            // Keep finding Shift Ending Ancestors (or the root) and push a new
            // shift operation for each one.
            prevAncestor = block;
            ancestor = getShiftEndingAncestor(block, translateDown, 0);
            // Gather temporary Blocks to shift until a Shift Ending Ancestor has
            // enough room already, or we've reached the root parent Block.
            while (translateDown > 0)
            {
                // Define the Block of Cells that will get shifted.
                startRowNum = prevAncestor.getBottomRowNum() + 1;
                startCellNum = prevAncestor.getLeftColNum();
                endCellNum = prevAncestor.getRightColNum();
                endRowNum = ancestor.getBottomRowNum();
                if (prevAncestor.getDirection() == Block.Direction.HORIZONTAL)
                {
                    // Below a Horizontal Ancestor, the range of columns in the
                    // block to shift downwards is bigger.  (This content has not
                    // been transformed yet.)
                    startCellNum = ancestor.getLeftColNum();
                    endCellNum = ancestor.getRightColNum();
                }

                // If the previous ancestor was already expanded, then the top edge
                // of this block hasn't been shifted yet.
                if (!shiftAmounts.isEmpty())
                    startRowNum -= shiftAmounts.peek();

                // Empty rows at the bottom mean less rows to shift and future
                // shifts will be smaller.  Only do this in the first loop.
                int emptyRowsAtBottom = getEmptyRowsAtBottom(sheet, startCellNum, endCellNum, startRowNum, endRowNum);
                if (emptyRowsAtBottom > 0)
                    endRowNum -= emptyRowsAtBottom;
                logger.debug("    emptyRowsAtBottom: {}", emptyRowsAtBottom);

                Block toShift = new Block(null, startCellNum, endCellNum, startRowNum, endRowNum);
                logger.debug("    Block to shift: {} by {} rows.", toShift, translateDown);
                blocksToShift.push(toShift);
                shiftAmounts.push(translateDown);
                // The shifting will fill the bottom of the block.  Reduce the
                // ancestor's expansion amount.
                if (emptyRowsAtBottom > 0)
                    translateDown -= emptyRowsAtBottom;
                // Manually expand the Shift Ending Ancestor.
                if (translateDown > 0)
                {
                    logger.debug("    Growing ancestor block ({}) by {} rows!", ancestor, translateDown);
                    ancestor.expand(0, translateDown);
                }

                // Prepare for next loop.
                prevAncestor = ancestor;
                if (ancestor.getParent() != null)
                    ancestor = getShiftEndingAncestor(ancestor, translateDown, 0);
                else  // Already reached root.
                    break;
            }

            // Perform the shifts in reverse order of found (LIFO).
            while (!blocksToShift.isEmpty())
            {
                Block toShift = blocksToShift.pop();
                translateDown = shiftAmounts.pop();

                // Don't copy the same row heights down again.  This would occur if
                // the parent is horizontal and already past its first iteration, e.g.
                //   |     |
                //   |     |
                //   v     v
                if (parent == null || parent.getDirection() != Block.Direction.HORIZONTAL || parent.getIterationNbr() == 0)
                {
                    copyRowHeightsDown(sheet, toShift.getTopRowNum(), toShift.getBottomRowNum(), translateDown);
                }
                shiftCellsDown(sheet, tagContext, context, toShift.getLeftColNum(), toShift.getRightColNum(),
                        toShift.getTopRowNum(), toShift.getBottomRowNum(), translateDown);
                FormulaUtil.shiftCellReferencesInRange(sheet.getSheetName(), context,
                        toShift.getLeftColNum(), toShift.getRightColNum(), toShift.getTopRowNum(), toShift.getBottomRowNum(),
                        0, translateDown, true, true);
            }
            break;
        case HORIZONTAL:
            // Cells will be shifted right.
            logger.trace("  Case Horizontal");
            // The number of shift operations could be as many as the number of
            // Shift Ending Ancestors + 1 (for the root parent of the Sheet).
            // Keep finding Shift Ending Ancestors (or the root) and push a new
            // shift operation for each one.
            prevAncestor = block;
            ancestor = getShiftEndingAncestor(block, 0, translateRight);
            // Gather temporary Blocks to shift until a Shift Ending Ancestor has
            // enough room already, or we've reached the root parent Block.
            while (translateRight > 0)
            {
                // Define the Block of Cells that will get shifted.
                startCellNum = prevAncestor.getRightColNum() + 1;
                startRowNum = prevAncestor.getTopRowNum();
                endRowNum = prevAncestor.getBottomRowNum();
                endCellNum = ancestor.getRightColNum();
                // To the right of a Vertical Ancestor, do not expand the row
                // range.  Content above and to the right has already been
                // transformed.  Content below and to the right will be on its own.

                // If the previous ancestor was already expanded, then the top edge
                // of this block hasn't been shifted yet.
                if (!shiftAmounts.isEmpty())
                    startCellNum -= shiftAmounts.peek();

                // Empty cols at the right mean less cols to shift and future
                // shifts will be smaller.   Only do this in the first loop.
                int emptyColsAtRight = getEmptyColumnsAtRight(sheet, startCellNum, endCellNum, startRowNum, endRowNum);
                if (emptyColsAtRight > 0)
                    endCellNum -= emptyColsAtRight;
                logger.debug("    emptyColsAtRight: {}", emptyColsAtRight);

                Block toShift = new Block(null, startCellNum, endCellNum, startRowNum, endRowNum);
                logger.debug("    Block to shift: {} by {} columns.", toShift, translateRight);
                blocksToShift.push(toShift);
                shiftAmounts.push(translateRight);
                // The shifting will fill the far right of the block.  Reduce
                // the ancestor's expansion amount.
                if (emptyColsAtRight > 0)
                    translateRight -= emptyColsAtRight;
                if (translateRight > 0)
                {
                    // Manually expand the Block Area ancestor (or the root!).
                    logger.debug("    Growing ancestor block ({}) by {} columns!", ancestor, translateRight);
                    ancestor.expand(translateRight, 0);
                }

                // Prepare for next loop.
                prevAncestor = ancestor;
                if (ancestor.getParent() != null)
                    ancestor = getShiftEndingAncestor(ancestor, 0, translateRight);
                else  // Already reached root.
                    break;
            }

            // Perform the shifts in reverse order of found (LIFO).
            while (!blocksToShift.isEmpty())
            {
                Block toShift = blocksToShift.pop();
                translateRight = shiftAmounts.pop();

                // Don't copy the same column widths right again.  This would occur if
                // the parent is vertical and already past its first iteration, e.g.
                //   ----->
                //   ----->
                if (parent == null || parent.getDirection() != Block.Direction.VERTICAL || parent.getIterationNbr() == 0)
                {
                    copyColumnWidthsRight(sheet, toShift.getLeftColNum(), toShift.getRightColNum(), translateRight);
                }
                shiftCellsRight(sheet, tagContext, context, toShift.getLeftColNum(), toShift.getRightColNum(),
                        toShift.getTopRowNum(), toShift.getBottomRowNum(), translateRight);
                FormulaUtil.shiftCellReferencesInRange(sheet.getSheetName(), context,
                        toShift.getLeftColNum(), toShift.getRightColNum(), toShift.getTopRowNum(), toShift.getBottomRowNum(),
                        translateRight, 0, true, true);
            }
            break;
        }
    }

    /**
     * Determine how many "empty" rows are at the bottom of the given
     * block of cells, between the left and right positions (inclusive).
     *
     * @param sheet  The <code>Sheet</code> on which the <code>Block</code> is
     *               located.
     * @param left   The 0-based column position to start looking for empty cells.
     * @param right  The 0-based column position to stop looking for empty cells.
     * @param top    The 0-based row index to stop looking for empty cells.
     * @param bottom The 0-based row index to start looking for empty cells.
     * @return The number of empty rows at the bottom of the <code>Block</code>.
     */
    private static int getEmptyRowsAtBottom(Sheet sheet, int left, int right, int top, int bottom)
    {
        int emptyRows = 0;
        for (int r = bottom; r >= top; r--)
        {
            boolean rowEmpty = true;
            for (int c = left; c <= right; c++)
            {
                if (!isCellImmaterial(sheet, r, c))
                {
                    logger.trace("      gERAB: Row {} is not empty because of cell {}", r, c);
                    rowEmpty = false;
                    break;
                }
            }
            if (rowEmpty)
                emptyRows++;
            else
                break;
        }
        return emptyRows;
    }

    /**
     * Determine how many "empty" columns are at the right of the given
     * block of cells, between the top and bottom positions (inclusive).
     *
     * @param sheet  The <code>Sheet</code> on which the <code>Block</code> is
     *               located.
     * @param left   The 0-based column position to stop looking for empty cells.
     * @param right  The 0-based column position to start looking for empty cells.
     * @param top    The 0-based row index to start looking for empty cells.
     * @param bottom The 0-based row index to stop looking for empty cells.
     * @return The number of empty columns at the right of the <code>Block</code>.
     */
    private static int getEmptyColumnsAtRight(Sheet sheet, int left, int right, int top, int bottom)
    {
        int emptyColumns = 0;
        for (int c = right; c >= left; c--)
        {
            boolean colEmpty = true;
            for (int r = top; r <= bottom; r++)
            {
                Row row = sheet.getRow(r);
                if (row != null)
                {
                    if (!isCellImmaterial(sheet, r, c))
                    {
                        logger.trace("      gECAR: Column {} is not empty because of row {}", c, r);
                        colEmpty = false;
                        break;
                    }
                }
            }
            if (colEmpty)
                emptyColumns++;
            else
                break;
        }
        return emptyColumns;
    }

    /**
     * Copies an entire <code>Block</code> the given number of blocks away on
     * the given <code>Sheet</code>.
     *
     * @param sheet         The <code>Sheet</code> on which to copy.
     * @param tagContext    A <code>TagContext</code>.
     * @param block         The <code>Block</code> to copy.
     * @param context       The <code>WorkbookContext</code>.
     * @param numBlocksAway The number of blocks (widths or lengths, depending
     *                      on the direction of <code>block</code>), away to copy.
     * @return The newly copied <code>Block</code>.
     */
    public static Block copyBlock(Sheet sheet, TagContext tagContext, Block block, WorkbookContext context, int numBlocksAway)
    {
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        Block parent = block.getParent();
        Block newBlock = null;
        String sheetName = sheet.getSheetName();
        int seqNbr = context.getSequenceNbr();
        String currSuffix = tagContext.getFormulaSuffix();
        String newSuffix = "[" + seqNbr + "," + numBlocksAway + "]";
        Map<String, String> tagLocationsMap = context.getTagLocationsMap();
        logger.trace("copyBlock: {}: {}, numBlocksAway={}", sheet.getSheetName(), block, numBlocksAway);

        // If copying down...
        int height = block.getBottomRowNum() - block.getTopRowNum() + 1;
        int translateDown = numBlocksAway * height;
        int newTop = top + translateDown;
        int newBottom = bottom + translateDown;
        // If copying right...
        int width = block.getRightColNum() - block.getLeftColNum() + 1;
        int translateRight = numBlocksAway * width;
        int newLeft = left + translateRight;
        int newRight = right + translateRight;

        switch (block.getDirection())
        {
        case VERTICAL:
            logger.trace("  Case Vertical");
            // Copy Cells.
            logger.debug("    Copying cells {} to {}, row {} to {} by {} rows.",
                    left, right, top, bottom, translateDown);
            for (int r = top; r <= bottom; r++)
            {
                Row oldRow = sheet.getRow(r);
                if (oldRow == null)
                    oldRow = sheet.createRow(r);
                Row newRow = sheet.getRow(r + translateDown);
                if (newRow == null)
                    newRow = sheet.createRow(r + translateDown);
                for (int c = left; c <= right; c++)
                {
                    Cell oldCell = oldRow.getCell(c);
                    if (oldCell == null)
                        oldCell = oldRow.createCell(c);
                    Cell newCell = newRow.getCell(c);
                    if (newCell == null)
                        newCell = newRow.createCell(c);
                    if (numBlocksAway > 0)
                        copyCell(oldCell, newCell);

                    String oldCellRef = getCellKey(oldCell);
                    String newCellRef = getCellKey(newCell);
                    String origCellRef = tagLocationsMap.get(oldCellRef);
                    if (origCellRef != null)
                    {
                        tagLocationsMap.put(newCellRef, origCellRef);
                        logger.debug("cB: Adding {} => {}", newCellRef, origCellRef);
                    }

                    // Append "[loop,iter]" on formulas.
                    if (newCell.getCellType() == CellType.STRING)
                    {
                        String cellText = newCell.getStringCellValue();
                        int startIdx = cellText.indexOf(Formula.BEGIN_FORMULA);
                        int endIdx = cellText.lastIndexOf(Formula.END_FORMULA);
                        if (startIdx != -1 && endIdx != -1 && startIdx < endIdx)
                        {
                            // If this is NOT the first iteration, then the copied
                            // text already has the previous iteration's suffix
                            // appended to it!  Remove it first.
                            if (numBlocksAway > 0)
                            {
                                int idx = cellText.lastIndexOf("[");
                                if (idx > -1)
                                    cellText = cellText.substring(0, idx);  // Lose the last suffix.
                            }
                            String newFormula = cellText + newSuffix;
                            setCellValue(context, newCell, newFormula);
                        }
                    }
                }
            }

            if (numBlocksAway > 0)
            {
                // Copy merged regions down.
                shiftMergedRegionsInRange(tagContext, left, right,
                        top, bottom, 0, translateDown, false, true);
                // Copy conditional formatting regions down.
                //copyConditionalFormattingRegionsInRange(sheet, left, right,
                //   top, bottom, 0, translateDown);
                // Don't copy the same row heights down again.  This would occur if
                // the parent is horizontal and already past its first iteration, e.g.
                //   |     |
                //   |     |
                //   v     v
                if (parent == null || parent.getDirection() != Block.Direction.HORIZONTAL || parent.getIterationNbr() == 0)
                {
                    copyRowHeightsDown(sheet, top, bottom, translateDown);
                }
                newBlock = new Block(parent, left, right, newTop, newBottom, numBlocksAway);
                newBlock.setDirection(block.getDirection());
            }
            else
                newBlock = block;

            FormulaUtil.copyCellReferencesInRange(sheetName, context,
                    left, right, top, bottom, 0, translateDown, currSuffix, newSuffix);
            break;
        case HORIZONTAL:
            logger.trace("  Case Horizontal");

            // Copy Cells.
            logger.debug("    Copying cells {} to {}, row {} to {} by {} columns.",
                    left, right, top, bottom, translateRight);
            for (int r = top; r <= bottom; r++)
            {
                Row row = sheet.getRow(r);
                if (row == null)
                    row = sheet.createRow(r);
                for (int col = left; col <= right; col++)
                {
                    Cell oldCell = row.getCell(col);
                    if (oldCell == null)
                        oldCell = row.createCell(col);
                    Cell newCell = row.getCell(col + translateRight);
                    if (newCell == null)
                        newCell = row.createCell(col + translateRight);
                    if (numBlocksAway > 0)
                        copyCell(oldCell, newCell);

                    String oldCellRef = getCellKey(oldCell);
                    String newCellRef = getCellKey(newCell);
                    String origCellRef = tagLocationsMap.get(oldCellRef);
                    if (origCellRef != null)
                    {
                        tagLocationsMap.put(newCellRef, origCellRef);
                        logger.debug("cB: Adding {} => {}", newCellRef, origCellRef);
                    }

                    // Append proper "[loop,iter]" on formulas.
                    if (newCell.getCellType() == CellType.STRING)
                    {
                        String cellText = newCell.getStringCellValue();
                        int startIdx = cellText.indexOf(Formula.BEGIN_FORMULA);
                        int endIdx = cellText.lastIndexOf(Formula.END_FORMULA);
                        if (startIdx != -1 && endIdx != -1 && startIdx < endIdx)
                        {
                            // If this is NOT the first iteration, then the copied
                            // text already has the previous iteration's suffix
                            // appended to it!  Remove it first.
                            if (numBlocksAway > 0)
                            {
                                int idx = cellText.lastIndexOf("[");
                                if (idx > -1)
                                    cellText = cellText.substring(0, idx);  // Lose the last suffix.
                            }
                            String newFormula = cellText + newSuffix;
                            setCellValue(context, newCell, newFormula);
                        }
                    }
                }
            }

            if (numBlocksAway > 0)
            {
                // Copy merged regions right.
                shiftMergedRegionsInRange(tagContext, left, right, top, bottom, translateRight, 0, false, true);
                // Copy conditional formatting regions down.
                //copyConditionalFormattingRegionsInRange(sheet, left, right,
                //   top, bottom, translateRight, 0);
                // Don't copy the same column widths right again.  This would occur if
                // the parent is vertical and already past its first iteration, e.g.
                //   ----->
                //   ----->
                if (parent == null || parent.getDirection() != Block.Direction.VERTICAL || parent.getIterationNbr() == 0)
                {
                    copyColumnWidthsRight(sheet, left, right, translateRight);
                }
                newBlock = new Block(parent, newLeft, newRight, top, bottom, numBlocksAway);
                newBlock.setDirection(block.getDirection());
            }
            else
                newBlock = block;

            FormulaUtil.copyCellReferencesInRange(sheetName, context,
                    left, right, top, bottom, translateRight, 0, currSuffix, newSuffix);
            break;
        }
        return newBlock;
    }

    /**
     * Replace all occurrences of the given collection expression name with the
     * given item name, in preparation for implicit collections processing
     * loops.
     *
     * @param sheet     The <code>Sheet</code> on which the <code>Block</code> lies.
     * @param block     The <code>Block</code> in which to perform the replacement.
     * @param collExprs The collection expression strings to replace.
     * @param itemNames The item names that replace the collection expressions.
     */
    public static void setUpBlockForImplicitCollectionAccess(Sheet sheet, Block block,
                                                             List<String> collExprs, List<String> itemNames)
    {
        int left = block.getLeftColNum();
        int right = block.getRightColNum();
        int top = block.getTopRowNum();
        int bottom = block.getBottomRowNum();
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        // Look at the given range of Cells in the given range of rows.
        for (int rowNum = top; rowNum <= bottom; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            if (row != null)
            {
                for (int cellNum = left; cellNum <= right; cellNum++)
                {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null && cell.getCellType() == CellType.STRING)
                    {
                        RichTextString value = cell.getRichStringCellValue();
                        for (int i = 0; i < collExprs.size(); i++)
                        {
                            value = RichTextStringUtil.replaceAll(
                                    value, helper, collExprs.get(i), itemNames.get(i), false, 0, true);
                        }
                        cell.setCellValue(value);
                    }
                }
            }
        }
    }

    /**
     * Replace all occurrences of the given collection expression name with the
     * given item name, in the entire <code>Sheet</code>, in preparation for
     * implicit cloning processing loops.  This doesn't replace any collection
     * expressions in the sheet name; this is assumed to have taken place
     * already, due to unique sheet naming requirements.
     *
     * @param sheet     The <code>Sheet</code>.
     * @param collExprs The <code>List</code> of collection expression strings
     *                  to replace.
     * @param itemNames The <code>List</code> of item names that replace the
     *                  collection expressions.
     * @since 0.9.1
     */
    public static void setUpSheetForImplicitCloningAccess(Sheet sheet, List<String> collExprs, List<String> itemNames)
    {
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();

        Header header = sheet.getHeader();
        header.setLeft(replacementHelper(helper, header.getLeft(), collExprs, itemNames));
        header.setCenter(replacementHelper(helper, header.getCenter(), collExprs, itemNames));
        header.setRight(replacementHelper(helper, header.getRight(), collExprs, itemNames));

        Footer footer = sheet.getFooter();
        footer.setLeft(replacementHelper(helper, footer.getLeft(), collExprs, itemNames));
        footer.setCenter(replacementHelper(helper, footer.getCenter(), collExprs, itemNames));
        footer.setRight(replacementHelper(helper, footer.getRight(), collExprs, itemNames));

        for (Row row : sheet)
        {
            for (Cell cell : row)
            {
                if (cell.getCellType() == CellType.STRING)
                {
                    RichTextString value = cell.getRichStringCellValue();
                    value = replacementHelper(helper, value, collExprs, itemNames);
                    cell.setCellValue(value);
                }
            }
        }
    }

    /**
     * Helper method to replace all collection expressions with their
     * replacements for one value.
     *
     * @param helper    The <code>CreationHelper</code>.
     * @param text      The rich text string containing expressions to replace.
     * @param collExprs The <code>List</code> of collection expression strings
     *                  to replace.
     * @param itemNames The <code>List</code> of item names that replace the
     *                  collection expressions.
     * @return The rich text string with expressions replaced.
     * @since 0.9.1
     */
    private static RichTextString replacementHelper(CreationHelper helper, RichTextString text, List<String> collExprs, List<String> itemNames)
    {
        for (int i = 0; i < collExprs.size(); i++)
        {
            text = RichTextStringUtil.replaceAll(text, helper, collExprs.get(i), itemNames.get(i), false, 0, true);
        }
        return text;
    }

    /**
     * Helper method to replace all collection expressions with their
     * replacements for one value.
     *
     * @param helper    The <code>CreationHelper</code>.
     * @param text      The text containing expressions to replace.
     * @param collExprs The <code>List</code> of collection expression strings
     *                  to replace.
     * @param itemNames The <code>List</code> of item names that replace the
     *                  collection expressions.
     * @return The text with expressions replaced.
     * @since 0.9.1
     */
    private static String replacementHelper(CreationHelper helper, String text, List<String> collExprs, List<String> itemNames)
    {
        RichTextString temp = helper.createRichTextString(text);
        temp = replacementHelper(helper, temp, collExprs, itemNames);
        return temp.toString();
    }

    /**
     * Group all rows on the sheet between the "begin" and "end" indices,
     * inclusive.  Optionally collapse the rows.
     *
     * @param sheet    The <code>Sheet</code> on which to group the rows.
     * @param begin    The 0-based index of the start row of the group.
     * @param end      The 0-based index of the end row of the group.
     * @param collapse Whether to collapse the group.
     * @since 0.2.0
     */
    public static void groupRows(Sheet sheet, int begin, int end, boolean collapse)
    {
        logger.trace("groupRows: {}, ({} - {}), collapse: {}",
                sheet.getSheetName(), begin, end, collapse);
        sheet.groupRow(begin, end);
        if (collapse)
        {
            if (sheet instanceof XSSFSheet)
            {
                // XSSFSheet - Must manually collapse the rows.
                for (int r = begin; r <= end; r++)
                {
                    Row row = sheet.getRow(r);
                    if (row == null)
                        row = sheet.createRow(r);
                    row.setZeroHeight(true);
                }
            }
            else
            {
                // HSSFSheet - setRowGroupCollapsed works.
                sheet.setRowGroupCollapsed(begin, true);
            }
        }
    }

    /**
     * Group all columns on the sheet between the "begin" and "end" indices,
     * inclusive.  Optionally collapse the columns.
     *
     * @param sheet    The <code>Sheet</code> on which to group the columns.
     * @param begin    The 0-based index of the start column of the group.
     * @param end      The 0-based index of the end column of the group.
     * @param collapse Whether to collapse the group.
     * @since 0.2.0
     */
    public static void groupColumns(Sheet sheet, int begin, int end, boolean collapse)
    {
        logger.trace("groupColumns: {}, ({} - {}), collapse: {}",
                sheet.getSheetName(), begin, end, collapse);
        // XSSFSheets will collapse the columns on "groupColumn".
        // Store the column widths to restore them after "groupColumn".
        Map<Integer, Integer> colWidths = new HashMap<>();
        if (sheet instanceof XSSFSheet)
        {
            logger.debug("Def. col width = {}", sheet.getDefaultColumnWidth());
            for (int c = begin; c <= end; c++)
            {
                int w = sheet.getColumnWidth(c);
                logger.debug("Col {}, {}", c, w);
                colWidths.put(c, w);
            }
        }
        if (sheet instanceof XSSFSheet)
        {
            // When nested, XSSFSheet's groupColumn doesn't do the whole range.
            for (int c = begin; c <= end; c++)
            {
                sheet.groupColumn(c, c);
                int w = colWidths.get(c);
                logger.debug("Setting Col {}, to width {}", c, w);
                sheet.setColumnWidth(c, w);
            }
        }
        else
        {
            // HSSFSheet works as expected.
            sheet.groupColumn(begin, end);
        }
        if (collapse)
        {
            if (sheet instanceof XSSFSheet)
            {
                // XSSFSheet - Must manually collapse the columns.
                for (int c = begin; c <= end; c++)
                {
                    logger.debug("Setting Col {} hidden", c);
                    sheet.setColumnHidden(c, true);
                }
            }
            else
            {
                // HSSFSheet - setColumnGroupCollapsed works.
                sheet.setColumnGroupCollapsed(begin, true);
            }
        }
    }

    /**
     * Get the hex string that represents the <code>Color</code>.
     *
     * @param color A POI <code>Color</code>.
     * @return The hex string that represents the <code>Color</code>.
     * @since 0.5.0
     */
    public static String getColorHexString(Color color)
    {
        if (color instanceof HSSFColor)
        {
            HSSFColor hssfColor = (HSSFColor) color;
            return getHSSFColorHexString(hssfColor);
        }
        else if (color instanceof XSSFColor)
        {
            XSSFColor xssfColor = (XSSFColor) color;
            return getXSSFColorHexString(xssfColor);
        }
        else if (color == null)
        {
            return "000000";
        }
        else
        {
            throw new IllegalArgumentException("Unexpected type of Color: " + color.getClass().getName());
        }
    }

    /**
     * Get the hex string for a <code>HSSFColor</code>.  Moved from test code.
     *
     * @param hssfColor A <code>HSSFColor</code>.
     * @return The hex string.
     * @since 0.5.0
     */
    private static String getHSSFColorHexString(HSSFColor hssfColor)
    {
        short[] shorts = hssfColor.getTriplet();
        StringBuilder hexString = new StringBuilder();
        for (short s : shorts)
        {
            String twoHex = Integer.toHexString(0x000000FF & s);
            if (twoHex.length() == 1)
                hexString.append('0');
            hexString.append(twoHex);
        }
        return hexString.toString();
    }

    /**
     * Get the hex string for a <code>XSSFColor</code>.  Moved from test code.
     *
     * @param xssfColor A <code>XSSFColor</code>.
     * @return The hex string.
     * @since 0.5.0
     */
    private static String getXSSFColorHexString(XSSFColor xssfColor)
    {
        if (xssfColor == null)
            return "000000";
        byte[] bytes;
        // As of Apache POI 3.8, there are Bugs 51236 and 52079 about font
        // color where somehow black and white get switched.  It appears to
        // have to do with the fact that black and white "theme" colors get
        // flipped.  Be careful, because XSSFColor(byte[]) does NOT call
        // "correctRGB", but XSSFColor.setRgb(byte[]) DOES call it, and so
        // does XSSFColor.getRgb(byte[]).
        // The private method "correctRGB" flips black and white, but no
        // other colors.  However, correctRGB is its own inverse operation,
        // i.e. correctRGB(correctRGB(rgb)) yields the same bytes as rgb.
        // XSSFFont.setColor(XSSFColor) calls "getRGB", but
        // XSSFCellStyle.set[Xx]BorderColor and
        // XSSFCellStyle.setFill[Xx]Color do NOT.
        // Solution: Correct the font color on the way out for themed colors
        // only.  For unthemed colors, bypass the "correction".
        if (xssfColor.getCTColor().isSetTheme())
            bytes = xssfColor.getRGB();
        else
            bytes = xssfColor.getCTColor().getRgb();
        // End of workaround for Bugs 51236 and 52079.
        if (bytes == null)
        {
            // Indexed Color - like HSSF
            HSSFColor hColor = ExcelColor.getHssfColorByIndex(xssfColor.getIndexed());
            if (hColor != null)
                return getHSSFColorHexString(ExcelColor.getHssfColorByIndex(xssfColor.getIndexed()));
            else
                return "000000";
        }
        if (bytes.length == 4)
        {
            // Lose the alpha.
            bytes = new byte[] {bytes[1], bytes[2], bytes[3]};
        }
        StringBuilder hexString = new StringBuilder();
        for (byte b : bytes)
        {
            String twoHex = Integer.toHexString(0x000000FF & b);
            if (twoHex.length() == 1)
                hexString.append('0');
            hexString.append(twoHex);
        }
        return hexString.toString();
    }

    /**
     * Creates a new <code>CellStyle</code> for the given <code>Workbook</code>,
     * with the given attributes.  Moved from <code>StyleTag</code> here for
     * 0.5.0.
     *
     * @param workbook            A <code>Workbook</code>.
     * @param alignment           A <code>short</code> alignment constant.
     * @param borderBottom        A <code>short</code> border type constant.
     * @param borderLeft          A <code>short</code> border type constant.
     * @param borderRight         A <code>short</code> border type constant.
     * @param borderTop           A <code>short</code> border type constant.
     * @param dataFormat          A data format string.
     * @param wrapText            Whether text is wrapped.
     * @param fillBackgroundColor A background <code>Color</code>.
     * @param fillForegroundColor A foreground <code>Color</code>.
     * @param fillPattern         A <code>short</code> pattern constant.
     * @param verticalAlignment   A <code>short</code> vertical alignment constant.
     * @param indention           A <code>short</code> number of indent characters.
     * @param rotationDegrees     A <code>short</code> degrees rotation of text.
     * @param bottomBorderColor   A border <code>Color</code> object.
     * @param leftBorderColor     A border <code>Color</code> object.
     * @param rightBorderColor    A border <code>Color</code> object.
     * @param topBorderColor      A border <code>Color</code> object.
     * @param locked              Whether the cell is locked.
     * @param hidden              Whether the cell is hidden.
     * @return A new <code>CellStyle</code>.
     */
    public static CellStyle createCellStyle(Workbook workbook, HorizontalAlignment alignment, BorderStyle borderBottom, BorderStyle borderLeft,
    										BorderStyle borderRight, BorderStyle borderTop, String dataFormat, boolean wrapText, Color fillBackgroundColor,
                                            Color fillForegroundColor, FillPatternType fillPattern, VerticalAlignment verticalAlignment, short indention,
                                            short rotationDegrees, Color bottomBorderColor, Color leftBorderColor,
                                            Color rightBorderColor, Color topBorderColor, boolean locked, boolean hidden)
    {
        CellStyle cs = workbook.createCellStyle();
        cs.setAlignment(alignment);
        cs.setBorderBottom(borderBottom);
        cs.setBorderLeft(borderLeft);
        cs.setBorderRight(borderRight);
        cs.setBorderTop(borderTop);
        cs.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(dataFormat));
        cs.setHidden(hidden);
        cs.setIndention(indention);
        cs.setLocked(locked);
        cs.setRotation(rotationDegrees);
        cs.setVerticalAlignment(verticalAlignment);
        cs.setWrapText(wrapText);
        // Certain properties need a type of workbook check.
        if (workbook instanceof HSSFWorkbook)
        {
            if (bottomBorderColor != null)
                cs.setBottomBorderColor(((HSSFColor) bottomBorderColor).getIndex());
            if (leftBorderColor != null)
                cs.setLeftBorderColor(((HSSFColor) leftBorderColor).getIndex());
            if (rightBorderColor != null)
                cs.setRightBorderColor(((HSSFColor) rightBorderColor).getIndex());
            if (topBorderColor != null)
                cs.setTopBorderColor(((HSSFColor) topBorderColor).getIndex());
            // Per POI Javadocs, set foreground color first!
            cs.setFillForegroundColor(((HSSFColor) fillForegroundColor).getIndex());
            cs.setFillBackgroundColor(((HSSFColor) fillBackgroundColor).getIndex());
        }
        else
        {
            // XSSFWorkbook
            XSSFCellStyle xcs = (XSSFCellStyle) cs;
            if (bottomBorderColor != null)
                xcs.setBottomBorderColor((XSSFColor) bottomBorderColor);
            if (leftBorderColor != null)
                xcs.setLeftBorderColor((XSSFColor) leftBorderColor);
            if (rightBorderColor != null)
                xcs.setRightBorderColor((XSSFColor) rightBorderColor);
            if (topBorderColor != null)
                xcs.setTopBorderColor((XSSFColor) topBorderColor);
            // Per POI Javadocs, set foreground color first!
            if (fillForegroundColor != null)
                xcs.setFillForegroundColor((XSSFColor) fillForegroundColor);
            if (fillBackgroundColor != null)
                xcs.setFillBackgroundColor((XSSFColor) fillBackgroundColor);
        }
        cs.setFillPattern(fillPattern);
        return cs;
    }

    /**
     * Creates a new <code>Font</code> for the given <code>Workbook</code>,
     * with the given attributes.  Moved from <code>StyleTag</code> here for
     * 0.5.0.
     *
     * @param workbook           A <code>Workbook</code>.
     * @param fontBoldweight     A <code>short</code> boldweight constant.
     * @param fontItalic         Whether the text is italic.
     * @param fontColor          A color <code>Color</code> object.
     * @param fontName           A font name.
     * @param fontHeightInPoints A <code>short</code> font height in points.
     * @param fontUnderline      A <code>byte</code> underline constant.
     * @param fontStrikeout      Whether the font is strikeout.
     * @param fontCharset        An <code>int</code> charset constant.
     * @param fontTypeOffset     A <code>short</code> type offset constant.
     * @return A new <code>Font</code>.
     */
    public static Font createFont(Workbook workbook, boolean fontBoldweight, boolean fontItalic, Color fontColor, String fontName, short fontHeightInPoints, byte fontUnderline,
                                  boolean fontStrikeout, int fontCharset, short fontTypeOffset)
    {
        logger.trace("createFont: {},{},{},{},{},{},{},{},{}",
                fontBoldweight, fontItalic,
                ((fontColor == null) ? "null" :
                        (fontColor instanceof HSSFColor) ? fontColor.toString() :
                                ((XSSFColor) fontColor).getCTColor().toString()
                ), fontName, fontHeightInPoints, fontUnderline, fontStrikeout, fontCharset, fontTypeOffset);

        Font f = workbook.createFont();
        f.setBold(fontBoldweight);
        f.setItalic(fontItalic);
        f.setFontName(fontName);
        f.setFontHeightInPoints(fontHeightInPoints);
        f.setUnderline(fontUnderline);
        f.setStrikeout(fontStrikeout);
        f.setCharSet(fontCharset);
        f.setTypeOffset(fontTypeOffset);
        // Color type check.
        if (fontColor instanceof HSSFColor)
        {
            f.setColor(((HSSFColor) fontColor).getIndex());
        }
        else
        {
            // XSSFWorkbook
            XSSFFont xf = (XSSFFont) f;
            XSSFColor xssfFontColor = (XSSFColor) fontColor;
            if (xssfFontColor != null)
            {
                // As of Apache POI 3.8, there are Bugs 51236 and 52079 about font
                // color where somehow black and white get switched.  It appears to
                // have to do with the fact that black and white "theme" colors get
                // flipped.  Be careful, because XSSFColor(byte[]) does NOT call
                // "correctRGB", but XSSFColor.setRgb(byte[]) DOES call it, and so
                // does XSSFColor.getRgb(byte[]).
                // The private method "correctRGB" flips black and white, but no
                // other colors.  However, correctRGB is its own inverse operation,
                // i.e. correctRGB(correctRGB(rgb)) yields the same bytes as rgb.
                // XSSFFont.setColor(XSSFColor) calls "getRGB", but
                // XSSFCellStyle.set[Xx]BorderColor and
                // XSSFCellStyle.setFill[Xx]Color do NOT.
                // Solution: Let setColor correct a theme color on the way in.
                // Un-correct other colors, so that setColor will correct it.
                if (xssfFontColor.getCTColor().isSetTheme())
                    xf.setColor(xssfFontColor);
                else
                    xf.setColor(new XSSFColor(xssfFontColor.getCTColor()));
                // End of workaround for Bugs 51236 and 52079.
            }
        }

        return f;
    }

    /**
     * Determines the proper POI <code>Color</code>, given a string value that
     * could be a color name, e.g. "aqua", or a hex string, e.g. "#FFCCCC".
     *
     * @param workbook A <code>Workbook</code>, used only to determine whether
     *                 to create an <code>HSSFColor</code> or an <code>XSSFColor</code>.
     * @param value    The color value, which could be one of the 48 pre-defined
     *                 color names, or a hex value of the format "#RRGGBB".
     * @return A <code>Color</code>, or <code>null</code> if an invalid color
     * name was given.
     */
    public static Color getColor(Workbook workbook, String value)
    {
        logger.trace("getColor: {}", value);
        Color color = null;
        if (workbook instanceof HSSFWorkbook)
        {
            // Create an HSSFColor.
            if (value.startsWith("#"))
            {
                ExcelColor best = ExcelColor.AUTOMATIC;
                int minDist = 255 * 3;
                String strRed = value.substring(1, 3);
                String strGreen = value.substring(3, 5);
                String strBlue = value.substring(5, 7);
                int red = Integer.parseInt(strRed, 16);
                int green = Integer.parseInt(strGreen, 16);
                int blue = Integer.parseInt(strBlue, 16);
                // Hex value.  Find the closest defined color.
                for (ExcelColor excelColor : ExcelColor.values())
                {
                    int dist = excelColor.distance(red, green, blue);
                    if (dist < minDist)
                    {
                        best = excelColor;
                        minDist = dist;
                    }
                }
                color = best.getHssfColor();
                logger.debug("  Best HSSFColor found: {}", color);
            }
            else
            {
                // Treat it as a color name.
                try
                {
                    ExcelColor excelColor = ExcelColor.valueOf(value);
                    color = excelColor.getHssfColor();
                    logger.debug("  HSSFColor name matched: {}", value);
                }
                catch (IllegalArgumentException e)
                {
                    logger.warn("  HSSFColor name not matched ({}): {}", value, e.getMessage());
                }
            }
        }
        else // XSSFWorkbook
        {
            // Create an XSSFColor.
            if (value.startsWith("#") && value.length() == 7)
            {
                // Create the corresponding XSSFColor.
                color = new XSSFColor(new byte[] {
                        Integer.valueOf(value.substring(1, 3), 16).byteValue(),
                        Integer.valueOf(value.substring(3, 5), 16).byteValue(),
                        Integer.valueOf(value.substring(5, 7), 16).byteValue()
                },new DefaultIndexedColorMap());
                logger.debug("  XSSFColor created: {}", color);
            }
            else
            {
                // Create an XSSFColor from the RGB values of the desired color.
                try
                {
                    ExcelColor excelColor = ExcelColor.valueOf(value);

                    color = new XSSFColor(new byte[]
                            {(byte) excelColor.getRed(), (byte) excelColor.getGreen(), (byte) excelColor.getBlue()}
                    ,new DefaultIndexedColorMap());
                    logger.debug("  XSSFColor name matched: {}", value);
                }
                catch (IllegalArgumentException e)
                {
                    logger.debug("  XSSFColor name not matched ({}): {}", value, e.toString());
                }
            }
        }
        return color;
    }

    /**
     * <p>Returns a <code>String</code> formatted in the following way:</p>
     * <p>
     * <code>" at " + cellReference</code>
     * </p>
     * <p>e.g. <code>" at Sheet2!C3"</code>.</p>
     *
     * @param cell The <code>Cell</code>
     * @return The formatted location string.
     * @since 0.7.0
     */
    public static String getCellLocation(Cell cell)
    {
        if (cell == null)
            return "";
        return " at " + getCellKey(cell);
    }

    /**
     * <p>Returns a <code>String</code> formatted in the following way:</p>
     * <code>[", at " + tagCellRef + " (originally at " + origCellRef + ")"]+</code>
     * <p>where each instance represents the parent tag of the tag before it, in
     * a "tag stack trace" kind of way.</p>
     *
     * @param tag The <code>Tag</code>.
     * @return The formatted location string.
     * @since 0.9.0
     */
    public static String getTagLocationWithHierarchy(Tag tag)
    {
        if (tag == null)
            return "";

        StringBuilder buf = new StringBuilder();
        WorkbookContext workbookContext = tag.getWorkbookContext();
        Map<String, String> tagLocationsMap = workbookContext.getTagLocationsMap();
        do
        {
            TagContext tagContext = tag.getContext();
            Sheet sheet = tagContext.getSheet();
            Block block = tagContext.getBlock();
            int row = block.getTopRowNum();
            int col = block.getLeftColNum();
            String cellRef = new CellReference(sheet == null ? "DNE" : sheet.getSheetName(), row, col, false, false).formatAsString();
            String origCellRef = tagLocationsMap.get(cellRef);
            buf.append(System.getProperty("line.separator"));
            buf.append("  inside tag \"");
            buf.append(tag.getName());
            buf.append("\" (");
            buf.append(tag.getClass().getName());
            buf.append("), at ");
            buf.append(cellRef);
            if (origCellRef != null)
            {
                buf.append(" (originally at ");
                buf.append(origCellRef);
                buf.append(")");
            }

            tag = tag.getParentTag();
        }
        while (tag != null);
        return buf.toString();
    }

    /**
     * Sets the name of the indicated <code>Sheet</code> in the workbook to a
     * safe, legal sheet name.  Invalid characters are replaced with spaces.  If
     * a sheet name is already taken, numbers are added as suffixes until a name
     * that isn't taken is found, e.g. "example" -&gt; "example-1" -&gt; "example-2".
     *
     * @param workbook The <code>Workbook</code> in which to set a sheet's name.
     * @param index    The 0-based index of the <code>Sheet</code>.
     * @param newName  The proposed new name.
     * @return The actual safe name used to rename the sheet.
     * @since 0.9.1
     */
    public static String safeSetSheetName(Workbook workbook, int index, String newName)
    {
        logger.trace("sSSN({}, \"{}\")", index, newName);

        // For the uniqueness of sheet names, they are case insensitive.
        // No change.
        if (workbook.getSheetName(index).equalsIgnoreCase(newName))
        {
            return newName;
        }
        // Ensure it's a valid name.
        try
        {
            WorkbookUtil.validateSheetName(newName);
        }
        catch (IllegalArgumentException e)
        {
            newName = WorkbookUtil.createSafeSheetName(newName);
        }
        // Ensure the name isn't already in the workbook.
        boolean alreadyExists = true;
        int suffix = 0;
        String finalName = newName;
        while (alreadyExists)
        {
            if (suffix > 0)
            {
                int addedLength = String.valueOf(suffix).length() + 1;
                if (newName.length() + addedLength > 31)
                {
                    newName = newName.substring(0, 31 - addedLength);
                }
                finalName = newName + "-" + suffix;
            }
            alreadyExists = false;
            for (int s = 0; s < workbook.getNumberOfSheets(); s++)
            {
                // For the uniqueness of sheet names, they are case insensitive.
                if (finalName.equalsIgnoreCase(workbook.getSheetName(s)))
                {
                    alreadyExists = true;
                    suffix++;
                    break;
                }
            }
        }
        // Validated and doesn't already exist.
        workbook.setSheetName(index, finalName);
        return finalName;
    }
}

