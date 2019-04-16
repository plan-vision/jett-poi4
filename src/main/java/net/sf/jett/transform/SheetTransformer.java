package net.sf.jett.transform;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.event.SheetEvent;
import net.sf.jett.event.SheetListener;
import net.sf.jett.expression.Expression;
import net.sf.jett.expression.ExpressionFactory;
import net.sf.jett.formula.Formula;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.parser.FormulaParser;
import net.sf.jett.parser.TagParser;
import net.sf.jett.tag.TagContext;
import net.sf.jett.util.FormulaUtil;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>SheetTransformer</code> knows how to transform one
 * <code>Sheet</code> in an Excel spreadsheet.  For cell processing, it creates
 * a <code>Block</code> representing the entire <code>Sheet</code>, then it
 * delegates processing to a <code>BlockTransformer</code>.  It is also
 * responsible for gathering all <code>Formulas</code> at the beginning, and
 * replacing all <code>Formulas</code> with Excel Formulas at the end.  It also
 * exposes the "sheet" object in the "beans" <code>Map</code>.
 *
 * @author Randy Gettman
 */
public class SheetTransformer
{
    private static final Logger logger = LoggerFactory.getLogger(SheetTransformer.class);

    /**
     * Specifies a callback interface that is called after all off-sheet
     * properties are set.  This is only necessary so the
     * <code>ExcelTransformer</code> can safely apply these off-sheet properties
     * that XSSF doesn't retain after the sheet name is changed.
     * @since 0.7.0
     */
    public interface AfterOffSheetProperties
    {
        /**
         * Apply settings to the given <code>Sheet</code> after all off-sheet
         * properties have been transformed.
         * @param sheet The given <code>Sheet</code>.
         */
        public void applySettings(Sheet sheet);
    }

    /**
     * Transforms the given <code>Sheet</code>, using the given <code>Map</code>
     * of bean names to bean objects.
     * @param sheet The <code>Sheet</code> to transform.
     * @param context The <code>WorkbookContext</code>.
     * @param beans The beans map.
     */
    public void transform(Sheet sheet, WorkbookContext context, Map<String, Object> beans)
    {
        transform(sheet, context, beans, null);
    }

    /**
     * Transforms the given <code>Sheet</code>, using the given <code>Map</code>
     * of bean names to bean objects.
     * @param sheet The <code>Sheet</code> to transform.
     * @param context The <code>WorkbookContext</code>.
     * @param beans The beans map.
     * @param cloner An optional <code>SheetCloner</code>.  This
     *    is only present so the <code>ExcelTransformer</code>, as the caller of
     *    this method, can safely apply certain off-sheet properties that XSSF
     *    doesn't retain after the sheet name is changed.
     * @since 0.7.0
     */
    public void transform(Sheet sheet, WorkbookContext context, Map<String, Object> beans, SheetCloner cloner)
    {
        AfterOffSheetProperties callback = null;
        if (cloner != null)
        {
            callback = cloner.getMissingPropertiesSetter();
        }

        exposeSheet(beans, sheet);

        boolean shouldProceed = fireBeforeSheetProcessedEvent(context, sheet, beans);

        if (shouldProceed)
            beans = transformOffSheetProperties(sheet, beans, context, cloner);

        // This will happen regardless.
        if (callback != null)
            callback.applySettings(sheet);

        if (!shouldProceed)
            return;

        // Create a Block to encompass the entire sheet of Cells.
        // Create a Block as if there was a start tag at the beginning of the
        // text in the first column of the first row and an end tag in the last
        // populated column of the last row of the sheet.
        Block block = new Block(null, 0, SheetUtil.getLastPopulatedColIndex(sheet), 0, sheet.getLastRowNum());
        block.setDirection(Block.Direction.NONE);
        if (logger.isDebugEnabled())
        {
            logger.debug("Transforming sheet {}", sheet.getSheetName());

            Set<String> keys = beans.keySet();
            for (String key : keys)
            {
                logger.debug("  Key: {}", key);
                try
                {
                    logger.debug("    Value: {}", beans.get(key));
                }
                catch (RuntimeException e)
                {
                    logger.debug("    Value: {}: {}", e.getClass().getName(), e.getMessage());
                }
            }
        }

        TagContext tagContext = new TagContext();
        tagContext.setSheet(sheet);
        tagContext.setBlock(block);
        tagContext.setBeans(beans);
        tagContext.setProcessedCellsMap(new HashMap<String, Cell>());
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        tagContext.setMergedRegions(mergedRegions);
        readMergedRegions(sheet, mergedRegions);
        List<List<CellRangeAddress>> conditionalFormattingRegions = new ArrayList<>();
        tagContext.setConditionalFormattingRegions(conditionalFormattingRegions);
        readConditionalFormattingRegions(sheet, conditionalFormattingRegions);
        BlockTransformer transformer = new BlockTransformer();
        transformer.transform(tagContext, context);
        writeMergedRegions(sheet, mergedRegions);

        fireSheetProcessedEvent(context, sheet, beans);
    }

    /**
     * Transform any expressions in "off-sheet" properties, including header/
     * footer and the sheet name itself.
     * @param sheet The <code>Sheet</code> to transform.
     * @param beans The beans map.
     * @param context The <code>WorkbookContext</code>.
     * @param cloner The <code>SheetCloner</code>.
     * @return A beans <code>Map</code> to use for transformation.  It will
     *    usually be <code>beans</code>, but it may be a different map in case
     *    of implicit sheet cloning and sheet specific beans.
     * @since 0.7.0
     */
    private Map<String, Object> transformOffSheetProperties(Sheet sheet, Map<String, Object> beans,
                                                            WorkbookContext context, SheetCloner cloner)
    {
        String text;
        Object result;
        ExpressionFactory factory = context.getExpressionFactory();
        if (cloner == null)
        {
            cloner = new SheetCloner(sheet.getWorkbook());
        }

        // Implicit cloning is handled first; it may influence any expressions in
        // off sheet properties.
        text = sheet.getSheetName();
        List<String> collExprs = Expression.getImplicitCollectionExpr(text, beans, context);
        if (!collExprs.isEmpty())
        {
            logger.debug("Implicit collection processing in sheet name \"{}\"", text);
            beans = cloner.setupForImplicitCloning(sheet, beans, context);

            // Sheet name is expected to be changed due to implicit cloning;
            // pick it up again.
            text = sheet.getSheetName();
        }

        // Sheet name.
        result = Expression.evaluateString(text, factory, beans);
        Workbook workbook = sheet.getWorkbook();
        if (result != null)
        {
            String newSheetName = result.toString();
            if (!sheet.getSheetName().equals(newSheetName))
            {
                String oldSheetName = sheet.getSheetName();
                newSheetName = SheetUtil.safeSetSheetName(workbook, workbook.getSheetIndex(sheet), newSheetName);
                // Apache POI seems to update all Excel formulas on a sheet name change.
                // The only exception is on named ranges that are scoped to a different
                // sheet than the sheet being renamed, and only on XSSFSheets (looks like
                // an Apache POI bug; it works on HSSFSheets).  JETT won't be messing
                // around with actual Excel formulas here.

                // We still need to update all JETT formula references in the cell ref map.
                FormulaUtil.replaceSheetNameRefs(context, oldSheetName, newSheetName);
            }
        }

        // Header/footer.
        Header header = sheet.getHeader();
        text = header.getLeft();
        result = Expression.evaluateString(text, factory, beans);
        header.setLeft(result.toString());
        text = header.getCenter();
        result = Expression.evaluateString(text, factory, beans);
        header.setCenter(result.toString());
        text = header.getRight();
        result = Expression.evaluateString(text, factory, beans);
        header.setRight(result.toString());

        Footer footer = sheet.getFooter();
        text = footer.getLeft();
        result = Expression.evaluateString(text, factory, beans);
        footer.setLeft(result.toString());
        text = footer.getCenter();
        result = Expression.evaluateString(text, factory, beans);
        footer.setCenter(result.toString());
        text = footer.getRight();
        result = Expression.evaluateString(text, factory, beans);
        footer.setRight(result.toString());

        return beans;
    }

    /**
     * Searches for all <code>Formulas</code> contained on the given
     * <code>Sheet</code>.  Adds them to the given formula map.  Searches for
     * tags on the given <code>Sheet</code>.  Adds them to the given tag
     * locations map.
     *
     * @param sheet The <code>Sheet</code> on which to search for
     *    <code>Formulas</code>.
     * @param formulaMap A <code>Map</code> of strings to <code>Formulas</code>,
     *    with the keys of the format "sheetName!formulaText".
     * @param tagLocationsMap A <code>Map</code> of cell reference strings to
     *    original cell reference strings.
     */
    public void gatherFormulasAndTagLocations(Sheet sheet, Map<String, Formula> formulaMap,
                                              Map<String, String> tagLocationsMap)
    {
        int top = sheet.getFirstRowNum();
        int bottom = sheet.getLastRowNum();
        int left, right;
        String sheetName = sheet.getSheetName();
        FormulaParser parser = new FormulaParser();

        for (int rowNum = top; rowNum <= bottom; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            if (row != null)
            {
                left = row.getFirstCellNum();
                // For some reason, "getLastCellNum()" returns the last cell num "PLUS ONE".
                right = row.getLastCellNum() - 1;
                for (int cellNum = left; cellNum <= right; cellNum++)
                {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null && cell.getCellType() == CellType.STRING)
                    {
                        String cellText = cell.getStringCellValue();
                        if (cellText != null)
                        {
                            // Formula?
                            int formulaStartIdx = cellText.indexOf(Formula.BEGIN_FORMULA);
                            if (formulaStartIdx != -1)
                            {
                                int formulaEndIdx = FormulaUtil.getEndOfJettFormula(cellText, formulaStartIdx);
                                if (formulaEndIdx != -1)  // End token after Begin token
                                {
                                    // Grab the formula, begin and end tokens and all, e.g. $[SUM(C3)]
                                    cellText = cellText.substring(formulaStartIdx, formulaEndIdx + Formula.END_FORMULA.length());
                                    // Formula text is cell text without the begin and end tokens.
                                    String formulaText = cellText.substring(Formula.BEGIN_FORMULA.length(), formulaEndIdx - formulaStartIdx);
                                    parser.setFormulaText(formulaText);
                                    parser.setCell(cell);
                                    parser.parse();
                                    Formula formula = new Formula(cellText, parser.getCellReferences());
                                    String key = sheetName + "!" + cellText;
                                    logger.debug("gF: Formula found: {} => {}", key, formula);
                                    formulaMap.put(key, formula);
                                }
                            }

                            // Tag?
                            int tagStartIdx = cellText.indexOf(TagParser.BEGIN_START_TAG);
                            if (tagStartIdx != -1 && tagStartIdx < cellText.length() - 1)
                            {
                                char next = cellText.charAt(tagStartIdx + 1);
                                // "<" followed by not whitespace, "=", "<", ">", "\""
                                // THIS MATCHES WHAT TagParser LOOKS FOR TO DETERMINE IF
                                // IT'S THE START OF A TAG.
                                // Also, don't count "/", because that is an end tag.
                                if (!Character.isWhitespace(next) &&
                                        "=<>\"/".indexOf(next) == -1)
                                {
                                    String cellRef = new CellReference(sheet.getSheetName(),
                                            cell.getRowIndex(), cell.getColumnIndex(), false, false).formatAsString();
                                    logger.debug("gF: Tag text found: {} for {}", cellText, cellRef);
                                    tagLocationsMap.put(cellRef, cellRef);
                                }
                            }
                        }  // End if cell text isn't null
                    }
                }  // End loop on cells
            }
        }  // End loop on rows
    }

    /**
     * Replace all <code>Formulas</code> found in the given <code>Sheet</code>
     * with Excel formulas.
     * @param sheet The <code>Sheet</code>.
     * @param context The <code>WorkbookContext</code>.
     */
    public void replaceFormulas(Sheet sheet, WorkbookContext context)
    {
        int top = sheet.getFirstRowNum();
        int bottom = sheet.getLastRowNum();
        int left, right;
        String sheetName = sheet.getSheetName();
        Map<String, Formula> formulaMap = context.getFormulaMap();
        logger.debug("rF: Rows from {} to {}", top, bottom);

        for (int rowNum = top; rowNum <= bottom; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            if (row != null)
            {
                left = row.getFirstCellNum();
                // For some reason, "getLastCellNum()" returns the last cell num "PLUS ONE".
                right = row.getLastCellNum() - 1;
                for (int cellNum = left; cellNum <= right; cellNum++)
                {
                    Cell cell = row.getCell(cellNum);
                    if (cell != null && cell.getCellType() == CellType.STRING)
                    {
                        String cellText = cell.getStringCellValue();
                        if (cellText != null && cellText.startsWith(Formula.BEGIN_FORMULA) &&
                                cellText.endsWith(Formula.END_FORMULA))
                        {
                            // Don't consider any suffixes (e.g. "[0,0]") when looking
                            // up the Formula.
                            int idx = FormulaUtil.getEndOfJettFormula(cellText, 0);
                            String cellTextNoSfx = cellText.substring(0, idx + 1);
                            String key = sheetName + "!" + cellTextNoSfx;
                            Formula formula = formulaMap.get(key);
                            if (formula != null)
                            {
                                // Replace all original cell references with translated cell references.
                                String excelFormula = FormulaUtil.createExcelFormulaString(cellText, formula, sheetName, context);
                                logger.debug("  At {}, row {}, cell {}, replacing formula text \"{}\" with excel formula \"{}\".",
                                        sheetName, rowNum, cellNum, cellText, excelFormula);
                                cell.setCellFormula(excelFormula);
                            }
                        }
                    }
                }  // End cell for loop.
            }
        }  // End row for loop.
    }

    /**
     * Make the <code>Sheet</code> object available as bean in the given
     * <code>Map</code> of beans.
     * @param beans The <code>Map</code> of beans.
     * @param sheet The <code>Sheet</code> to expose.
     */
    private void exposeSheet(Map<String, Object> beans, Sheet sheet)
    {
        beans.put("sheet", sheet);
    }

    /**
     * Calls all <code>SheetListeners'</code> <code>beforeSheetProcessed</code>
     * method, sending a <code>SheetEvent</code>.
     * @param context The <code>WorkbookContext</code> object.
     * @param sheet The <code>Sheet</code> to be processed.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @return Whether processing of the <code>Sheet</code> should occur.  If
     *    any <code>SheetListener's</code> <code>beforeSheetProcessed</code>
     *    method returns <code>false</code>, then this method returns
     *    <code>false</code>.
     * @since 0.8.0
     */
    private boolean fireBeforeSheetProcessedEvent(WorkbookContext context, Sheet sheet, Map<String, Object> beans)
    {
        boolean shouldProceed = true;
        List<SheetListener> listeners = context.getSheetListeners();
        SheetEvent event = new SheetEvent(sheet, beans);
        for (SheetListener listener : listeners)
        {
            shouldProceed &= listener.beforeSheetProcessed(event);
        }
        return shouldProceed;
    }

    /**
     * Calls all <code>SheetListeners'</code> <code>sheetProcessed</code>
     * method, sending a <code>SheetEvent</code>.
     * @param context The <code>WorkbookContext</code> object.
     * @param sheet The <code>Sheet</code> to be processed.
     * @param beans A <code>Map</code> of bean names to bean values.
     * @since 0.8.0
     */
    private void fireSheetProcessedEvent(WorkbookContext context, Sheet sheet, Map<String, Object> beans)
    {
        List<SheetListener> listeners = context.getSheetListeners();
        SheetEvent event = new SheetEvent(sheet, beans);
        for (SheetListener listener : listeners)
        {
            listener.sheetProcessed(event);
        }
    }

    /**
     * Reads all merged regions from the given <code>Sheet</code> and populates
     * the given <code>List</code> with them.  All transformation that
     * manipulates merged regions will be done on this cache of merged regions,
     * instead of directly on the <code>Sheet</code>, for performance reasons.
     * @param sheet The <code>Sheet</code>.
     * @param mergedRegions A <code>List</code> of
     *    <code>CellRangeAddress</code>es, which is modified.
     * @since 0.8.0
     */
    private void readMergedRegions(Sheet sheet, List<CellRangeAddress> mergedRegions)
    {
        int numMergedRegions = sheet.getNumMergedRegions();
        for (int i = 0; i < numMergedRegions; i++)
        {
            mergedRegions.add(sheet.getMergedRegion(i));
        }
    }

    /**
     * Clears all merged regions on the given <code>Sheet</code> and populates
     * the <code>Sheet</code> with the given <code>List</code> of merged
     * regions.
     * @param sheet The <code>Sheet</code>.
     * @param mergedRegions A <code>List</code> of
     *    <code>CellRangeAddresses</code>.
     * @since 0.8.0
     */
    private void writeMergedRegions(Sheet sheet, List<CellRangeAddress> mergedRegions)
    {
        // Clear the existing merged regions on the sheet.
        // Remove them last item first, in an attempt to avoid internal ArrayList
        // shifting.
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--)
        {
            sheet.removeMergedRegion(i);
        }
        // Send in the new regions.
        for (CellRangeAddress mergedRegion : mergedRegions)
        {
            sheet.addMergedRegion(mergedRegion);
        }
    }

    /**
     * Reads all conditional formatting regions from the given <code>Sheet</code>
     * and populates the given <code>List</code> with them.  All transformation
     * that manipulates conditional formatting regions will be done on this
     * cache of regions, instead of directly on the <code>Sheet</code>.
     * @param sheet The <code>Sheet</code>.
     * @param regions A <code>List</code> of <code>Lists</code> of
     *    <code>CellRangeAddress</code>es, which is modified.
     * @since 0.9.0
     */
    private void readConditionalFormattingRegions(Sheet sheet, List<List<CellRangeAddress>> regions)
    {
        SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
        int numConditionalFormattings = scf.getNumConditionalFormattings();
        for (int i = 0; i < numConditionalFormattings; i++)
        {
            ConditionalFormatting cf = scf.getConditionalFormattingAt(i);
            CellRangeAddress[] ranges = cf.getFormattingRanges();
            List<CellRangeAddress> copies = new ArrayList<>(ranges.length);
            for (CellRangeAddress range : ranges)
            {
                copies.add(range.copy());
            }
            regions.add(copies);
        }
    }
}
