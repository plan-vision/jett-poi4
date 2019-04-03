package net.sf.jett.util;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.formula.SheetNameFormatter;

import net.sf.jett.formula.CellRef;
import net.sf.jett.formula.CellRefRange;
import net.sf.jett.formula.Formula;
import net.sf.jett.model.WorkbookContext;

/**
 * The <code>FormulaUtil</code> utility class provides methods for Excel
 * formula creation and manipulation.
 *
 * @author Randy Gettman
 */
public class FormulaUtil
{
    private static final Logger logger = LogManager.getLogger();

    // Prevents a mapping of "A1" => "A21" and "A2" => "A22" from yielding
    // "A1 + A2" => "A21 + A2" => "A221 + A22".
    private static final String NEGATIVE_LOOKBEHIND_ALPHA = "(?<![A-Za-z])";
    private static final String NEGATIVE_LOOKAHEAD_ALPHAN = "(?![A-Za-z0-9])";

    // Prefixes for cell keys in the cell ref map.
    /**
     * Prefix for explicit cell map references.
     * @since 0.9.1
     */
    public static final String EXPLICIT_REF_PREFIX = "e/";
    /**
     * Prefix for implicit cell map references.
     * @since 0.9.1
     */
    public static final String IMPLICIT_REF_PREFIX = "i/";

    /**
     * Finds unique cell references in all <code>Formulas</code> in the given
     * formula map.  The string "e/" (explicit) or "i/" (implicit) is prepended
     * to the cell key to distinguish when both a formula with an explicit sheet
     * name and another formula with an implicit sheet name would otherwise
     * resolve to the same cell key.
     * @param formulaMap A formula map.
     * @return A cell reference map, a <code>Map</code> of cell key strings to
     *    <code>Lists</code> of <code>CellRefs</code>.  Each <code>List</code>
     *    is initialized to contain only one <code>CellRef</code>, the original
     *    from the cell key string, e.g. "Sheet1!C2" =&gt; [Sheet1!C2]
     */
    public static Map<String, List<CellRef>> createCellRefMap(Map<String, Formula> formulaMap)
    {
        logger.trace("cCRM");
        Map<String, List<CellRef>> cellRefMap = new HashMap<>();
        for (String key : formulaMap.keySet())
        {
            Formula formula = formulaMap.get(key);
            logger.debug("  Processing key {} => {}", key, formula);

            // Formula keys always are of the format "Sheet!CellRef".
            // The key was created internally, so "!" is expected.
            String keySheetName = key.substring(0, key.indexOf("!"));

            for (CellRef cellRef : formula.getCellRefs())
            {
                String sheetName = cellRef.getSheetName();
                String cellKey = getCellKey(cellRef, keySheetName);
                if (sheetName != null)
                {
                    cellKey = EXPLICIT_REF_PREFIX + cellKey;
                }
                else
                {
                    cellKey = IMPLICIT_REF_PREFIX + cellKey;
                }
                if (!cellRefMap.containsKey(cellKey))
                {
                    List<CellRef> cellRefs = new ArrayList<>();
                    CellRef mappedCellRef;
                    if (sheetName == null || "".equals(sheetName))
                    {
                        // Local sheet reference.
                        mappedCellRef = new CellRef(cellRef.getRow(), cellRef.getCol(),
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                    }
                    else
                    {
                        // Refer to the template sheet name in the CellRef for now.
                        // If necessary, it will be translated into resultant sheet
                        // name(s) later in "updateSheetNameRefsAfterClone".
                        mappedCellRef = new CellRef(sheetName, cellRef.getRow(), cellRef.getCol(),
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                    }
                    cellRefs.add(mappedCellRef);
                    logger.debug("    New CellRefMap entry: {} => [{}]", cellKey, mappedCellRef.formatAsString());

                    cellRefMap.put(cellKey, cellRefs);
                }
            }
        }
        return cellRefMap;
    }

    /**
     * Creates a "cell key" from a cell ref, with a sheet name supplied if the
     * cell ref doesn't refer to a sheet name.  The returned string is suitable
     * as a key in the cell ref map.
     * @param cellRef The <code>CellRef</code>.
     * @param sheetName The sheet name to use if the <code>CellRef</code>
     *    doesn't supply one.
     * @return A string of the format "sheetName!cellRef", where "sheetName" is
     *    from the <code>CellRef</code> or it defaults to the
     *    <code>sheetName</code> parameter if it doesn't exist.  No single-
     *    quotes are in the cell key.
     * @since 0.8.0
     */
    public static String getCellKey(CellRef cellRef, String sheetName)
    {
        String cellKey;
        String cellRefSheetName = cellRef.getSheetName();
        // If no sheet name, then prepend the sheet name from the Formula key.
        if (cellRefSheetName == null || "".equals(cellRefSheetName))
        {
            // Prepend sheet name from formula key.
            cellKey = sheetName + "!" + cellRef.formatAsString();
        }
        else
        {
            // Single quotes may be in the cell reference.
            // Don't store single-quotes in cell key:
            // "'Test Sheet'!C3" => "Test Sheet!C3"
            cellKey = cellRef.formatAsString().replace("'", "");
        }
        return cellKey;
    }

    /**
     * Replaces cell references in the given formula text with the translated
     * cell references, and returns the formula string.
     * @param formula The <code>Formula</code>, for its access to its original
     *    <code>CellRefs</code>.
     * @param sheetName The name of the <code>Sheet</code> on which the formula
     *    exists.
     * @param context The <code>WorkbookContext</code>, for its access to the
     *    cell reference map.
     * @return A string suitable for an Excel formula, for use in the method
     *    <code>Cell.setCellFormula()</code>.
     * @since 0.9.1
     */
    public static String createExcelFormulaString(Formula formula,
                                                  String sheetName, WorkbookContext context)
    {
        return createExcelFormulaString(formula.getFormulaText(), formula, sheetName, context);
    }

    /**
     * Replaces cell references in the given formula text with the translated
     * cell references, and returns the formula string.
     * @param formulaText The <code>Formula</code> text, e.g. "SUM(C2)".
     * @param formula The <code>Formula</code>, for its access to its original
     *    <code>CellRefs</code>.
     * @param sheetName The name of the <code>Sheet</code> on which the formula
     *    exists.
     * @param context The <code>WorkbookContext</code>, for its access to the
     *    cell reference map.
     * @return A string suitable for an Excel formula, for use in the method
     *    <code>Cell.setCellFormula()</code>.
     */
    public static String createExcelFormulaString(String formulaText, Formula formula,
                                                  String sheetName, WorkbookContext context)
    {
        Map<String, List<CellRef>> cellRefMap = context.getCellRefMap();
        List<CellRef> origCellRefs = formula.getCellRefs();
        StringBuilder buf = new StringBuilder();
        String excelFormula, suffix;
        int endFormulaIdx = getEndOfJettFormula(formulaText, 0);
        int idx = formulaText.indexOf("[", endFormulaIdx);  // Get pos of any suffixes (e.g. "[0,0]").
        if (idx > -1)
        {
            excelFormula = formulaText.substring(0, idx);
            suffix = formulaText.substring(idx);
        }
        else
        {
            excelFormula = formulaText;
            suffix = "";
        }
        // Strip any $[ and ] off the Excel Formula, which at this point has been
        // stripped of any suffixes already.
        if (excelFormula.startsWith(Formula.BEGIN_FORMULA) && excelFormula.endsWith(Formula.END_FORMULA))
            excelFormula = excelFormula.substring(Formula.BEGIN_FORMULA.length(),
                    excelFormula.length() - Formula.END_FORMULA.length());

        logger.debug("cEFS: Formula text:\"{}\" on sheet {}", formulaText, sheetName);
        logger.debug("  excelFormula: \"{}\"", excelFormula);

        for (CellRef origCellRef : origCellRefs)
        {
            logger.debug("  Original cell ref: {}", origCellRef.formatAsString());

            // Look up the translated cells by cell key, which requires a sheet name.
            String cellKey = getCellKey(origCellRef, sheetName);
            boolean isExplicit = origCellRef.getSheetName() != null;
            if (!isExplicit)
            {
                cellKey = IMPLICIT_REF_PREFIX + cellKey;
            }
            else
            {
                cellKey = EXPLICIT_REF_PREFIX + cellKey;
            }
            // Append the suffix to the cell key to look up the correct references.
            cellKey += suffix;

            // Find the appropriate cell references.
            // It may be necessary to remove suffixes iteratively, if the cell key
            // represents a formula cell reference outside of a looping tag.
            List<CellRef> transCellRefs;
            do
            {
                transCellRefs = cellRefMap.get(cellKey);
                logger.debug("  cellKey: {} => {}", cellKey, transCellRefs);

                // Remove suffixes, one at a time if it's not found.
                if (transCellRefs == null)
                {
                    int lastSuffixIdx = cellKey.lastIndexOf("[");
                    if (lastSuffixIdx != -1)
                    {
                        cellKey = cellKey.substring(0, lastSuffixIdx);
                    }
                    else
                    {
                        throw new IllegalStateException("Unable to find cell references for cell key \"" + cellKey + "\"!");
                    }
                }
            }
            while (transCellRefs == null);

            // Construct the replacement string.
            String cellRefs;
            // Avoid re-allocation of the internal buffer.
            buf.delete(0, buf.length());
            int numCellRefs = transCellRefs.size();
            logger.debug("  Number of translated cell refs: {}", numCellRefs);
            if (numCellRefs > 0)
            {
                for (int i = 0; i < numCellRefs; i++)
                {
                    if (i > 0)
                        buf.append(",");
                    String cellRef = transCellRefs.get(i).formatAsString();
                    logger.debug("    Appending cell ref string: \"{}\".", cellRef);
                    buf.append(cellRef);
                }
                cellRefs = buf.toString();
            }
            else
            {
                // All cell references were deleted.  Must use the cell reference's
                // default value.  If that doesn't exist, that means that a default
                // value wasn't specified.  Use the "default" default.
                cellRefs = origCellRef.getDefaultValue();
                if (cellRefs == null)
                    cellRefs = CellRef.DEF_DEFAULT_VALUE;
                logger.debug("    Appending default value: \"{}\".", cellRefs);
            }
            // Replace the formula text, including any default value, with the
            // updated cell references.
            logger.debug("Regex: {}", NEGATIVE_LOOKBEHIND_ALPHA +
                    Pattern.quote(origCellRef.formatAsStringWithDef()) +
                    NEGATIVE_LOOKAHEAD_ALPHAN);
            logger.debug("cellRefs: {}", Matcher.quoteReplacement(cellRefs));

            excelFormula = excelFormula.replaceAll(
                    NEGATIVE_LOOKBEHIND_ALPHA +
                            Pattern.quote(origCellRef.formatAsStringWithDef()) +
                            NEGATIVE_LOOKAHEAD_ALPHAN,
                    Matcher.quoteReplacement(cellRefs));
        }
        return excelFormula;
    }

    /**
     * Examines all <code>CellRefs</code> in each <code>List</code>.  If a group
     * of <code>CellRefs</code> represent a linear range, horizontally or
     * vertically, then they are replaced with a <code>CellRefRange</code>.
     * @param cellRefMap The cell reference map.
     */
    public static void findAndReplaceCellRanges(Map<String, List<CellRef>> cellRefMap)
    {
        for (String key : cellRefMap.keySet())
        {
            List<CellRef> cellRefs = cellRefMap.get(key);
            // This will put cells that should be part of a range in consecutive
            // positions.
            Collections.sort(cellRefs);

            logger.debug("fARCR: Replacing cell ref ranges for \"{}\".", key);
            logger.debug("  cellRefs: {}", cellRefs);

            boolean vertical = false;
            boolean horizontal = false;
            CellRef first = null, prev = null;
            int firstIdx = -1;
            int size = cellRefs.size();

            for (int i = 0; i < size; i++)
            {
                CellRef curr = cellRefs.get(i);
                logger.debug("  curr is {}", curr.formatAsString());
                if (first == null)
                {
                    vertical = false;
                    horizontal = false;
                    first = curr;
                    firstIdx = i;
                    logger.debug("    Case first was null; first: {}, firstIdx = {}",
                            first.formatAsString(), firstIdx);
                }
                else if (vertical)
                {
                    logger.debug("    Case vertical; first: {}, firstIdx = {}",
                            first.formatAsString(), firstIdx);
                    if (!isBelow(prev, curr))
                    {
                        // End of range.  Replace sequence of vertically arranged
                        // CellRefs with a single CellRefRange.
                        replaceRange(cellRefs, firstIdx, i - 1);
                        // The list has shrunk.
                        int shrink = size - cellRefs.size();
                        size -= shrink;
                        i -= shrink;
                        // Setup for next range.
                        vertical = false;
                        first = curr;
                        firstIdx = i;
                    }
                }
                else if (horizontal)
                {
                    logger.debug("    Case horizontal; first: {}, firstIdx = {}",
                            first.formatAsString(), firstIdx);
                    if (!isRightOf(prev, curr))
                    {
                        // End of range.  Replace sequence of vertically arranged
                        // CellRefs with a single CellRefRange.
                        replaceRange(cellRefs, firstIdx, i - 1);
                        // The list has shrunk.
                        int shrink = size - cellRefs.size();
                        size -= shrink;
                        i -= shrink;
                        // Setup for next range.
                        horizontal = false;
                        first = curr;
                        firstIdx = i;
                    }
                }
                else
                {
                    // Decide on the proper direction, if any.
                    if (isRightOf(prev, curr))
                        horizontal = true;
                    else if (isBelow(prev, curr))
                        vertical = true;
                    else
                    {
                        first = curr;
                        firstIdx = i;
                    }
                    logger.debug("    Case none; first: {}, firstIdx = {}, horizontal = {}, vertical = {}",
                            first.formatAsString(), firstIdx, horizontal, vertical);
                }
                prev = curr;
            }

            // Don't forget the last one!
            if (horizontal || vertical)
                replaceRange(cellRefs, firstIdx, size - 1);
        }
    }

    /**
     * After sheets have been cloned, all sheets could have been renamed,
     * leaving the situation where in the cell ref map, all cell keys are of the
     * new sheet names, but the <code>CellRefs</code> still refer to the
     * template sheet names.  This updates all <code>CellRefs</code> in the cell
     * ref map to the new sheet names, cloning the references if necessary.
     * @param context The <code>WorkbookContext</code>, which contains the cell
     *    ref map, the template sheet names, and the new sheet names.
     * @since 0.8.0
     */
    public static void updateSheetNameRefsAfterClone(WorkbookContext context)
    {
        Map<String, List<CellRef>> cellRefMap = context.getCellRefMap();
        List<String> templateSheetNamesList = context.getTemplateSheetNames();
        List<String> newSheetNamesList = context.getSheetNames();
        logger.trace("uSNRAC...");
        for (String key : cellRefMap.keySet())
        {
            logger.debug("key: \"{}\".", key);

            // Formula keys always are of the format "e/Sheet!CellRef" or "i/Sheet!CellRef".
            // The key was created internally, so "!" is expected.
            // 1. e/templateSheet!cellKey => templateSheet!cellRef
            // Must update cell refs to resultant sheet cell ref(s).
            // 2. i/resultantSheet!cellKey => cellRef
            // Don't update these.
            // Determine if it's a template sheet name.
            boolean isExplicitRef = key.startsWith(EXPLICIT_REF_PREFIX);
            if (!isExplicitRef)
            {
                // Assumed to be a resultant/implicit sheet name; skip.
                continue;
            }
            // Bypass the explicit/implicit prefix.
            String templateSheetName = key.substring(2, key.indexOf("!"));
            // At this point, keySheetName is known to be a template sheet name.
            // No transformation has taken place yet, so there should be exactly
            // one cell ref in the list.
            List<CellRef> cellRefs = cellRefMap.get(key);
            List<CellRef> addedCellRefs = new ArrayList<>();
            CellRef cellRef = cellRefs.get(0);
            logger.debug("  cellRef: \"{}\".", cellRef);
            String templateRefSheetName = cellRef.getSheetName();

            // No cell ref sheet reference means a simple reference, e.g "B2",
            // meaning "this sheet", which means don't update.
            if (templateRefSheetName != null)
            {
                // Update the reference, plus clone the reference too, if more
                // than one template sheet name matches.
                boolean updatedFirstAlready = false;
                for (int j = 0; j < templateSheetNamesList.size(); j++)
                {
                    String sheetName = templateSheetNamesList.get(j);
                    if (sheetName.equals(templateSheetName))
                    {
                        String newSheetName = newSheetNamesList.get(j);
                        CellRef newCellRef = new CellRef(newSheetName, cellRef.getRow(), cellRef.getCol(),
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                        if (updatedFirstAlready)
                        {
                            logger.debug("    refers to other sheet: Adding \"{}\".", newCellRef);
                            addedCellRefs.add(newCellRef);
                        }
                        else
                        {
                            logger.debug("    refers to other sheet: Replacing \"{}\" with \"{}\" keyed by {}.",
                                    cellRef, newCellRef, key);
                            cellRefs.set(0, newCellRef);  // The only one so far.
                            updatedFirstAlready = true;
                        }
                    }
                }  // End for loop on template sheet names
            }  // End null check on templateSheetRefName
            cellRefs.addAll(addedCellRefs);
        }  // End for loop on cell keys.
    }

    /**
     * After a sheet has been implicitly cloned, there is a sheet that is
     * unaccounted for in the template sheet names, new sheet names, the formula
     * map, and the cell ref map.  This inserts new <code>CellRefs</code> in the
     * cell ref map to the new sheet name, adds new keys in the formula map and
     * the cell ref map, and inserts the "template" sheet name and new sheet
     * name.
     * @param context The <code>WorkbookContext</code>, which contains the cell
     *    ref map, the template sheet names, and the new sheet names.
     * @param origSheetName The current name of the <code>Sheet</code> that was
     *    copied.
     * @param newSheetName The new name of the <code>Sheet</code> that is a
     *    clone of the sheet that was copied.
     * @param clonePos The 0-based index of the sheet that is a clone of the
     *    sheet that was copied.
     * @since 0.9.1
     */
    public static void addSheetNameRefsAfterClone(WorkbookContext context, String origSheetName,
                                                  String newSheetName, int clonePos)
    {
        logger.trace("aSNRAC(context, {}, {}, {})", origSheetName, newSheetName, clonePos);

        // Insert into the template and new sheet name lists (local copies,
        // doesn't affect the original list passed in to ExcelTransformer.transform).
        List<String> templateSheetNames = context.getTemplateSheetNames();
        List<String> newSheetNames = context.getSheetNames();
        int index = newSheetNames.indexOf(origSheetName);
        if (index != -1)
        {
            newSheetNames.add(clonePos, newSheetName);
            templateSheetNames.add(clonePos, templateSheetNames.get(index));
        }

        // Formula map: insert new keys.  Make it look like these formulas were
        // here since the beginning of the transformation.
        Map<String, Formula> formulaMap = context.getFormulaMap();
        Map<String, Formula> addToFormulaMap = new HashMap<>();
        for (String key : formulaMap.keySet())
        {
            index = key.indexOf("!");  // Expected to be present in all cell keys
            String sheetPartOfKey = key.substring(0, index);
            if (sheetPartOfKey.equals(origSheetName))
            {
                Formula formula = formulaMap.get(key);
                String newKey = newSheetName + "!" + key.substring(index + 1);
                logger.debug("aSNRAC: Adding formula map key {} referring to formula {}", newKey, formula);
                addToFormulaMap.put(newKey, formula);
            }
        }
        formulaMap.putAll(addToFormulaMap);

        // Cell Ref Map:
        // Add keys for implicit cell references, copying the references.
        // Add cell refs to existing explicit cell references.
        Map<String, List<CellRef>> cellRefMap = context.getCellRefMap();
        Map<String, List<CellRef>> addToCellRefMap = new HashMap<>();
        for (String key : cellRefMap.keySet())
        {
            List<CellRef> cellRefs = cellRefMap.get(key);
            List<CellRef> newCellRefs;  // Set if a new entry will be made.
            index = key.indexOf("!");  // Expected to be present in all cell keys
            boolean isExplicit = key.startsWith(EXPLICIT_REF_PREFIX);
            // Bypass explicit/implicit indicator.
            String sheetPartOfKey = key.substring(2, index);
            // If the sheet name in the key changed, then we must replace the entry.
            // This occurs with JETT formulas in a sheet whose name was changed
            // via an expression, when those formulas refer to local sheet cells.
            if (!isExplicit && sheetPartOfKey.equals(origSheetName))
            {
                String newKey = IMPLICIT_REF_PREFIX + newSheetName + "!" + key.substring(index + 1);
                logger.debug("aSNRAC: Adding cell ref map key {} referring to {}", newKey, cellRefs);
                // Shallow copy is ok; CellRefs aren't changed; they are replaced if needed.
                newCellRefs = new ArrayList<>(cellRefs);
                addToCellRefMap.put(newKey, newCellRefs);
            }
            else
            {
                List<CellRef> addToCellRefs = new ArrayList<>();
                for (int i = 0; i < cellRefs.size(); i++)
                {
                    CellRef cellRef = cellRefs.get(i);
                    String cellRefSheetName = cellRef.getSheetName();
                    if (cellRefSheetName != null && cellRefSheetName.equals(origSheetName))
                    {
                        CellRef newCellRef = new CellRef(newSheetName, cellRef.getRow(), cellRef.getCol(),
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                        logger.debug("aSNRAC: adding cell ref {} to list keyed by {}", newCellRef, key);
                        addToCellRefs.add(i, newCellRef);
                    }
                }
                cellRefs.addAll(addToCellRefs);
            }
        }
        // Add the new entries.
        cellRefMap.putAll(addToCellRefMap);
    }

    /**
     * When a <code>Sheet</code> is renamed, then this updates all
     * <code>CellRefs</code> in the cell reference map need to be updated too.
     * @param context The <code>WorkbookContext</code>, on which the formula map
     *    and the cell ref map can be found.
     * @param oldSheetName The old sheet name.
     * @param newSheetName The new sheet name.
     * @since 0.8.0
     */
    public static void replaceSheetNameRefs(WorkbookContext context, String oldSheetName, String newSheetName)
    {
        // Update new sheet name list (local copy, doesn't affect the original
        // list passed in to ExcelTransformer.transform).
        List<String> newSheetNames = context.getSheetNames();
        int index = newSheetNames.indexOf(oldSheetName);
        if (index != -1)
        {
            newSheetNames.set(index, newSheetName);
        }

        // Formula map: update keys.
        Map<String, Formula> formulaMap = context.getFormulaMap();
        List<String> removeFromFormulaMap = new ArrayList<>();
        Map<String, Formula> addToFormulaMap = new HashMap<>();
        for (String key : formulaMap.keySet())
        {
            index = key.indexOf("!");  // Expected to be present in all cell keys
            String sheetPartOfKey = key.substring(0, index);
            if (sheetPartOfKey.equals(oldSheetName))
            {
                Formula formula = formulaMap.get(key);
                removeFromFormulaMap.add(key);
                String newKey = newSheetName + "!" + key.substring(index + 1);
                logger.debug("rSNR: Replacing formula map key {} with {}", key, newKey);
                addToFormulaMap.put(newKey, formula);
            }
        }
        // Now remove all the keys marked to be removed, now that we're past the
        // Iterator and a possible ConcurrentModificationException.
        for (String key : removeFromFormulaMap)
        {
            formulaMap.remove(key);
        }
        // Put back all the replacements.
        formulaMap.putAll(addToFormulaMap);

        // Cell Ref Map: update keys and cell refs.
        Map<String, List<CellRef>> cellRefMap = context.getCellRefMap();
        List<String> removeFromCellRefMap = new ArrayList<>();
        Map<String, List<CellRef>> addToCellRefMap = new HashMap<>();
        for (String key : cellRefMap.keySet())
        {
            List<CellRef> cellRefs = cellRefMap.get(key);
            index = key.indexOf("!");  // Expected to be present in all cell keys
            boolean isExplicit = key.startsWith(EXPLICIT_REF_PREFIX);
            String sheetPartOfKey = key.substring(2, index);
            // If the sheet name in the key changed, then we must replace the entry.
            // This occurs with JETT formulas in a sheet whose name was changed
            // via an expression, when those formulas refer to local sheet cells.
            if (!isExplicit && sheetPartOfKey.equals(oldSheetName))
            {
                removeFromCellRefMap.add(key);
                String newKey = IMPLICIT_REF_PREFIX + newSheetName + "!" + key.substring(index + 1);
                logger.debug("rSNR: Replacing cell ref map key {} with {}", key, newKey);
                addToCellRefMap.put(newKey, cellRefs);
            }
            else
            {
                for (int i = 0; i < cellRefs.size(); i++)
                {
                    CellRef cellRef = cellRefs.get(i);
                    String cellRefSheetName = cellRef.getSheetName();
                    if (cellRefSheetName != null && cellRefSheetName.equals(oldSheetName))
                    {
                        CellRef newCellRef = new CellRef(newSheetName, cellRef.getRow(), cellRef.getCol(),
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                        logger.debug("rSNR: replacing cell ref {} with {} for key {}", cellRef, newCellRef, key);
                        cellRefs.set(i, newCellRef);
                    }
                }
            }
        }
        // Now remove all the keys marked to be removed, now that we're past the
        // Iterator and a possible ConcurrentModificationException.
        for (String key : removeFromCellRefMap)
        {
            cellRefMap.remove(key);
        }
        // Put back all the replacements.
        cellRefMap.putAll(addToCellRefMap);
    }

    /**
     * Returns <code>true</code> if <code>curr</code> is directly to the right
     * of <code>prev</code>, i.e., all of the following are true:
     * <ul>
     * <li>The sheet names match or they are both <code>null</code>.
     * <li>The row indexes match.
     * <li>The column index of <code>curr</code> is one more than the column
     *    index of <code>prev</code>.
     * </ul>
     * @param prev The previous <code>CellRef</code>.
     * @param curr The current <code>CellRef</code>.
     * @return <code>true</code> if <code>curr</code> is directly to the right
     *    of <code>prev</code>, else <code>false</code>.
     */
    private static boolean isRightOf(CellRef prev, CellRef curr)
    {
        return (curr.getRow() == prev.getRow() && curr.getCol() == prev.getCol() + 1 &&
                ((curr.getSheetName() == null && prev.getSheetName() == null) ||
                        (curr.getSheetName() != null && curr.getSheetName().equals(prev.getSheetName()))));
    }

    /**
     * Returns <code>true</code> if <code>curr</code> is directly below
     * <code>prev</code>, i.e., all of the following are true:
     * <ul>
     * <li>The sheet names match or they are both <code>null</code>.
     * <li>The column indexes match.
     * <li>The row index of <code>curr</code> is one more than the row
     *    index of <code>prev</code>.
     * </ul>
     * @param prev The previous <code>CellRef</code>.
     * @param curr The current <code>CellRef</code>.
     * @return <code>true</code> if <code>curr</code> is directly below
     *    <code>prev</code>, else <code>false</code>.
     */
    private static boolean isBelow(CellRef prev, CellRef curr)
    {
        return (curr.getCol() == prev.getCol() && curr.getRow() == prev.getRow() + 1 &&
                ((curr.getSheetName() == null && prev.getSheetName() == null) ||
                        (curr.getSheetName() != null && curr.getSheetName().equals(prev.getSheetName()))));
    }

    /**
     * Replace the <code>CellRefs</code> in the given <code>List</code> of
     * <code>CellRefs</code>, in the range of indexes between
     * <code>startIdx</code> and <code>endIdx</code> with a single
     * <code>CellRefRange</code>.
     * @param cellRefs Modifies this <code>List</code> of <code>CellRefs</code>.
     * @param startIdx The <code>CellRef</code> at this index is treated as the
     *    start of the range (inclusive).
     * @param endIdx The <code>CellRef</code> at this index is treated as the
     *    end of the range (inclusive).
     */
    private static void replaceRange(List<CellRef> cellRefs, int startIdx, int endIdx)
    {
        // Create the range.
        CellRef first = cellRefs.get(startIdx);
        CellRef prev = cellRefs.get(endIdx);
        CellRefRange range = new CellRefRange(first.getSheetName(), first.getRow(), first.getCol(),
                first.isRowAbsolute(), first.isColAbsolute());
        range.setRangeEndCellRef(prev);
        logger.debug("  Replacing {} through {} with {}",
                first.formatAsString(), prev.formatAsString(), range.formatAsString());
        // Replace the first with the range.
        cellRefs.set(startIdx, range);
        // Remove the others in the range.  The end index for the "subList"
        // method is exclusive.
        cellRefs.subList(startIdx + 1, endIdx + 1).clear();
    }

    /**
     * Shifts all <code>CellRefs</code> that are in range and on the same
     * <code>Sheet</code> by the given number of rows and/or columns (usually
     * one of those two will be zero).  Modifies the <code>Lists</code> that are
     * the values of <code>cellRefMap</code>.
     * @param sheetName The name of the <code>Sheet</code> on which to shift
     *    cell references.
     * @param context The <code>WorkbookContext</code> which holds the cell ref
     *    map, template sheet names, and new sheet names.
     * @param left The 0-based index of the column on which to start shifting
     *    cell references.
     * @param right The 0-based index of the column on which to end shifting
     *    cell references.
     * @param top The 0-based index of the row on which to start shifting
     *    cell references.
     * @param bottom The 0-based index of the row on which to end shifting
     *    cell references.
     * @param numCols The number of columns to shift the cell reference (can be
     *    negative).
     * @param numRows The number of rows to shift the cell reference (can be
     *    negative).
     * @param remove Determines whether to remove the old cell reference,
     *    resulting in a shift, or not to remove the old cell reference,
     *    resulting in a copy.
     * @param add Determines whether to add the new cell reference, resulting in
     *    a copy, or not to add the new cell reference, resulting in a shift.
     */
    public static void shiftCellReferencesInRange(String sheetName, WorkbookContext context,
                                                  int left, int right, int top, int bottom, int numCols, int numRows,
                                                  boolean remove, boolean add)
    {
        logger.trace("    sCRIR: left {}, right {}, top {}, bottom {}, numCols {}, numRows {}, remove {}, add {}",
                left, right, top, bottom, numCols, numRows, remove, add);
        Map<String, List<CellRef>> cellRefMap = context.getCellRefMap();
        List<String> templateSheetNames = context.getTemplateSheetNames();
        List<String> newSheetNames = context.getSheetNames();
        if (numCols == 0 && numRows == 0 && remove && add)
            return;
        for (String cellKey : cellRefMap.keySet())
        {
            // All cell keys have the sheet name in them.
            boolean isExplicit = cellKey.startsWith(EXPLICIT_REF_PREFIX);
            // Bypass explicit/implicit indicator.
            String keySheetName = cellKey.substring(2, cellKey.indexOf("!"));
            if (!keySheetName.equals(sheetName))
            {
                // No exact match.  Check the corresponding template sheet name, if
                // it exists.
                int index = newSheetNames.indexOf(sheetName);
                if (isExplicit || (index != -1 && keySheetName.equals(templateSheetNames.get(index))))
                {
                    // Template sheet name match.
                    // Update keySheetName (the template sheet name) to the new sheet name.
                    keySheetName = sheetName;
                }
                else
                {
                    continue;
                }
            }

            List<CellRef> cellRefs = cellRefMap.get(cellKey);
            List<CellRef> delete = new ArrayList<>();
            List<CellRef> insert = new ArrayList<>();
            for (CellRef cellRef : cellRefs)
            {
                String cellRefSheetName = cellRef.getSheetName();
                int row = cellRef.getRow();
                int col = cellRef.getCol();
                if ((cellRefSheetName == null || keySheetName.equals(cellRefSheetName)) &&
                        (row >= top && row <= bottom && col >= left && col <= right))
                {
                    if (remove)
                    {
                        logger.debug("      Deleting cell reference: {} for cell key {}", cellRef.formatAsString(), cellKey);
                        delete.add(cellRef);
                    }
                    if (add)
                    {
                        CellRef adjCellRef = new CellRef(cellRefSheetName, row + numRows, col + numCols,
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                        logger.debug("      Adding cell reference: {} for cell key {}", adjCellRef.formatAsString(), cellKey);
                        insert.add(adjCellRef);
                    }
                }
            }
            if (remove)
                cellRefs.removeAll(delete);
            if (add)
                cellRefs.addAll(insert);
        }
    }

    /**
     * Copies cell references that are on the same <code>Sheet</code> in the
     * given cell reference map by the given number of rows and/or columns
     * (usually one of those two will be zero).  Modifies the <code>Lists</code>
     * that are the values of <code>cellRefMap</code>.
     * @param sheetName The name of the <code>Sheet</code> on which to copy
     *    references.
     * @param context The <code>WorkbookContext</code> which holds the cell ref
     *    map, template sheet names, and new sheet names.
     * @param left The 0-based index of the column on which to start shifting
     *    cell references.
     * @param right The 0-based index of the column on which to end shifting
     *    cell references.
     * @param top The 0-based index of the row on which to start shifting
     *    cell references.
     * @param bottom The 0-based index of the row on which to end shifting
     *    cell references.
     * @param numCols The number of columns to shift the cell reference (can be
     *    negative).
     * @param numRows The number of rows to shift the cell reference (can be
     *    negative).
     * @param currSuffix The current "[loop,iter]*" suffix we're already in.
     * @param newSuffix The new "[loop,iter]" suffix to add for new entries.
     */
    public static void copyCellReferencesInRange(String sheetName, WorkbookContext context,
                                                 int left, int right, int top, int bottom, int numCols, int numRows, String currSuffix, String newSuffix)
    {
        logger.trace("    cCRIR: left {}, right {}, top {}, bottom {}, numCols {}, numRows {}, currSuffix: \"{}\", newSuffix: \"{}\"",
                left, right, top, bottom, numCols, numRows, currSuffix, newSuffix);
        Map<String, List<CellRef>> cellRefMap = context.getCellRefMap();
        Map<String, List<CellRef>> newCellRefEntries = new HashMap<>();
        List<String> templateSheetNames = context.getTemplateSheetNames();
        List<String> newSheetNames = context.getSheetNames();
        for (String cellKey : cellRefMap.keySet())
        {
            // All cell keys have the sheet name in them.
            boolean isExplicit = cellKey.startsWith(EXPLICIT_REF_PREFIX);
            // Bypass the explicit/implicit indicator.
            String keySheetName = cellKey.substring(2, cellKey.indexOf("!"));
            if (!keySheetName.equals(sheetName))
            {
                // No exact match.  Check the corresponding template sheet name, if
                // it exists.
                int index = newSheetNames.indexOf(sheetName);
                if (isExplicit || (index != -1 && keySheetName.equals(templateSheetNames.get(index))))
                {
                    // Template sheet name match.
                    // Update keySheetName (the template sheet name) to the new sheet name.
                    keySheetName = sheetName;
                }
                else
                {
                    continue;
                }
            }

            // A cell key may have a suffix, e.g. [0,1].
            String keySuffix = "";
            int idx = cellKey.indexOf("[");
            if (idx > -1)
                keySuffix = cellKey.substring(idx);
            if (currSuffix.startsWith(keySuffix)) // Suffix matches
            {
                List<CellRef> cellRefs = cellRefMap.get(cellKey);
                List<CellRef> insert = new ArrayList<>();
                for (CellRef cellRef : cellRefs)
                {
                    String cellRefSheetName = cellRef.getSheetName();
                    int row = cellRef.getRow();
                    int col = cellRef.getCol();
                    if ((cellRefSheetName == null || keySheetName.equals(cellRefSheetName)) &&    // Sheet matches
                            (row >= top && row <= bottom && col >= left && col <= right))             // In cell range
                    {
                        CellRef adjCellRef = new CellRef(cellRefSheetName, row + numRows, col + numCols,
                                cellRef.isRowAbsolute(), cellRef.isColAbsolute());
                        // Only add the reference if being translated!
                        if (numRows != 0 || numCols != 0)
                        {
                            logger.debug("      Adding cell reference: {} for cell key {}", adjCellRef.formatAsString(), cellKey);
                            insert.add(adjCellRef);
                        }
                        // Introduce new mappings with the new suffix, e.g. [2,0], appended to
                        // the current suffix, e.g. [0,1][2,0].
                        // Look for formulas in the range.
                        // Only do this once (pick out those without suffixes to accomplish this).
                        if (idx == -1)
                        {
                            String newCellKey = cellKey + currSuffix + newSuffix;
                            List<CellRef> newCellRefs = new ArrayList<>();
                            newCellRefs.add(adjCellRef);
                            logger.debug("      Adding new entry: {} => [{}]", newCellKey, adjCellRef.formatAsString());
                            newCellRefEntries.put(newCellKey, newCellRefs);
                        }
                    }
                }
                cellRefs.addAll(insert);
            }
        }
        cellRefMap.putAll(newCellRefEntries);
    }

    /**
     * Finds the end of the JETT formula substring.  This accounts for bracket
     * characters (<code>[]</code>) that may be nested inside the JETT formula;
     * they are legal characters in Excel formulas.  It also accounts for Excel
     * string literals, by ignoring bracket characters inside Excel string
     * literals, which are enclosed in double-quotes.  Note that escaped
     * double-quote characters (<code>""</code>) don't change the "inside double
     * quotes" variable, once both double-quotes have been processed.
     * @param cellText The cell text.
     * @param formulaStartIdx The start of the formula.
     * @return The index of the ']' character that ends the JETT formula, or
     *    <code>-1</code> if not found.
     * @since 0.9.1
     */
    public static int getEndOfJettFormula(String cellText, int formulaStartIdx)
    {
        int numUnMatchedBrackets = 0;
        boolean insideDoubleQuotes = false;

        for (int i = formulaStartIdx + Formula.BEGIN_FORMULA.length(); i < cellText.length(); i++)
        {
            char c = cellText.charAt(i);
            switch (c)
            {
            case '[':
                if (!insideDoubleQuotes)
                {
                    numUnMatchedBrackets++;
                }
                break;
            case ']':
                if (!insideDoubleQuotes)
                {
                    if (numUnMatchedBrackets == 0)
                        return i;

                    numUnMatchedBrackets--;
                }
                break;
            case '"':
                insideDoubleQuotes = !insideDoubleQuotes;
                break;
            default:
                break;
            }
        }
        // End of cell text without matching end-bracket.  Not found.
        return -1;
    }

    /**
     * It's possible that a JETT formula was entered that wouldn't be accepted
     * by Excel because the sheet name needs to be formatted -- enclosed in
     * single quotes, e.g. <code>$[SUM(${dvs.name}$@i=n;l=10;v=s;r=DNE!B3)]</code>
     * -&gt; <code>$[SUM('${dvs.name}$@i=n;l=10;v=s;r=DNE'!B3)]</code>
     * @param formula The original JETT formula text, as entered in the template.
     * @param cellReferences The <code>List</code> of <code>CellRefs</code>
     *    already found by the <code>FormulaParser</code>.
     * @return Formula text with sheet names formatted properly for Excel.
     * @since 0.9.1
     */
    public static String formatSheetNames(String formula, List<CellRef> cellReferences)
    {
        for (CellRef cellRef : cellReferences)
        {
            String sheetName = cellRef.getSheetName();
            if (sheetName != null && !sheetName.startsWith("'") && !sheetName.endsWith("'"))
            {
                String formattedSheetName = SheetNameFormatter.format(sheetName);
                // If not already in single quotes.
                formula = formula.replaceAll("(?<!')" + Pattern.quote(sheetName) + "(?!')", Matcher.quoteReplacement(formattedSheetName));
            }
        }
        return formula;
    }
}
