package net.sf.jett.transform;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import net.sf.jett.exception.MetadataParseException;
import net.sf.jett.expression.Expression;
import net.sf.jett.model.BaseLoopTagStatus;
import net.sf.jett.model.HashMapWrapper;
import net.sf.jett.model.MissingCloneSheetProperties;
import net.sf.jett.model.PastEndValue;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.parser.MetadataParser;
import net.sf.jett.parser.SheetNameMetadataParser;
import net.sf.jett.tag.NameTag;
import net.sf.jett.util.FormulaUtil;
import net.sf.jett.util.RichTextStringUtil;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>SheetCloner</code> clone sheets and can set them up for implicit
 * collections processing when a collection is detected as part of an
 * expression in a sheet name.
 *
 * @author Randy Gettman
 * @since 0.9.1
 */
public class SheetCloner
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Determines the beginning of metadata text inside sheet names only.  It's
     * different from the beginning of metadata text inside normal cells
     * ({@link net.sf.jett.transform.CollectionsTransformer#BEGIN_METADATA})
     * because the normal cells beginning of metadata contains a character that
     * isn't allowed in Excel sheet names.
     */
    public static final String BEGIN_METADATA = "$@";

    private Workbook myWorkbook;
    private List<MissingCloneSheetProperties> myMissingPropertiesList;

    /**
     * Constructs an <code>SheetCloner</code> that will work on cloning
     * the given source <code>Sheet</code>, in the given <code>WorkbookContext</code>.
     * @param workbook The <code>Workbook</code> that contains <code>Sheets</code>
     *    that may be cloned/removed.
     */
    public SheetCloner(Workbook workbook)
    {
        myWorkbook = workbook;
        myMissingPropertiesList = new ArrayList<>();
    }

    /**
     * Returns an object that can set the missing properties on a <code>Sheet</code>
     * after it has been moved and/or renamed.  This was moved to
     * <code>SheetCloner</code> for version 0.9.1.
     * @return An object that can set the missing properties on a <code>Sheet</code>.
     */
    public SheetTransformer.AfterOffSheetProperties getMissingPropertiesSetter()
    {
        return new SheetTransformer.AfterOffSheetProperties() {
            /**
             * Apply the missing clone sheet properties.
             * @param sheet The given <code>Sheet</code>.
             * @since 0.7.0
             */
            @Override
            public void applySettings(Sheet sheet)
            {
                replaceMissingCloneSheetProperties(sheet, myMissingPropertiesList.get(sheet.getWorkbook().getSheetIndex(sheet)));
            }
        };
    }

    /**
     * Copies the properties that weren't properly copied upon cloning and/or
     * moving the given <code>Sheet</code> back into the sheet.  This was moved
     * to <code>SheetCloner</code> for version 0.9.1.
     * @param sheet The <code>Sheet</code> on which to restore properties.
     * @param mcsp The properties to copy back to the sheet.
     * @since 0.7.0
     */
    private void replaceMissingCloneSheetProperties(Sheet sheet, MissingCloneSheetProperties mcsp)
    {
        PrintSetup ps = sheet.getPrintSetup();

        // Missing properties for any case.
        sheet.setRepeatingColumns(mcsp.getRepeatingColumns());
        sheet.setRepeatingRows(mcsp.getRepeatingRows());

        // Missing properties for XSSF only.
        if (sheet instanceof XSSFSheet)
        {
            ps.setCopies(mcsp.getCopies());
            ps.setDraft(mcsp.isDraft());
            ps.setFitHeight(mcsp.getFitHeight());
            ps.setFitWidth(mcsp.getFitWidth());
            ps.setHResolution(mcsp.getHResolution());
            ps.setLandscape(mcsp.isLandscape());
            ps.setLeftToRight(mcsp.isLeftToRight());
            ps.setNoColor(mcsp.isNoColor());
            ps.setNotes(mcsp.isNotes());
            ps.setPageStart(mcsp.getPageStart());
            ps.setPaperSize(mcsp.getPaperSize());
            ps.setScale(mcsp.getScale());
            ps.setUsePage(mcsp.isUsePage());
            ps.setValidSettings(mcsp.isValidSettings());
            ps.setVResolution(mcsp.getVResolution());
        }
    }

    /**
     * Clones and moves <code>Sheets</code> around as necessary according to
     * the given template and new sheet names lists.  This method initializes
     * the list of Missing Clone Sheet Properties.  The logic for this method
     * was extracted out of <code>ExcelTransformer</code> for version 0.9.1.
     * @param templateSheetNames The <code>List</code> of template sheet names.
     * @param newSheetNames The <code>List</code> of new sheet names.
     */
    public void cloneForSheetSpecificBeans(List<String> templateSheetNames, List<String> newSheetNames)
    {
        Map<String, Integer> firstReferencedSheets = new HashMap<>();
        // Note down any sheet properties that are known to be "messed up" when a
        // Sheet is cloned and/or moved.
        for (int i = 0; i < myWorkbook.getNumberOfSheets(); i++)
        {
            myMissingPropertiesList.add(getMissingCloneSheetProperties(myWorkbook.getSheetAt(i)));
        }

        // Clone and/or move sheets.
        for (int i = 0; i < templateSheetNames.size(); i++)
        {
            if (logger.isTraceEnabled())
            {
                for (int j = 0; j < myWorkbook.getNumberOfSheets(); j++)
                    logger.trace("  Before: Sheet({}): \"{}\".",
                            j, myWorkbook.getSheetAt(j).getSheetName());
            }

            String templateSheetName = templateSheetNames.get(i);
            String newSheetName = newSheetNames.get(i);
            if (firstReferencedSheets.containsKey(templateSheetName))
            {
                int prevIndex = firstReferencedSheets.get(templateSheetName);
                // Clone the previously referenced sheet, name it, and reposition it.
                logger.debug("Cloning sheet at position {}.", prevIndex);

                MissingCloneSheetProperties cloned = new MissingCloneSheetProperties(myMissingPropertiesList.get(prevIndex));

                myWorkbook.cloneSheet(prevIndex);
                logger.debug("Setting sheet name at position {} to \"{}\".",
                        myWorkbook.getNumberOfSheets() - 1, newSheetName);

                int clonePos = myWorkbook.getNumberOfSheets() - 1;
                newSheetName = SheetUtil.safeSetSheetName(myWorkbook, clonePos, newSheetName);
                cloneNamedRanges(myWorkbook, prevIndex);

                logger.debug("Moving sheet \"{}\" to position {}.",
                        newSheetName, i);

                myWorkbook.setSheetOrder(newSheetName, i);
                updateNamedRangesScope(myWorkbook, clonePos, i);

                myMissingPropertiesList.add(i, cloned);
            }
            else
            {
                // Find the sheet.
                int index = myWorkbook.getSheetIndex(templateSheetName);
                if (index == -1)
                    throw new RuntimeException("Template Sheet \"" + templateSheetName + "\" not found!");

                // Rename the sheet and move it to the current position.
                logger.debug("Renaming sheet at position {} to \"{}\".",
                        index, newSheetName);

                newSheetName = SheetUtil.safeSetSheetName(myWorkbook, index, newSheetName);

                if (index != i)
                {
                    logger.debug("Moving sheet at position {} to {}.", index, i);

                    MissingCloneSheetProperties move = myMissingPropertiesList.remove(index);

                    myWorkbook.setSheetOrder(newSheetName, i);
                    updateNamedRangesScope(myWorkbook, index, i);

                    myMissingPropertiesList.add(i, move);
                }
                firstReferencedSheets.put(templateSheetName, i);
            }
            if (logger.isTraceEnabled())
            {
                for (int j = 0; j < myWorkbook.getNumberOfSheets(); j++)
                    logger.trace("  After: Sheet({}): \"{}\".",
                            j, myWorkbook.getSheetAt(j).getSheetName());
            }
        }
    }

    /**
     * Copies the properties that won't be properly copied upon cloning and/or
     * moving the given <code>Sheet</code>.  This was moved to
     * <code>SheetCloner</code> for version 0.9.1.
     * @param sheet The <code>Sheet</code> on which to copy properties.
     * @return A <code>MissingCloneSheetProperties</code>.
     * @since 0.7.0
     */
    private MissingCloneSheetProperties getMissingCloneSheetProperties(Sheet sheet)
    {
        MissingCloneSheetProperties mcsp = new MissingCloneSheetProperties();
        PrintSetup ps = sheet.getPrintSetup();

        mcsp.setRepeatingColumns(sheet.getRepeatingColumns());
        mcsp.setRepeatingRows(sheet.getRepeatingRows());

        mcsp.setCopies(ps.getCopies());
        mcsp.setDraft(ps.getDraft());
        mcsp.setFitHeight(ps.getFitHeight());
        mcsp.setFitWidth(ps.getFitWidth());
        mcsp.setHResolution(ps.getHResolution());
        mcsp.setLandscape(ps.getLandscape());
        mcsp.setNoColor(ps.getNoColor());
        mcsp.setLeftToRight(ps.getLeftToRight());
        mcsp.setNotes(ps.getNotes());
        mcsp.setPageStart(ps.getPageStart());
        mcsp.setPaperSize(ps.getPaperSize());
        mcsp.setScale(ps.getScale());
        mcsp.setUsePage(ps.getUsePage());
        mcsp.setValidSettings(ps.getValidSettings());
        mcsp.setVResolution(ps.getVResolution());

        return mcsp;
    }

    /**
     * Clones all named ranges that are scoped to the <code>Sheet</code> at the
     * given index, and scopes the newly cloned named ranges to the last sheet
     * in the workbook, where it is assumed that the cloned sheet still exists.
     * This was moved to <code>SheetCloner</code> for version 0.9.1.
     * @param workbook A <code>Workbook</code>.
     * @param prevIndex The 0-based sheet index from which to clone named
     *    ranges.
     * @since 0.8.0
     */
    private void cloneNamedRanges(Workbook workbook, int prevIndex)
    {
        int numNamedRanges = workbook.getNumberOfNames();
        int clonedSheetIndex = workbook.getNumberOfSheets() - 1;
        for (int i = 0; i < numNamedRanges; i++)
        {
            Name name = workbook.getNameAt(i);
            // Avoid copying Excel's "built-in" (and hidden) named ranges.
            if (name.getSheetIndex() == prevIndex && !NameTag.EXCEL_BUILT_IN_NAMES.contains(name.getNameName()))
            {
                Name clone = workbook.createName();
                // This will be a sheet-scoped clone of a name that could be workbook-scoped.
                clone.setSheetIndex(clonedSheetIndex);
                clone.setNameName(name.getNameName());
                clone.setComment(name.getComment());
                clone.setFunction(name.isFunctionName());
                clone.setRefersToFormula(name.getRefersToFormula());
            }
        }
    }

    /**
     * The sheet order has changed; a <code>Sheet</code> has been moved from one
     * position to another.  Apache POI doesn't change the scopes of named
     * ranges to match this change.  This accomplishes the task here.  This was
     * moved to <code>SheetCloner</code> for version 0.9.1.
     * @param workbook The <code>Workbook</code> on which a sheet was moved.
     * @param fromIndex The 0-based previous index of the <code>Sheet</code>
     *    that was moved.
     * @param toIndex The 0-based current index of the <code>Sheet</code> that
     *    was moved.
     * @since 0.8.0
     */
    private void updateNamedRangesScope(Workbook workbook, int fromIndex, int toIndex)
    {
        if (fromIndex != toIndex)
        {
            int numNamedRanges = workbook.getNumberOfNames();
            for (int i = 0; i < numNamedRanges; i++)
            {
                Name name = workbook.getNameAt(i);
                int scopeIndex = name.getSheetIndex();
                if (scopeIndex == fromIndex)
                {
                    name.setSheetIndex(toIndex);
                }
                else if (fromIndex < scopeIndex && scopeIndex < toIndex)
                {
                    name.setSheetIndex(scopeIndex - 1);
                }
                else if (toIndex < scopeIndex && scopeIndex < fromIndex)
                {
                    name.setSheetIndex(scopeIndex + 1);
                }
            }
        }
    }

    /**
     * Performs all manipulation necessary for implicit cloning, so that the
     * <code>SheetTransformer</code> can transform the resultant sheets as if
     * they were already there.
     * @param sheet The <code>Sheet</code> on which to perform implicit cloning.
     * @param beans The beans map.
     * @param context The <code>WorkbookContext</code>.
     * @return A beans <code>Map</code> to use for transformation, which may
     *    be <code>beans</code>.
     */
    @SuppressWarnings("unchecked")
    public Map<String, Object> setupForImplicitCloning(Sheet sheet, Map<String, Object> beans, WorkbookContext context)
    {
        // The SheetTransformer that calls this method may need to replace its
        // beans map with what is created here.
        Map<String, Object> useThisBeansMap = beans;

        // 1. Use a MetadataParser to extract any metadata.
        MetadataParser parser = extractMetadata(sheet, context);

        // 2. Extract parser properties and validate.
        String replacementValue = "";  // default
        String indexVarName = null;
        int limit = -1;
        String varStatusName = null;
        if (parser != null)
        {
            replacementValue = parser.getReplacementValue();
            indexVarName = parser.getIndexVarName();
            try
            {
                limit = Integer.parseInt(parser.getLimit());
            }
            catch (NumberFormatException e)
            {
                throw new MetadataParseException("Limit must be a number: " + parser.getLimit(), e);
            }
            varStatusName = parser.getVarStatusName();
        }

        // 3. Find all collection names in the current sheet name.
        List<String> collExprs = findCollectionsInSheetName(sheet, beans, context);
        List<Collection<Object>> collections = new ArrayList<>();
        logger.debug("collExprs: {}", collExprs);
        for (String collExpression : collExprs)
        {
            Object result = Expression.evaluateString(
                    Expression.BEGIN_EXPR + collExpression.trim() + Expression.END_EXPR,
                    context.getExpressionFactory(), beans);

            if (result == null)
            {
                // Allow null to be interpreted as an empty collection.
                result = new ArrayList<>(0);
            }
            if (!(result instanceof Collection))
            {
                throw new MetadataParseException("One of the items in the sheet name is not a Collection: \"" + collExpression);
            }
            collections.add((Collection<Object>) result);
        }
        int maxSize = 0;
        for (Collection<Object> coll : collections)
        {
            if (coll.size() > maxSize)
                maxSize = coll.size();
        }
        if (limit == -1)
        {
            limit = maxSize;
        }

        // 4. Clone the sheets as necessary.
        boolean isSheetSpecificBeans = !context.getTemplateSheetNames().isEmpty();
        String sheetName = sheet.getSheetName();
        int index = myWorkbook.getSheetIndex(sheetName);
        CreationHelper helper = myWorkbook.getCreationHelper();
        List<Map<String, Object>> beansMaps = context.getBeansMaps();

        // Ensure MissingCloneSheetProperties exists for this sheet, in the
        // proper position.
        if (myMissingPropertiesList.size() <= index)
        {
            if (myMissingPropertiesList.size() < index)
            {
                myMissingPropertiesList.addAll(Collections.<MissingCloneSheetProperties>nCopies(index, null));
            }
            myMissingPropertiesList.add(getMissingCloneSheetProperties(myWorkbook.getSheetAt(index)));
        }

        if (limit >= 1)
        {
            // Find the sheet.
            if (index == -1)
                throw new RuntimeException("Implicit cloning Sheet \"" + sheetName + "\" not found!");

            // Clone the sheet.
            String origSheetName = null;
            for (int i = 0; i < limit; i++)
            {
                MissingCloneSheetProperties cloned = null;
                if (i > 0)
                {
                    // Clone the sheet, and move it into position.
                    logger.debug("Implicitly cloning sheet at position {}.", index);

                    cloned = new MissingCloneSheetProperties(myMissingPropertiesList.get(index));
                    myWorkbook.cloneSheet(index);
                }
                // New name for the sheet.
                RichTextString temp = helper.createRichTextString(sheetName);
                for (String collExpression : collExprs)
                {
                    String replacement = collExpression + "." + i;
                    temp = RichTextStringUtil.replaceAll(temp, helper, collExpression, replacement, false, 0, true);
                }
                String newSheetName = temp.getString();

                if (i == 0)
                {
                    logger.debug("Setting sheet name at position {} to \"{}\".",
                            index, newSheetName);
                    newSheetName = SheetUtil.safeSetSheetName(myWorkbook, index, newSheetName);
                    FormulaUtil.replaceSheetNameRefs(context, sheetName, newSheetName);
                    origSheetName = newSheetName;
                }
                else
                {
                    logger.debug("Setting new sheet name at position {} to \"{}\".",
                            myWorkbook.getNumberOfSheets() - 1, newSheetName);
                    int clonePos = myWorkbook.getNumberOfSheets() - 1;
                    newSheetName = SheetUtil.safeSetSheetName(myWorkbook, clonePos, newSheetName);
                    cloneNamedRanges(myWorkbook, index);

                    logger.debug("Moving sheet \"{}\" to position {}.",
                            newSheetName, index + i);

                    myWorkbook.setSheetOrder(newSheetName, index + i);
                    updateNamedRangesScope(myWorkbook, clonePos, index + i);

                    myMissingPropertiesList.add(cloned);
                    FormulaUtil.addSheetNameRefsAfterClone(context, origSheetName, newSheetName, index + i);
                }
            }

            // Set up sheets and beans.
            if (isSheetSpecificBeans)
            {
                List<Iterator<Object>> iterators = new ArrayList<>();
                Map<String, Object> source = beansMaps.get(index);
                for (Collection<Object> collection : collections)
                {
                    iterators.add(collection.iterator());
                }
                for (int i = 0; i < limit; i++)
                {
                    Sheet cloned = myWorkbook.getSheetAt(index + i);
                    // Expose indexVar/varStatus.
                    List<String> varNames = CollectionsTransformer.getImplicitVarNames(collExprs);
                    Map<String, Object> wrappingMap = new HashMapWrapper<>(source);
                    for (int c = 0; c < collExprs.size(); c++)
                    {
                        Iterator<Object> itr = iterators.get(c);
                        Object value;
                        if (itr.hasNext())
                        {
                            value = itr.next();
                        }
                        else
                        {
                            value = PastEndValue.PAST_END_VALUE;
                        }
                        String varName = varNames.get(c);
                        wrappingMap.put(varName, value);
                    }
                    if (indexVarName != null && !indexVarName.isEmpty())
                    {
                        wrappingMap.put(indexVarName, i);
                    }
                    if (varStatusName != null && !varStatusName.isEmpty())
                    {
                        wrappingMap.put(varStatusName, new BaseLoopTagStatus(i, limit));
                    }
                    if (i == 0)
                    {
                        beansMaps.set(index, wrappingMap);
                        useThisBeansMap = wrappingMap;
                    }
                    else
                    {
                        beansMaps.add(index + i, wrappingMap);
                    }
                    SheetUtil.setUpSheetForImplicitCloningAccess(cloned, collExprs, varNames);

                    List<String> pastEndRefs = new ArrayList<>();
                    for (int c = 0; c < collections.size(); c++)
                    {
                        if (i >= collections.get(c).size())
                        {
                            pastEndRefs.add(varNames.get(c));
                            // This covers the sheet name too.
                            pastEndRefs.add(collExprs.get(c) + "." + i);
                        }
                    }
                    sheetName = cloned.getSheetName();
                    String newSheetName = SheetUtil.takePastEndAction(cloned, pastEndRefs, replacementValue);
                    if (!sheetName.equals(newSheetName))
                    {
                        FormulaUtil.replaceSheetNameRefs(context, sheetName, newSheetName);
                    }
                }
            }
            else
            {
                // One beans map to rule them all.
                // Expose indexVar/varStatus.
                if (indexVarName != null && !indexVarName.isEmpty())
                {
                    List<Integer> indexVars = new ArrayList<>();
                    for (int i = 0; i < limit; i++)
                    {
                        indexVars.add(i);
                    }
                    beans.put(indexVarName, indexVars);
                }
                if (varStatusName != null && !varStatusName.isEmpty())
                {
                    List<BaseLoopTagStatus> varStatuses = new ArrayList<>();
                    for (int i = 0; i < limit; i++)
                    {
                        varStatuses.add(new BaseLoopTagStatus(i, limit));
                    }
                    beans.put(varStatusName, varStatuses);
                }
                for (int i = 0; i < limit; i++)
                {
                    Sheet cloned = myWorkbook.getSheetAt(index + i);

                    // Set up the sheet for implicit cloning access.
                    List<String> localCollExprs = new ArrayList<>(collExprs);
                    List<String> replacementExprs = new ArrayList<>(localCollExprs.size());
                    for (String collExpr : localCollExprs)
                    {
                        replacementExprs.add(collExpr + "." + i);
                    }

                    if (indexVarName != null && !indexVarName.isEmpty())
                    {
                        localCollExprs.add(indexVarName);
                        replacementExprs.add(indexVarName + "." + i);
                    }
                    if (varStatusName != null && !varStatusName.isEmpty())
                    {
                        localCollExprs.add(varStatusName);
                        replacementExprs.add(varStatusName + "." + i);
                    }
                    SheetUtil.setUpSheetForImplicitCloningAccess(cloned, localCollExprs, replacementExprs);
                    List<String> pastEndRefs = new ArrayList<>();
                    for (int c = 0; c < collections.size(); c++)
                    {
                        if (i >= collections.get(c).size())
                        {
                            pastEndRefs.add(replacementExprs.get(c));
                        }
                    }
                    sheetName = cloned.getSheetName();
                    String newSheetName = SheetUtil.takePastEndAction(cloned, pastEndRefs, replacementValue);
                    if (!sheetName.equalsIgnoreCase(newSheetName))
                    {
                        FormulaUtil.replaceSheetNameRefs(context, sheetName, newSheetName);
                    }
                }  // end for loop to limit
            }  // end case not sheet specific beans
        }
        else if (limit == 0)
        {
            // Blank out the entire sheet?
            // Sheet name uses replaceValue?
            Sheet toBlank = myWorkbook.getSheetAt(index);
            String oldSheetName = toBlank.getSheetName();
            List<String> augmented = new ArrayList<>(collExprs);
            if (indexVarName != null && !indexVarName.isEmpty())
            {
                augmented.add(indexVarName);
            }
            if (varStatusName != null && !varStatusName.isEmpty())
            {
                augmented.add(varStatusName);
            }
            String newSheetName = SheetUtil.takePastEndAction(toBlank, augmented, replacementValue);
            if (!oldSheetName.equals(newSheetName))
            {
                FormulaUtil.replaceSheetNameRefs(context, sheetName, newSheetName);
            }
        }
        return useThisBeansMap;
    }

    /**
     * Returns a <code>MetadataParser</code> that holds the metadata
     * information from the sheet name, if present.
     * @param sheet The <code>Sheet</code>.
     * @param context The <code>WorkbookContext</code>.
     * @return A <code>MetadataParser</code>, or <code>null</code> if no
     *    metadata was present.
     */
    private MetadataParser extractMetadata(Sheet sheet, WorkbookContext context)
    {
        MetadataParser parser = null;
        String name = sheet.getSheetName();
        int metadataIndIdx = name.indexOf(BEGIN_METADATA);
        if (metadataIndIdx != -1)
        {
            // Evaluate any Expressions in the metadata.
            String metadata = name.substring(metadataIndIdx + BEGIN_METADATA.length());
            logger.debug("  SC: Metadata found: {} on sheet {}",
                    metadata, sheet.getSheetName());
            // Parse the Metadata.
            parser = new SheetNameMetadataParser(metadata);
            parser.setCell(null);
            parser.parse();
            // Remove the metadata text from the sheet name.
            String metadataRemoved = name.replaceAll(Pattern.quote(BEGIN_METADATA + metadata), "");
            Workbook workbook = sheet.getWorkbook();
            metadataRemoved = SheetUtil.safeSetSheetName(workbook, workbook.getSheetIndex(sheet), metadataRemoved);
            FormulaUtil.replaceSheetNameRefs(context, name, metadataRemoved);
        }

        return parser;
    }

    /**
     * Finds all <code>Collection</code> names in the name of the source sheet.
     * starting with the given <code>Cell</code>.
     * @param sheet The <code>Sheet</code>.
     * @param beans The beans map for the <code>Sheet</code>.
     * @param context The <code>WorkbookContext</code>.
     * @return A <code>List</code> of all <code>Collection</code> names found.
     */
    private List<String> findCollectionsInSheetName(Sheet sheet, Map<String, Object> beans, WorkbookContext context)
    {
        String metadataRemoved = sheet.getSheetName();
        return Expression.getImplicitCollectionExpr(metadataRemoved, beans, context);
    }
}
