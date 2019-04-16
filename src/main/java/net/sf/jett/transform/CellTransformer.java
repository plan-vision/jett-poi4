package net.sf.jett.transform;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Stack;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import net.sf.jett.event.CellEvent;
import net.sf.jett.event.CellListener;
import net.sf.jett.exception.ParseException;
import net.sf.jett.exception.TagParseException;
import net.sf.jett.exception.TransformException;
import net.sf.jett.expression.Expression;
import net.sf.jett.model.Block;
import net.sf.jett.model.WorkbookContext;
import net.sf.jett.parser.TagParser;
import net.sf.jett.tag.Tag;
import net.sf.jett.tag.TagContext;
import net.sf.jett.tag.TagLibraryRegistry;
import net.sf.jett.util.RichTextStringUtil;
import net.sf.jett.util.SheetUtil;

/**
 * A <code>CellTransformer</code> knows how to transform a <code>Cell</code>
 * inside of a <code>Sheet</code>.  If a <code>Tag</code> is found, then the
 * <code>CellTransformer</code> will process it.
 *
 * @author Randy Gettman
 */
public class CellTransformer
{
    private static final Logger logger = LoggerFactory.getLogger(CellTransformer.class);
    private static final Logger tagLogger = LoggerFactory.getLogger(CellTransformer.class.getName() + ".Tag");

    /**
     * Transforms the given <code>Cell</code>, using the given <code>Map</code>
     * of bean names to bean objects.
     * @param cell The <code>Cell</code> to transform.
     * @param workbookContext The <code>WorkbookContext</code> that provides the
     *    <code>Map</code> of <code>Formulas</code>, the
     *    <code>TagLibraryRegistry</code>, the <code>CellListeners</code>, the
     *    fixed size collection names, and the turned off implicit collection
     *    names.
     * @param cellContext The <code>TagContext</code> that provides the
     *    <code>Map</code> of beans data,  the <code>Map</code> of processed
     *    <code>Cells</code>, and the parent <code>Block</code>.
     * @return <code>true</code> if this <code>Cell</code> was transformed,
     *    <code>false</code> if it needs to be transformed again.  This may
     *    happen if the <code>Block</code> associated with the <code>Tag</code>
     *    was removed.
     */
    public boolean transform(Cell cell, WorkbookContext workbookContext, TagContext cellContext)
    {
        Map<String, Object> beans = cellContext.getBeans();
        Map<String, Cell> processedCells = cellContext.getProcessedCellsMap();

        // Make sure this Cell hasn't already been processed.
        String key = SheetUtil.getCellKey(cell);
        if (processedCells.containsKey(key))
            return true;

        exposeCell(beans, cell);

        Object oldValue = null;
        switch(cell.getCellType())
        {
	        case STRING:
	            oldValue = cell.getStringCellValue();
	            break;
	        case NUMERIC:
	            if (DateUtil.isCellDateFormatted(cell))
	                oldValue = cell.getDateCellValue();  // java.util.Date
	            else
	                oldValue = cell.getNumericCellValue();  // double
	            break;
	        case BLANK:
	            oldValue = null;
	            break;
	        case FORMULA:
	            oldValue = cell.getCellFormula();  // java.lang.String
	            break;
	        case BOOLEAN:
	            oldValue = cell.getBooleanCellValue();  // boolean
	            break;
	        case ERROR:
	            oldValue = cell.getErrorCellValue();  // byte
	            break;
        }

        if (!fireBeforeCellProcessedEvent(workbookContext, cell, beans, oldValue))
        {
            // Mark as processed without actually processing it.
            processedCells.put(key, cell);
            return true;
        }

        logger.debug("Processing row={}, col={} on sheet {}",
                cell.getRowIndex(), cell.getColumnIndex(), cell.getSheet().getSheetName());
        logger.debug("Parent Block: {}", cellContext.getBlock());

        Sheet sheet = cell.getSheet();
        boolean cellProcessed = true;
        Object newValue = null;
        switch(cell.getCellType())
        {
        case STRING:
            TagParser parser = new TagParser(cell);
            parser.parse();

            if (parser.isTag() && !parser.isEndTag())
            {
                // Transform the Tag.
                logger.trace("Transforming tag cell tag.");
                cellProcessed = transformCellTag(cell, workbookContext, cellContext, parser);
            }
            else
            {
                // Not a tag.  Evaluate any Expressions embedded in the value.
                RichTextString richString = cell.getRichStringCellValue();
                List<String> collExprs = Expression.getImplicitCollectionExpr(richString.toString(),
                        beans, workbookContext);
                if (!collExprs.isEmpty())
                {
                    logger.trace("  Transforming implicit collection(s).");
                    CollectionsTransformer collTransformer = new CollectionsTransformer();
                    collTransformer.transform(cell, workbookContext, cellContext);
                    // The implicit collection processing has already processed this Cell.
                    cellProcessed = false;
                }
                else
                {
                    // Evaluate.
                    logger.trace("  Transforming string cell.");
                    CreationHelper helper = sheet.getWorkbook().getCreationHelper();
                    Object result = Expression.evaluateString(richString, helper, workbookContext.getExpressionFactory(), beans);
                    newValue = SheetUtil.setCellValue(workbookContext, cell, result, richString);
                }
            }
            break;
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell))
                newValue = cell.getDateCellValue();  // java.util.Date
            else
                newValue = cell.getNumericCellValue();  // double
            break;
        case BLANK:
            newValue = null;
            break;
        case FORMULA:
            newValue = cell.getCellFormula();  // java.lang.String
            break;
        case BOOLEAN:
            newValue = cell.getBooleanCellValue();  // boolean
            break;
        case ERROR:
            newValue = cell.getErrorCellValue();  // byte
            break;
        }  // End switch on cell type
        if (cellProcessed)
        {
            fireCellProcessedEvent(workbookContext, cell, beans, oldValue, newValue);
            // Only mark it as processed if the Cell has actually been processed.
            processedCells.put(key, cell);
        }
        return cellProcessed;
    }

    /**
     * Calls all <code>CellListeners'</code> <code>beforeCellProcessed</code>
     * method, sending a <code>CellEvent</code>.  The new cell value is not
     * available because the cell hasn't been processed yet.
     * @param context The <code>WorkbookContext</code> object.
     * @param cell The <code>Cell</code> that is about to be processed.
     * @param beans A <code>Map</code> of bean names to bean values
     * @param oldValue The old cell value.
     * @return Whether processing of the <code>Sheet</code> should occur.  If
     *    any <code>SheetListener's</code> <code>beforeSheetProcessed</code>
     *    method returns <code>false</code>, then this method returns
     *    <code>false</code>.
     * @since 0.8.0
     */
    private boolean fireBeforeCellProcessedEvent(WorkbookContext context, Cell cell, Map<String, Object> beans,
                                                 Object oldValue)
    {
        boolean shouldProceed = true;
        List<CellListener> cellListeners = context.getCellListeners();
        CellEvent event = new CellEvent(cell, beans, oldValue, null);
        for (CellListener listener : cellListeners)
        {
            shouldProceed &= listener.beforeCellProcessed(event);
        }
        return shouldProceed;
    }

    /**
     * Calls all <code>CellListeners'</code> <code>cellProcessed</code>
     * method, sending a <code>CellEvent</code>.  This functionality was
     * extracted out of the <code>transform</code> method.
     * @param context The <code>WorkbookContext</code> object.
     * @param cell The <code>Cell</code> that was processed.
     * @param beans A <code>Map</code> of bean names to bean values
     * @param oldValue The old cell value.
     * @param newValue The new cell value.
     * @since 0.8.0
     */
    private void fireCellProcessedEvent(WorkbookContext context, Cell cell, Map<String, Object> beans,
                                        Object oldValue, Object newValue)
    {
        List<CellListener> cellListeners = context.getCellListeners();
        CellEvent event = new CellEvent(cell, beans, oldValue, newValue);
        for (CellListener listener : cellListeners)
        {
            listener.cellProcessed(event);
        }
    }

    /**
     * Transforms the <code>Tag</code> defined in the given <code>Cell</code>.
     * @param cell The <code>Cell</code> on which the <code>Tag</code> is defined.
     * @param workbookContext The <code>WorkbookContext</code>.
     * @param cellContext The <code>CellContext</code>.
     * @param parser The <code>TagParser</code> used to parse the tag's text.
     * @return <code>true</code> if this <code>Cell</code> was transformed,
     *    <code>false</code> if it needs to be transformed again.  This may
     *    happen if the <code>Block</code> associated with the <code>Tag</code>
     *    was removed.
     */
    private boolean transformCellTag(Cell cell, WorkbookContext workbookContext,
                                     TagContext cellContext, TagParser parser)
    {
        Block parentBlock = cellContext.getBlock();
        TagLibraryRegistry registry = workbookContext.getRegistry();
        Map<String, Object> beans = cellContext.getBeans();
        Sheet sheet = cellContext.getSheet();
        Map<String, Cell> processedCells = cellContext.getProcessedCellsMap();
        String value = cell.getStringCellValue();
        RichTextString richTextString = cell.getRichStringCellValue();
        Block newBlock;
        Tag tag = null;
        try
        {
            if (parser.isBodiless())
            {
                // Results in a 1x1 block of 1 cell.
                newBlock = new Block(parentBlock, cell);
            }
            else
            {
                // Remove start tag text.
                SheetUtil.setCellValue(workbookContext, cell, RichTextStringUtil.replaceAll(richTextString,
                        sheet.getWorkbook().getCreationHelper(), parser.getTagText(), "", true));
                tagLogger.debug("Cell text after tag removal is \"{}\".", cell.getStringCellValue());
                // Search for matching end tag.  If found, remove the end tag.
                Cell match = findMatchingEndTag(workbookContext, cell, parentBlock, parser.getNamespaceAndTagName());
                if (match == null)
                    throw new TagParseException("Matching tag not found for tag: " + parser.getTagText() +
                            ", located" + SheetUtil.getCellLocation(cell) + ", within block " + parentBlock);

                tagLogger.debug("  Match found at row {} and column {}",
                        match.getRowIndex(), match.getColumnIndex());
                newBlock = new Block(parentBlock, cell, match);
            }
            TagContext context = new TagContext();
            context.setBeans(beans);
            context.setBlock(newBlock);
            context.setSheet(sheet);
            context.setProcessedCellsMap(processedCells);
            context.setDrawing(cellContext.getDrawing());
            context.setMergedRegions(cellContext.getMergedRegions());
            context.setFormulaSuffix(cellContext.getFormulaSuffix());

            tag = registry.createTag(parser, context, workbookContext);
            if (tag == null)
            {
                Map<String, String> tagLocationsMap = workbookContext.getTagLocationsMap();
                String cellRef = SheetUtil.getCellKey(cell);
                String location = " at " + cellRef;
                String origCellRef = tagLocationsMap.get(cellRef);
                if (origCellRef != null)
                {
                    location += " (originally located at " + origCellRef + ")";
                }
                throw new TagParseException("Invalid tag: " + value + location +
                        SheetUtil.getTagLocationWithHierarchy(cellContext.getCurrentTag()));
            }
            else
            {
                context.setCurrentTag(tag);
                tag.setParentTag(cellContext.getCurrentTag());
            }

            // Process the Tag.
            return tag.processTag();
        }
        catch (ParseException | TransformException e)
        {
            tagLogger.error("An error has occurred: " + e.getMessage(), e);
            // Don't re-wrap.
            throw e;
        }
        catch (RuntimeException re)
        {
            tagLogger.error("An error has occurred: " + re.getMessage(), re);
            throw new TransformException("A " + re.getClass().getName() +
                    " was caught during transformation" +
                    ((tag != null) ?
                            SheetUtil.getTagLocationWithHierarchy(tag) :
                            SheetUtil.getCellLocation(cell)),
                    re);
        }
    }

    /**
     * Finds the end tag that matches the given start tag.  The end tag must
     * reside inside the given <code>parentBlock</code>.
     * @param context The <code>WorkbookContext</code>.
     * @param startTag The <code>Cell</code> with the start tag.
     * @param parentBlock The parent <code>Block</code> in which the given
     *    <code>Cell</code> is contained.  The end tag must also be contained
     *    within this <code>Block</code>.
     * @param namespaceAndTagName The namespace and tag name of the start tag,
     *    e.g. "namespace:tagName".
     * @return The <code>Cell</code> containing the matching end tag, or
     *    <code>null</code> if there is no matching end tag.
     */
    private Cell findMatchingEndTag(WorkbookContext context, Cell startTag, Block parentBlock, String namespaceAndTagName)
    {
        int startColumnIndex = startTag.getColumnIndex();
        int startRowIndex = startTag.getRowIndex();
        int right = parentBlock.getRightColNum();
        int bottom = parentBlock.getBottomRowNum();

        tagLogger.debug("fMET: Matching tag {} in {}, starting tag found at row {}, cell {}",
                namespaceAndTagName, parentBlock, startRowIndex, startColumnIndex);

        List<TagParser> innerTags = new ArrayList<>();

        // Look for candidate matches in current Cell, to its right, below it, or
        // both.
        Sheet sheet = startTag.getSheet();
        for (int rowNum = startRowIndex; rowNum <= bottom; rowNum++)
        {
            Row row = sheet.getRow(rowNum);
            if (row != null)
            {
                for (int cellNum = startColumnIndex; cellNum <= right; cellNum++)
                {
                    tagLogger.trace("  Trying cell: row {}, col {}", rowNum, cellNum);
                    Cell candidate = row.getCell(cellNum);
                    if (candidate != null && isMatchingEndTag(context, candidate, namespaceAndTagName, innerTags))
                        return candidate;
                }
            }
        }
        // If we got here, then there wasn't a match.
        return null;
    }

    /**
     * Helper method to determine if the given candidate <code>Cell</code> is an
     * end tag that matches the given namespace and tag name, considering the
     * given <code>List</code> of unmatched inner tags already encountered.
     *
     * @param context The <code>WorkbookContext</code>.
     * @param candidate The candidate <code>Cell</code>.
     * @param namespaceAndTagName The namespace and tag name to match.
     * @param innerTags A <code>List</code> of inner tags which must be matched
     *    prior to matching the given namespace and tag name.  This stack may be
     *    modified if <code>candidate</code> is itself a start tag, or if
     *    <code>candidate</code> is an end tag that matches an inner tag.
     * @return <code>true</code> if it matches, <code>false</code> otherwise.
     */
    private boolean isMatchingEndTag(WorkbookContext context, Cell candidate, String namespaceAndTagName,
                                     List<TagParser> innerTags)
    {
        if (candidate.getCellType() != CellType.STRING)
            return false;
        TagParser candidateParser = new TagParser(candidate);
        candidateParser.parse();
        int rightMostCol = candidate.getColumnIndex();
        int afterTagIdx = 0;
        tagLogger.debug("    iMET: afterTagIdx={}, parser's tag text is \"{}\".",
                afterTagIdx, candidateParser.getTagText());

        // Look for possibly multiple tags on the same Cell.
        while (candidateParser.isTag())
        {
            if (candidateParser.isEndTag())
            {
                // Found matching end tag with no unclosed intervening start tags.
                if (namespaceAndTagName.equals(candidateParser.getNamespaceAndTagName()) &&
                        doAllInnerTagsMatch(innerTags, rightMostCol))
                {
                    // This is the matching end tag.  Remove it from the Cell.
                    SheetUtil.setCellValue(context, candidate, RichTextStringUtil.replaceAll(candidate.getRichStringCellValue(),
                            candidate.getSheet().getWorkbook().getCreationHelper(), candidateParser.getTagText(), "", true, afterTagIdx));
                    return true;
                }
                else
                {
                    // End tag matches an intervening start tag.
                    if (innerTags.isEmpty())
                    {
                        throw new TagParseException("End tag found \"" + candidateParser.getNamespaceAndTagName() +
                                "\" does not match start tag \"" + namespaceAndTagName + "\"" + SheetUtil.getCellLocation(candidate) + ".");
                    }
                    tagLogger.trace("    Adding end tag to list: {}", candidateParser.getNamespaceAndTagName());
                    innerTags.add(candidateParser);
                }
            }
            else if (!candidateParser.isEndTag())
            {
                // Found another start tag.  If bodiless, don't bother pushing it.
                // If it is not bodiless, then it now needs to be matched BEFORE we
                // can match the original start tag.  Push it onto the "stack".
                if (!candidateParser.isBodiless())
                {
                    tagLogger.trace("    Adding start tag to list: {}", candidateParser.getNamespaceAndTagName());
                    innerTags.add(candidateParser);
                }
            }
            // Setup for next loop.  Advance past this tag.
            afterTagIdx += candidateParser.getAfterTagIdx();
            candidateParser = new TagParser(candidate, afterTagIdx);
            candidateParser.parse();
            tagLogger.trace("    afterTagIdx is now {}, parser's tag text is \"{}\".",
                    afterTagIdx, candidateParser.getTagText());
        }
        // If we got here, then we did not match.
        return false;
    }

    /**
     * Determines whether all tags in the given <code>List</code>, disregarding
     * any tags found to the right of the given column index, i.e.
     * <code>parser.getCell().getColumnIndex() &gt; rightMostCol</code>.
     *
     * @param innerTags The <code>List</code> of <code>TagParsers</code>
     *    containing tags to match.
     * @param rightMostCol Disregard all tags found to the right of this column
     *    index (0-based).  Pass -1 to consider all tags, no matter how far to
     *    the right they are.
     * @return <code>true</code> if all considered tags match,
     *    <code>false</code> otherwise.
     */
    private boolean doAllInnerTagsMatch(List<TagParser> innerTags, int rightMostCol)
    {
        Stack<TagParser> tagsToMatch = new Stack<>();
        tagLogger.trace("    dAITM:");
        for (TagParser parser : innerTags)
        {
            Cell candidateCell = parser.getCell();
            if (candidateCell.getColumnIndex() <= rightMostCol)
            {
                tagLogger.trace("      dAITM: Considering tag: {} at row {}, col {}",
                        parser.getNamespaceAndTagName(), parser.getCell().getRowIndex(), parser.getCell().getColumnIndex());
                if (parser.isEndTag())
                {
                    // Unmatched end tag.
                    if (tagsToMatch.isEmpty())
                    {
                        tagLogger.trace("      dAITM: Unmatched end tag.");
                        return false;
                    }

                    String namespaceAndTagName = parser.getNamespaceAndTagName();
                    TagParser startParser = tagsToMatch.peek();
                    tagLogger.trace("      dAITM: Comparing start: {} to end: {}",
                            startParser.getNamespaceAndTagName(), namespaceAndTagName);
                    if (namespaceAndTagName.equals(startParser.getNamespaceAndTagName()))
                    {
                        tagLogger.trace("      dAITM: Popped {}", startParser.getNamespaceAndTagName());
                        tagsToMatch.pop();
                    }
                }
                else if (!parser.isBodiless())
                {
                    tagLogger.trace("      dAITM: Pushed {}", parser.getNamespaceAndTagName());
                    tagsToMatch.push(parser);
                }
            }
        }
        return tagsToMatch.isEmpty();
    }

    /**
     * Make the <code>Cell</code> object available as bean in the given
     * <code>Map</code> of beans.
     * @param beans The <code>Map</code> of beans.
     * @param cell The <code>Cell</code> to expose.
     * @since 0.4.0
     */
    private void exposeCell(Map<String, Object> beans, Cell cell)
    {
        beans.put("cell", cell);
    }
}
