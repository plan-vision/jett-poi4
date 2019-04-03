package net.sf.jett.util;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.logging.log4j.Logger;
import org.apache.logging.log4j.LogManager;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import net.sf.jett.expression.Expression;
import net.sf.jett.model.CellStyleCache;
import net.sf.jett.model.FontCache;

/**
 * The <code>RichTextStringUtil</code> utility class provides methods for
 * RichTextString manipulation.
 *
 * @author Randy Gettman
 */
public class RichTextStringUtil
{
    private static final Logger logger = LogManager.getLogger();

    /**
     * Replaces all occurrences of the given target string with the replacement
     * string.
     * Preserves rich text formatting as much as possible.
     * @param richTextString The <code>RichTextString</code> to manipulate.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param target The string to replace.
     * @param replacement The replacement string.
     * @return A new <code>RichTextString</code> with replaced values, or the
     *    same <code>RichTextString</code> if <code>replace</code> is
     *    <code>null</code> or empty.
     */
    public static RichTextString replaceAll(RichTextString richTextString,
                                            CreationHelper helper, String target, String replacement)
    {
        return replaceAll(richTextString, helper, target, replacement, false);
    }

    /**
     * Replaces all occurrences of the given target string with the replacement
     * string.
     * Preserves rich text formatting as much as possible.
     * @param richTextString The <code>RichTextString</code> to manipulate.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param target The string to replace.
     * @param replacement The replacement string.
     * @param firstOnly Whether to stop after replacing the first found instance
     *    of <code>target</code>.
     * @return A new <code>RichTextString</code> with replaced values, or the
     *    same <code>RichTextString</code> if <code>replace</code> is
     *    <code>null</code> or empty.
     */
    public static RichTextString replaceAll(RichTextString richTextString,
                                            CreationHelper helper, String target, String replacement, boolean firstOnly)
    {
        return replaceAll(richTextString, helper, target, replacement, firstOnly, 0);
    }

    /**
     * Replaces all occurrences of the given target string with the replacement
     * string.
     * Preserves rich text formatting as much as possible.
     * @param richTextString The <code>RichTextString</code> to manipulate.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param target The string to replace.
     * @param replacement The replacement string.
     * @param firstOnly Whether to stop after replacing the first found instance
     *    of <code>target</code>.
     * @param startIdx Start replacing after this 0-based index into the
     *    <code>richTextString</code>.
     * @return A new <code>RichTextString</code> with replaced values, or the
     *    same <code>RichTextString</code> if <code>replace</code> is
     *    <code>null</code> or empty.
     */
    public static RichTextString replaceAll(RichTextString richTextString,
                                            CreationHelper helper, String target, String replacement, boolean firstOnly, int startIdx)
    {
        return replaceAll(richTextString, helper, target, replacement, firstOnly, startIdx, false);
    }

    /**
     * Replaces all occurrences of the given target string with the replacement
     * string.
     * Preserves rich text formatting as much as possible.
     * @param richTextString The <code>RichTextString</code> to manipulate.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param target The string to replace.
     * @param replacement The replacement string.
     * @param firstOnly Whether to stop after replacing the first found instance
     *    of <code>target</code>.
     * @param startIdx Start replacing after this 0-based index into the
     *    <code>richTextString</code>.
     * @param identifierMode If true, makes sure that the <code>target</code> is
     *    NOT replaced if the target string is part of a larger identifier.
     *    E.g. if <code>target</code> is <code>"activity"</code>, don't replace
     *    the <code>"activity"</code> substring within <code>"activityDay"</code>.
     *    Only replace the target if found within expressions.
     * @return A new <code>RichTextString</code> with replaced values, or the
     *    same <code>RichTextString</code> if <code>replace</code> is
     *    <code>null</code> or empty.
     *
     * @since 0.5.2
     */
    public static RichTextString replaceAll(RichTextString richTextString,
                                            CreationHelper helper, String target, String replacement, boolean firstOnly, int startIdx,
                                            boolean identifierMode)
    {
        if (target == null || target.length() == 0)
            return richTextString;

        int numFormattingRuns = richTextString.numFormattingRuns();
        String value = richTextString.getString();
        logger.trace("replaceAll: \"{}\" ({}): numFormattingRuns={}, replacing \"{}\" with \"{}\".",
                value, value.length(), numFormattingRuns, target, replacement);

        List<FormattingRun> formattingRuns = determineFormattingRunStats(richTextString);

        // Replace target(s) with the replacement.
        logger.debug("  Replacing \"{}\" with \"{}\".", target, replacement);
        int change = replacement.length() - target.length();
        int beginIdx = value.indexOf(target, startIdx);

        while (beginIdx != -1)
        {
            logger.trace("    beginIdx={}", beginIdx);

            // Identifier Mode: If there is a "Java Identifier Part" just before
            // or just after the target, then don't replace it, because we've
            // found part of a larger identifier that is not equal to the target.
            // If the identifier is not found in expression delimiters, then don't
            // replace it, because it's literal text not to be modified.
            int exprBeginIdx = value.substring(0, beginIdx).lastIndexOf(Expression.BEGIN_EXPR);
            int exprEndIdx = value.indexOf(Expression.END_EXPR, exprBeginIdx + 1);
            if (identifierMode &&
                    ((exprBeginIdx == -1 || exprBeginIdx > beginIdx || exprEndIdx == -1 || exprEndIdx < beginIdx + target.length()) ||
                            (beginIdx > 0 && Character.isJavaIdentifierPart(value.charAt(beginIdx - 1))) ||
                            (beginIdx + target.length() < value.length() && Character.isJavaIdentifierPart(value.charAt(beginIdx + target.length())))
                    )
                    )
            {
                // Still setup for next loop.
                beginIdx = value.indexOf(target, beginIdx + target.length());

                continue;
            }

            // Take care to skip any already processed part of the value string.
            value = value.substring(0, beginIdx) + replacement + value.substring(beginIdx + target.length());

            updateFormattingRuns(formattingRuns, beginIdx, change);

            if (firstOnly)
                break;
            // Setup for next loop.
            // Look for the next occurrence of the target, taking care to skip the
            // replacement string.  That avoids an infinite loop if the target
            // string can be found inside the replacement string.
            beginIdx = value.indexOf(target, beginIdx + replacement.length());
        }

        return createFormattedString(numFormattingRuns, helper, value, formattingRuns);
    }

    /**
     * Replaces all strings in the given <code>List</code> of strings to replace
     * with the corresponding replacement string in the given <code>List</code>.
     * Preserves rich text formatting as much as possible.
     * @param richTextString The <code>RichTextString</code> to manipulate.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param targets The <code>List</code> of strings to replace.
     * @param replacements The corresponding <code>List</code> of replacement
     *    strings.
     * @param replaceAll If <code>true</code> replace all occurrences, else only
     *    replace the first occurrence.
     * @return A new <code>RichTextString</code> with replaced values, or the
     *    same <code>RichTextString</code> if <code>replace</code> is
     *    <code>null</code> or empty.
     */
    public static RichTextString replaceValues(RichTextString richTextString,
                                               CreationHelper helper, List<String> targets, List<String> replacements, boolean replaceAll)
    {
        if (targets == null || targets.size() == 0)
            return richTextString;

        int numFormattingRuns = richTextString.numFormattingRuns();
        String value = richTextString.getString();
        logger.trace("replaceValues: \"{}\" ({}): numFormattingRuns={}, replacements: {}",
                value, value.length(), numFormattingRuns, targets.size());

        List<FormattingRun> formattingRuns = determineFormattingRunStats(richTextString);

        // Replace targets with replacements.
        for (int i = 0; i < targets.size(); i++)
        {
            int beginIdx = value.indexOf(targets.get(i));
            if (beginIdx != -1)
            {
                String replaceMe = targets.get(i);
                String replaceWith = replacements.get(i);
                int change = replaceWith.length() - replaceMe.length();
                logger.debug("  Replacing \"{}\" with \"{}\".", replaceMe, replaceWith);
                if (replaceAll)
                {
                    value = value.replaceAll(Expression.NEGATIVE_LOOKBEHIND_BACKSLASH + Pattern.quote(replaceMe),
                            Matcher.quoteReplacement(replaceWith));
                }
                else
                {
                    value = value.replaceFirst(Expression.NEGATIVE_LOOKBEHIND_BACKSLASH + Pattern.quote(replaceMe),
                            Matcher.quoteReplacement(replaceWith));
                }

                updateFormattingRuns(formattingRuns, beginIdx, change);
            }
        }

        // Replace "\${" with "${".
        int beginIdx = value.indexOf("\\" + Expression.BEGIN_EXPR);
        while (beginIdx != -1)
        {
            value = value.replace("\\" + Expression.BEGIN_EXPR, Expression.BEGIN_EXPR);
            updateFormattingRuns(formattingRuns, beginIdx, -1);

            beginIdx = value.indexOf("\\" + Expression.BEGIN_EXPR);
        }

        return createFormattedString(numFormattingRuns, helper, value, formattingRuns);
    }

    /**
     * <p>Update all <code>FormattingRuns</code> affected by a change to a
     * <code>RichTextString</code> at the given index by the given change
     * amount.</p>
     * <p>This code was extracted from duplicated code</p>.
     * @param formattingRuns A <code>List</code> of <code>FormattingRuns</code>.
     * @param beginIdx The 0-based position at which a change occurred.
     * @param change The change amount.  If negative, the string is shrinking.
     *    If positive, the string is growing.
     * @since 0.7.0
     */
    private static void updateFormattingRuns(List<FormattingRun> formattingRuns, int beginIdx, int change)
    {
        int numFormattingRuns = formattingRuns.size();
        // Find the formatting run that applies at beginIdx.
        int fmtIndex = -1;
        for (int j = 0; j < numFormattingRuns; j++)
        {
            FormattingRun run = formattingRuns.get(j);
            int currBeginIdx = run.getBegin();
            int currLength = run.getLength();
            logger.debug("    j={}, currBeginIdx={}, currLength={}, beginIdx={}",
                    j, currBeginIdx, currLength, beginIdx);
            // Don't pick a zero-length run.
            if (beginIdx >= currBeginIdx && beginIdx < currBeginIdx + currLength)
            {
                fmtIndex = j;
                break;
            }
        }
        // Found a run to apply.  The length has changed.
        if (fmtIndex != -1)
        {
            // Change the length of the formatting run.
            FormattingRun run = formattingRuns.get(fmtIndex);
            run.setLength(run.getLength() + change);

            // This affects the beginning positions of all subsequent runs!
            for (int j = fmtIndex + 1; j < numFormattingRuns; j++)
            {
                run = formattingRuns.get(j);
                logger.debug("    updateFormattingRuns: Changing beginning of formatting run {} from {} to {}",
                        j, run.getBegin(), run.getBegin() + change);
                run.setBegin(run.getBegin() + change);
            }

            logger.debug("  updateFormattingRuns: Formatting run length changed ({}): newLength={}",
                    fmtIndex, formattingRuns.get(fmtIndex).getLength());
        }
        else
        {
            // It's possible for there to be no formatting run if it's an
            // HSSFRichTextString and the value to replace is before any
            // formatting.  The first formatting "run" is considered to be the
            // format of the Cell and is not present in the formatting runs.

            // This affects the beginning positions of all runs!
            for (int j = 0; j < numFormattingRuns; j++)
            {
                FormattingRun run = formattingRuns.get(j);
                logger.debug("    updateFormattingRuns: changing beginning of formatting run {} from {} to {}",
                        j, run.getBegin(), run.getBegin() + change);
                run.setBegin(run.getBegin() + change);
            }
        }
    }

    /**
     * Extracts a substring of a <code>RichTextString</code> as another
     * <code>RichTextString</code>.  Preserves the formatting that is in place
     * from the given string.
     * @param richTextString The <code>RichTextString</code> of which to take a
     *    substring.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param beginIndex The beginning index, inclusive.
     * @param endIndex The ending index, exclusive.
     * @return The specified substring as a <code>RichTextString</code>, with
     *    the original formatting from the original string intact.
     * @since 0.2.0
     */
    public static RichTextString substring(RichTextString richTextString,
                                           CreationHelper helper, int beginIndex, int endIndex)
    {
        int numFormattingRuns = richTextString.numFormattingRuns();
        String value = richTextString.getString();
        logger.trace("substring: \"{}\" ({}): numFormattingRuns={}, beginIndex: {}, endIndex: {}",
                value, value.length(), numFormattingRuns, beginIndex, endIndex);

        List<FormattingRun> formattingRuns = determineFormattingRunStats(richTextString);

        // Determine which runs apply in the new substring's range.
        List<FormattingRun> substrFormattingRuns = new ArrayList<>();
        int begin, end;
        for (int i = 0; i < numFormattingRuns; i++)
        {
            FormattingRun run = formattingRuns.get(i);
            begin = run.getBegin();
            end = begin + run.getLength();
            if ((begin < beginIndex && end < beginIndex) ||
                    (begin >= endIndex && end >= endIndex))
            {
                // Not copied to the new substring.
                continue;
            }
            if (begin < beginIndex && end >= beginIndex)
            {
                // Partial cover at beginning.
                begin = beginIndex;
            }
            if (begin < endIndex && end >= endIndex)
            {
                // Partial cover at end.
                end = endIndex;
            }
            substrFormattingRuns.add(new FormattingRun(begin - beginIndex, end - begin, run.getFont()));
        }
        return createFormattedString(substrFormattingRuns.size(), helper, value.substring(beginIndex, endIndex),
                substrFormattingRuns);
    }

    /**
     * Determine formatting run statistics for the given
     * <code>RichTextString</code>.  Adds elements to the arrays.
     * @param richTextString The <code>RichTextString</code>.
     * @return A <code>List</code> of all <code>FormattingRun</code>s found.
     */
    public static List<FormattingRun> determineFormattingRunStats(RichTextString richTextString)
    {
        int numFormattingRuns = richTextString.numFormattingRuns();
        List<FormattingRun> formattingRuns = new ArrayList<>(numFormattingRuns);
        if (richTextString instanceof HSSFRichTextString)
        {
            HSSFRichTextString hssfRichTextString = (HSSFRichTextString) richTextString;
            // Determine formatting run statistics.
            for (int fmtIdx = 0; fmtIdx < numFormattingRuns; fmtIdx++)
            {
                int begin = richTextString.getIndexOfFormattingRun(fmtIdx);
                short fontIndex = (Short) getFontOfFormattingRun(hssfRichTextString, fmtIdx);

                // Determine font formatting run length.
                int length = 0;
                for (int j = begin; j < richTextString.length(); j++)
                {
                    short currFontIndex = (Short) getFontAtIndex(hssfRichTextString, j);
                    logger.trace("    Comparing j={}, currFont={}, font={}", j, currFontIndex, fontIndex);
                    if (currFontIndex == fontIndex)
                        length++;
                    else
                        break;
                }

                logger.trace("  dFRS: HSSF Formatting run found: ({}) begin={}, length={}, font={}",
                        fmtIdx, begin, length, fontIndex);
                formattingRuns.add(new FormattingRun(begin, length, fontIndex));
            }
        }
        else if (richTextString instanceof XSSFRichTextString)
        {
            XSSFRichTextString xssfRichTextString = (XSSFRichTextString) richTextString;
            // Determine formatting run statistics.
            for (int fmtIdx = 0; fmtIdx < numFormattingRuns; fmtIdx++)
            {
                int begin = richTextString.getIndexOfFormattingRun(fmtIdx);
                logger.trace("  fmtIdx: {}", fmtIdx);
                logger.trace("    begin: {}", begin);
                XSSFFont fontIndex = (XSSFFont) getFontOfFormattingRun(xssfRichTextString, fmtIdx);

                // Determine font formatting run length.
                int length = 0;
                for (int j = begin; j < richTextString.length(); j++)
                {
                    XSSFFont currFontIndex = (XSSFFont) getFontAtIndex(xssfRichTextString, j);
                    logger.trace("    Comparing j={}, currFont={}, font={}",
                            j, ((currFontIndex != null) ? currFontIndex.toString() : "(null)"), fontIndex);
                    if ((currFontIndex == null && fontIndex == null) ||
                            (currFontIndex != null && currFontIndex.equals(fontIndex)))
                        length++;
                    else
                        break;
                }

                logger.trace("  dFRS: XSSF Formatting run found: ({}) begin={}, length={}, font={}",
                        fmtIdx, begin, length, fontIndex);
                formattingRuns.add(new FormattingRun(begin, length, fontIndex));
            }
        }
        return formattingRuns;
    }

    /**
     * Construct a <code>RichTextString</code> of the same type as
     * <code>richTextString</code>, format it, and return it.
     * @param numFormattingRuns The number of formatting runs.
     * @param value The new string value of the new <code>RichTextString</code>
     *    to construct.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @param formattingRuns A <code>List</code> of <code>FormattingRuns</code>.
     * @return A new <code>RichTextString</code>, the same type as
     *    <code>richTextString</code>, with <code>value</code> as it contents,
     *    formatted as specified.
     */
    public static RichTextString createFormattedString(int numFormattingRuns,
                                                       CreationHelper helper, String value, List<FormattingRun> formattingRuns)
    {
        // Construct the proper RichTextString.
        RichTextString newString = helper.createRichTextString(value);

        formatString(newString, numFormattingRuns, formattingRuns);
        return newString;
    }

    /**
     * Format a <code>RichTextString</code> that has already been created.
     * @param string A <code>RichTextString</code>.
     * @param numFormattingRuns The number of formatting runs.
     * @param formattingRuns A <code>List</code> of <code>FormattingRuns</code>.
     */
    public static void formatString(RichTextString string, int numFormattingRuns,
                                    List<FormattingRun> formattingRuns)
    {
        // Apply the formatting runs.
        for (int i = 0; i < numFormattingRuns; i++)
        {
            FormattingRun run = formattingRuns.get(i);
            int begin = run.getBegin();
            int end = begin + run.getLength();
            Object font = run.getFont();
            logger.trace("  formatString: Applying format ({}): begin={}, length={}, font={} to string \"{}\".",
                    i, begin, run.getLength(), font, string.getString());
            if (string instanceof HSSFRichTextString)
                string.applyFont(begin, end, (Short) font);
            else if (string instanceof XSSFRichTextString)
            {
                if (font != null)
                    string.applyFont(begin, end, (XSSFFont) font);
            }
            else throw new IllegalArgumentException("Unexpected RichTextString type: " +
                        string.getClass().getName() + ": " + string.getString());
        }
    }

    /**
     * Gets the font index of the specified formatting run in the given
     * <code>RichTextString</code>.
     * @param richTextString The <code>RichTextString</code>.
     * @param fmtIndex The 0-based index of the formatting run.
     * @return The font index.  If HSSF, a <code>short</code>.  If XSSF, an
     *    <code>XSSFFont</code>.
     */
    private static Object getFontOfFormattingRun(RichTextString richTextString, int fmtIndex)
    {
        if (richTextString instanceof HSSFRichTextString)
        {
            return ((HSSFRichTextString) richTextString).getFontOfFormattingRun(fmtIndex);
        }
        else if (richTextString instanceof XSSFRichTextString)
        {
            try
            {
                // Instead of returning null, getFontOfFormattingRun (eventually)
                // throws a NullPointerException.  It extracts a "CTRElt" from an
                // array, and it extracts a "CTRPrElt" from the "CTRElt".  The
                // "CTRprElt" can be null if there is no font at the formatting
                // run.  Then, when creating a "CTFont", it calls a method on the
                // null "CTRPrElt".
                // Return the XSSFFont.
                return ((XSSFRichTextString) richTextString).getFontOfFormattingRun(fmtIndex);
            }
            catch (NullPointerException e)
            {
                logger.debug("    NullPointerException caught calling getFontOfFormattingRun({})!", fmtIndex);
                return null;
            }
        }
        else
            throw new IllegalArgumentException("Unexpected RichTextString type: " +
                    richTextString.getClass().getName() + ": " + richTextString.getString());
    }

    /**
     * Gets the font index of the <code>Font</code> in use at the specified
     * position in the given <code>RichTextString</code>.
     * @param richTextString The <code>RichTextString</code>.
     * @param fmtIndex The 0-based index of the formatting run.
     * @return The font index: If HSSF, a <code>short</code>.  If XSSF, an
     *    <code>XSSFFont</code>.
     */
    public static Object getFontAtIndex(RichTextString richTextString, int fmtIndex)
    {
        if (richTextString instanceof HSSFRichTextString)
        {
            // Returns a short.
            return ((HSSFRichTextString) richTextString).getFontAtIndex(fmtIndex);
        }
        else if (richTextString instanceof XSSFRichTextString)
        {
            try
            {
                // Instead of returning null, getFontAtIndex (eventually) throws a
                // NullPointerException.  It extracts a "CTRElt" from an array, and
                // it extracts a "CTRPrElt" from the "CTRElt".  The "CTRprElt" can
                // be null if there is no font at the formatting run.  Then, when
                // creating a "CTFont", it calls a method on the null "CTRPrElt".
                // Return an XSSFFont.
                return ((XSSFRichTextString) richTextString).getFontAtIndex(fmtIndex);
            }
            catch (NullPointerException e)
            {
                logger.debug("    NullPointerException caught calling getFontOfFormattingRun({}).", fmtIndex);
                return null;
            }
        }
        else
            throw new IllegalArgumentException("Unexpected RichTextString type: " +
                    richTextString.getClass().getName() + ": " + richTextString.getString());
    }

    /**
     * Take the first <code>Font</code> from the given
     * <code>RichTextString</code> and apply it to the given <code>Cell's</code>
     * <code>CellStyle</code>.
     * @param cell The <code>Cell</code>.
     * @param richTextString The <code>RichTextString</code> that contains the
     *    desired <code>Font</code>.
     * @param cellStyleCache The <code>CellStyleCache</code> in which to look
     *    for another <code>CellStyle</code>.
     * @param fontCache The <code>FontCache</code> in which to look for another
     *    <code>Font</code>.
     * @since 0.2.0
     */
    public static void applyFont(RichTextString richTextString, Cell cell,
                                 CellStyleCache cellStyleCache, FontCache fontCache)
    {
        logger.trace("applyFont: richTextString = {}, sheet {}, cell at row {}, col {}",
                richTextString, cell.getSheet().getSheetName(), cell.getRowIndex(), cell.getColumnIndex());
        Font font;
        short fontIdx;
        if (richTextString == null)
            return;
        if (richTextString instanceof HSSFRichTextString)
        {
            fontIdx = ((HSSFRichTextString) richTextString).getFontAtIndex(0);
            font = cell.getSheet().getWorkbook().getFontAt(fontIdx);
        }
        else if (richTextString instanceof XSSFRichTextString)
        {
            try
            {
                // Instead of returning null, getFontAtIndex (eventually) throws a
                // NullPointerException.  It extracts a "CTRElt" from an array, and
                // it extracts a "CTRPrElt" from the "CTRElt".  The "CTRprElt" can
                // be null if there is no font at the formatting run.  Then, when
                // creating a "CTFont", it calls a method on the null "CTRPrElt".
                font = ((XSSFRichTextString) richTextString).getFontAtIndex(0);
            }
            catch (NullPointerException e)
            {
                logger.debug("    NullPointerException caught calling getFontAtIndex(0)!");
                font = null;
            }
        }
        else
        {
            throw new IllegalArgumentException("Unexpected RichTextString type: " +
                    richTextString.getClass().getName() + ": " + richTextString.getString());
        }
        if (font != null)
        {
            logger.trace("  Font is {}, index {}", font.toString(), font.getIndex());
            CellStyle cellStyle = cell.getCellStyle();
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle newCellStyle = cellStyleCache.findCellStyleWithFont(cellStyle, font);
            if (newCellStyle == null)
            {
                newCellStyle = workbook.createCellStyle();
                newCellStyle.cloneStyleFrom(cellStyle);
                // For some reason, just setting the Font directly doesn't work.
                //newCellStyle.setFont(font);
                Font foundFont = fontCache.findFont(font);
                newCellStyle.setFont(foundFont);
                cellStyleCache.cacheCellStyle(newCellStyle);
            }
            cell.setCellStyle(newCellStyle);
        }
    }

    /**
     * <p>Performs escaping.  Preserves rich text formatting as much as
     * possible.  The following escape sequences are recognized:</p>
     * <ul>
     *   <li><code>\"</code> =&gt; <code>"</code></li>
     *   <li><code>\'</code> =&gt; <code>'</code></li>
     *   <li><code>\\</code> =&gt; <code>\</code></li>
     *   <li><code>\b</code> =&gt; <code>(backspace)</code></li>
     *   <li><code>\f</code> =&gt; <code>(form feed)</code></li>
     *   <li><code>\n</code> =&gt; <code>(newline)</code></li>
     *   <li><code>\r</code> =&gt; <code>(carriage return)</code></li>
     *   <li><code>\t</code> =&gt; <code>(tab)</code></li>
     * </ul>
     * @param richTextString The <code>RichTextString</code> to manipulate.
     * @param helper A <code>CreationHelper</code> that can create the proper
     *    <code>RichTextString</code>.
     * @return A new <code>RichTextString</code> with escape sequences replaced
     *    with the actual characters.
     * @since 0.7.0
     */
    public static RichTextString performEscaping(RichTextString richTextString,
                                                 CreationHelper helper)
    {
        int numFormattingRuns = richTextString.numFormattingRuns();
        String value = richTextString.getString();
        logger.trace("performEscaping: \"{}\" ({}): numFormattingRuns={}",
                value, value.length(), numFormattingRuns);

        List<FormattingRun> formattingRuns = determineFormattingRunStats(richTextString);
        StringBuilder buf = new StringBuilder();

        for (int i = 0; i < value.length(); i++)
        {
            char curr = value.charAt(i);
            // Beginning of escape sequence
            if (curr == '\\' &&  // Backslash to start escape sequence found
                    (i + 1) < value.length() &&  // Not at end of string
                    "\\\"\'bfnrt".indexOf(value.charAt(i + 1)) != -1)   // Valid 2nd character of escape sequence
            {
                switch (value.charAt(i + 1))
                {
                case '\\':
                    buf.append('\\');
                    break;
                case '\"':
                    buf.append('\"');
                    break;
                case '\'':
                    buf.append('\'');
                    break;
                case 'b':
                    buf.append('\b');
                    break;
                case 'f':
                    buf.append('\f');
                    break;
                case 'n':
                    buf.append('\n');
                    break;
                case 'r':
                    buf.append('\r');
                    break;
                case 't':
                    buf.append('\t');
                    break;
                default:
                    throw new IllegalStateException("Accidentally recognized invalid escape sequence: \"\\" +
                            value.charAt(i + 1) + "\"!");
                }

                updateFormattingRuns(formattingRuns, i, -1);
                // Bypass the second character!
                i++;
            }
            else
            {
                buf.append(curr);
            }
        }

        return createFormattedString(numFormattingRuns, helper, buf.toString(), formattingRuns);
    }
}

/**
 * A <code>FormattingRun</code> holds information about one "run" of a
 * <code>RichTextString</code> that is of the same <code>Font</code>.
 */
class FormattingRun
{
    private int myBeginIdx;
    private int myLength;
    private Object myFont;

    /**
     * Construct a <code>FormattingRun</code> with
     * @param beginIdx The beginning 0-based index of the run.
     * @param length The length of the run.
     * @param font The font.  It should be a <code>Short</code> index for
     *    <code>HSSF</code> or an <code>XSSFFont</code> for <code>XSSF</code>.
     */
    public FormattingRun(int beginIdx, int length, Object font) {
        myBeginIdx = beginIdx;
        myLength = length;
        myFont = font;
    }

    /**
     * Returns the beginning 0-based index of the run.
     * @return The beginning 0-based index of the run.
     */
    public int getBegin()
    {
        return myBeginIdx;
    }

    /**
     * Sets the beginning 0-based index of the run.
     * @param begin The beginning 0-based index of the run.
     */
    public void setBegin(int begin)
    {
        myBeginIdx = begin;
    }

    /**
     * Returns the length of the run.
     * @return The length of the run.
     */
    public int getLength()
    {
        return myLength;
    }

    /**
     * Sets the new formatting run length.
     * @param length The new formatting run length.
     */
    public void setLength(int length)
    {
        myLength = length;
    }

    /**
     * Returns the font.
     * @return A <code>Short</code> index for <code>HSSF</code> or an
     *    <code>XSSFFont</code> for <code>XSSF</code>.
     */
    public Object getFont()
    {
        return myFont;
    }
}